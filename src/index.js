import puppeteer from "@cloudflare/puppeteer";
import { Zip, ZipPassThrough } from "fflate";
import { Buffer } from "buffer";

const CORS = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, X-API-Key",
    "Access-Control-Expose-Headers": "Content-Disposition, Content-Type, X-Output-Url, X-Output-Id",
    "Access-Control-Max-Age": "86400",
};

function json(obj, status = 200) {
    return new Response(JSON.stringify(obj), {
        status,
        headers: { ...CORS, "Content-Type": "application/json; charset=utf-8" },
    });
}

function errJson(message, status = 400, extra = {}) {
    return json({ ok: false, error: message, ...extra }, status);
}

function requireKey(req, env) {
    if (!env.API_KEY) return null;
    const key = req.headers.get("x-api-key") || "";
    return key === env.API_KEY ? null : "Unauthorized";
}

async function appsPost(env, payload) {
    if (!env.APPS_SCRIPT_URL) throw new Error("Missing env.APPS_SCRIPT_URL");

    const r = await fetch(env.APPS_SCRIPT_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
        redirect: "follow",
    });

    const ct = (r.headers.get("content-type") || "").toLowerCase();
    const text = await r.text();
    const trimmed = (text || "").trim();

    // Apps Script deploy sai/quyền sai hay trả HTML / login
    const looksHtml =
        ct.includes("text/html") ||
        trimmed.startsWith("<!doctype") ||
        trimmed.startsWith("<html") ||
        trimmed.startsWith("<head") ||
        trimmed.startsWith("<");

    if (looksHtml) {
        throw new Error(
            "Apps Script trả HTML (thường do deploy chưa để Anyone hoặc dùng sai /exec). Snippet: " +
            trimmed.slice(0, 200)
        );
    }

    let data;
    try {
        data = JSON.parse(trimmed);
    } catch {
        throw new Error("Apps Script không trả JSON. Snippet: " + trimmed.slice(0, 200));
    }

    if (!data.ok) throw new Error(data.error || "Apps Script error");
    return data;
}

function streamZipResponse(filename, outputUrl, outputId) {
    const { readable, writable } = new TransformStream();
    const writer = writable.getWriter();

    const zip = new Zip((err, chunk, final) => {
        if (err) {
            try {
                writer.abort(err);
            } catch { }
            return;
        }
        // chunk là Uint8Array
        writer.write(chunk);
        if (final) writer.close();
    });

    const headers = {
        ...CORS,
        "Content-Type": "application/zip",
        "Content-Disposition": `attachment; filename="${filename}"`,
        "Cache-Control": "no-store",
    };
    if (outputUrl) headers["X-Output-Url"] = outputUrl;
    if (outputId) headers["X-Output-Id"] = outputId;

    return { zip, response: new Response(readable, { status: 200, headers }) };
}

async function renderPng(browser, html, viewport, waitMs) {
    const page = await browser.newPage();
    await page.setViewport(viewport);
    await page.setContent(html, { waitUntil: "networkidle0" });
    try {
        await page.evaluate(() => document.fonts?.ready ?? Promise.resolve());
    } catch { }
    if (waitMs > 0) await new Promise((r) => setTimeout(r, waitMs));
    try {
        const png = await page.screenshot({ type: "png", fullPage: true });
        // puppeteer cf thường trả Uint8Array
        return png instanceof Uint8Array ? png : new Uint8Array(png);
    } finally {
        await page.close();
    }
}

export default {
    async fetch(request, env, ctx) {
        try {
            if (request.method === "OPTIONS") return new Response(null, { status: 204, headers: CORS });

            const url = new URL(request.url);

            // Health / quick check
            if (request.method === "GET") {
                return json({ ok: true, endpoints: ["/api/prepare", "/api/render", "/api/run"] });
            }

            const authErr = requireKey(request, env);
            if (authErr) return errJson(authErr, 401);

            if (request.method !== "POST") return errJson("Method Not Allowed", 405);

            // =============== /api/prepare ===============
            // Upload xlsx -> Apps Script prepare -> trả JSON (output_url, headers, ...)
            if (url.pathname === "/api/prepare") {
                const form = await request.formData();
                const file = form.get("file");
                if (!file || typeof file === "string") return errJson("Thiếu file");

                const headerRow = Number(form.get("headerRow") || 6);
                const sheetPrefix = String(form.get("sheetPrefix") || "CT");
                const outputName = String(form.get("outputName") || "output_tong_hop");

                const ab = await file.arrayBuffer();
                // ✅ FIX: không dùng spread ... (tràn stack)
                const fileBase64 = Buffer.from(ab).toString("base64");
                const fileName = file.name || "upload.xlsx";

                const data = await appsPost(env, {
                    action: "prepare",
                    fileBase64,
                    fileName,
                    headerRow,
                    sheetPrefix,
                    outputName,
                });

                return json(data);
            }

            // =============== /api/render ===============
            // input: { outputSpreadsheetId, selectedHeaders?, cfg?, zipName? }
            // output: ZIP PNG (stream)
            if (url.pathname === "/api/render") {
                const body = await request.json();
                const outputSpreadsheetId = String(body.outputSpreadsheetId || "").trim();
                if (!outputSpreadsheetId) return errJson("Thiếu outputSpreadsheetId");

                const selectedHeaders = Array.isArray(body.selectedHeaders) ? body.selectedHeaders : null;
                const cfg = body.cfg || {};

                const viewport = {
                    width: Number(cfg.minWidth || 980),
                    height: 900,
                    deviceScaleFactor: Number(cfg.deviceScaleFactor || 1.8),
                };
                const waitMs = Number(cfg.waitMs || 60);

                const zipName = String(body.zipName || "bang_con_png.zip");
                const { zip, response } = streamZipResponse(zipName, body.output_url || "", outputSpreadsheetId);

                ctx.waitUntil(
                    (async () => {
                        const browser = await puppeteer.launch(env.MYBROWSER);
                        try {
                            let offset = 0;
                            const limit = 10;

                            while (true) {
                                const pageRes = await appsPost(env, {
                                    action: "buildPages",
                                    outputSpreadsheetId,
                                    selectedHeaders,
                                    cfg,
                                    offset,
                                    limit,
                                });

                                for (const p of pageRes.pages || []) {
                                    const name = p.name || `page_${offset}.png`;
                                    const entry = new ZipPassThrough(name);
                                    zip.add(entry);

                                    const pngBytes = await renderPng(browser, p.html, viewport, waitMs);
                                    entry.push(pngBytes, true);
                                }

                                if (pageRes.done) break;
                                offset = pageRes.nextOffset;
                            }
                        } catch (e) {
                            // nếu lỗi, vẫn end zip để client không treo
                            console.log("render error:", e?.message || String(e));
                        } finally {
                            try {
                                await browser.close();
                            } catch { }
                            zip.end();
                        }
                    })()
                );

                return response;
            }

            // =============== /api/run ===============
            // 1 phát ăn ngay: upload xlsx -> prepare -> buildPages -> render -> zip
            if (url.pathname === "/api/run") {
                const form = await request.formData();
                const file = form.get("file");
                if (!file || typeof file === "string") return errJson("Thiếu file");

                const headerRow = Number(form.get("headerRow") || 6);
                const sheetPrefix = String(form.get("sheetPrefix") || "CT");
                const outputName = String(form.get("outputName") || "output_tong_hop");

                const cfg = form.get("cfgJson") ? JSON.parse(String(form.get("cfgJson"))) : {};
                const selectedHeaders = form.get("selectedHeadersJson")
                    ? JSON.parse(String(form.get("selectedHeadersJson")))
                    : null;

                const ab = await file.arrayBuffer();
                // ✅ FIX: không dùng spread ... (tràn stack)
                const fileBase64 = Buffer.from(ab).toString("base64");
                const fileName = file.name || "upload.xlsx";

                const prep = await appsPost(env, {
                    action: "prepare",
                    fileBase64,
                    fileName,
                    headerRow,
                    sheetPrefix,
                    outputName,
                });

                const outputSpreadsheetId = prep.outputSpreadsheetId;
                const output_url = prep.output_url;

                const viewport = {
                    width: Number(cfg.minWidth || 980),
                    height: 900,
                    deviceScaleFactor: Number(cfg.deviceScaleFactor || 1.8),
                };
                const waitMs = Number(cfg.waitMs || 60);

                const zipName = "bang_con_png.zip";
                const { zip, response } = streamZipResponse(zipName, output_url, outputSpreadsheetId);

                ctx.waitUntil(
                    (async () => {
                        const browser = await puppeteer.launch(env.MYBROWSER);
                        try {
                            let offset = 0;
                            const limit = 10;

                            while (true) {
                                const pageRes = await appsPost(env, {
                                    action: "buildPages",
                                    outputSpreadsheetId,
                                    selectedHeaders: selectedHeaders || prep.headers,
                                    cfg,
                                    offset,
                                    limit,
                                });

                                for (const p of pageRes.pages || []) {
                                    const name = p.name || `page_${offset}.png`;
                                    const entry = new ZipPassThrough(name);
                                    zip.add(entry);

                                    const pngBytes = await renderPng(browser, p.html, viewport, waitMs);
                                    entry.push(pngBytes, true);
                                }

                                if (pageRes.done) break;
                                offset = pageRes.nextOffset;
                            }
                        } catch (e) {
                            console.log("run error:", e?.message || String(e));
                        } finally {
                            try {
                                await browser.close();
                            } catch { }
                            zip.end();
                        }
                    })()
                );

                return response;
            }

            return errJson("Not Found", 404);
        } catch (e) {
            return errJson(e?.message || String(e), 500);
        }
    },
};
// ok
