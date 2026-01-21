import puppeteer from "@cloudflare/puppeteer";
import { Zip, ZipPassThrough } from "fflate";

const corsHeaders = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, OPTIONS, GET",
    "Access-Control-Allow-Headers": "Content-Type, X-API-Key",
    "Access-Control-Expose-Headers": "Content-Disposition, Content-Type, X-Output-Url, X-Output-Id",
    "Access-Control-Max-Age": "86400",
};

function json(obj, status = 200) {
    return new Response(JSON.stringify(obj), {
        status,
        headers: { ...corsHeaders, "Content-Type": "application/json; charset=utf-8" },
    });
}

function bad(msg, status = 400) {
    return json({ ok: false, error: msg }, status);
}

function requireKey(req, env) {
    if (!env.API_KEY) return null;
    const key = req.headers.get("x-api-key") || "";
    if (key !== env.API_KEY) return "Unauthorized";
    return null;
}

async function appsPost(env, payload) {
    const r = await fetch(env.APPS_SCRIPT_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
        redirect: "follow",
    });
    const text = await r.text();
    let data = null;
    try { data = JSON.parse(text); } catch { }
    if (!data) throw new Error("Apps Script không trả JSON: " + text.slice(0, 200));
    if (!data.ok) throw new Error(data.error || "Apps Script error");
    return data;
}

async function renderPng(browser, html, viewport, waitMs) {
    const page = await browser.newPage();
    await page.setViewport(viewport);
    await page.setContent(html, { waitUntil: "networkidle0" });
    try { await page.evaluate(() => document.fonts?.ready ?? Promise.resolve()); } catch { }
    if (waitMs > 0) await new Promise(r => setTimeout(r, waitMs));
    try {
        return await page.screenshot({ type: "png", fullPage: true });
    } finally {
        await page.close();
    }
}

function streamZipResponse(filename, outputUrl, outputId) {
    const { readable, writable } = new TransformStream();
    const writer = writable.getWriter();

    const zip = new Zip((err, data, final) => {
        if (err) { writer.abort(err); return; }
        writer.write(data);
        if (final) writer.close();
    });

    const headers = {
        ...corsHeaders,
        "Content-Type": "application/zip",
        "Content-Disposition": `attachment; filename="${filename}"`,
        "Cache-Control": "no-store",
    };
    if (outputUrl) headers["X-Output-Url"] = outputUrl;
    if (outputId) headers["X-Output-Id"] = outputId;

    return { zip, response: new Response(readable, { status: 200, headers }) };
}

export default {
    async fetch(request, env, ctx) {
        try {
            if (request.method === "OPTIONS") return new Response(null, { status: 204, headers: corsHeaders });

            const url = new URL(request.url);

            if (request.method === "GET") {
                return json({ ok: true, endpoints: ["/api/prepare", "/api/render", "/api/run"] });
            }

            const authErr = requireKey(request, env);
            if (authErr) return bad(authErr, 401);
            if (request.method !== "POST") return bad("Method Not Allowed", 405);

            // 1) /api/prepare: upload xlsx -> Apps Script prepare -> trả JSON
            if (url.pathname === "/api/prepare") {
                const form = await request.formData();
                const file = form.get("file");
                if (!file || typeof file === "string") return bad("Thiếu file");

                const headerRow = Number(form.get("headerRow") || 6);
                const sheetPrefix = String(form.get("sheetPrefix") || "CT");
                const outputName = String(form.get("outputName") || "output_tong_hop");

                const ab = await file.arrayBuffer();
                const fileBase64 = btoa(String.fromCharCode(...new Uint8Array(ab)));
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

            // 2) /api/render: lấy outputSpreadsheetId -> buildPages -> render -> zip
            if (url.pathname === "/api/render") {
                const body = await request.json();
                const outputSpreadsheetId = String(body.outputSpreadsheetId || "");
                if (!outputSpreadsheetId) return bad("Thiếu outputSpreadsheetId");

                const selectedHeaders = Array.isArray(body.selectedHeaders) ? body.selectedHeaders : null;
                const cfg = body.cfg || {};

                const viewport = {
                    width: Number(cfg.minWidth || 980),
                    height: 900,
                    deviceScaleFactor: Number(cfg.deviceScaleFactor || 1.8),
                };
                const waitMs = Number(cfg.waitMs || 60);

                const zipName = String(body.zipName || "bang_cong_png.zip");
                const { zip, response } = streamZipResponse(zipName, body.output_url || "", outputSpreadsheetId);

                ctx.waitUntil((async () => {
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

                            for (const p of (pageRes.pages || [])) {
                                const entry = new ZipPassThrough(p.name || `page_${offset}.png`);
                                zip.add(entry);

                                const png = await renderPng(browser, p.html, viewport, waitMs);
                                const bytes = (png instanceof ArrayBuffer) ? new Uint8Array(png) : new Uint8Array(png);
                                entry.push(bytes, true);
                            }

                            if (pageRes.done) break;
                            offset = pageRes.nextOffset;
                        }
                    } finally {
                        try { await browser.close(); } catch { }
                        zip.end();
                    }
                })());

                return response;
            }

            // 3) /api/run: 1 phát ăn ngay -> trả zip
            if (url.pathname === "/api/run") {
                const form = await request.formData();
                const file = form.get("file");
                if (!file || typeof file === "string") return bad("Thiếu file");

                const headerRow = Number(form.get("headerRow") || 6);
                const sheetPrefix = String(form.get("sheetPrefix") || "CT");
                const outputName = String(form.get("outputName") || "output_tong_hop");

                const cfg = form.get("cfgJson") ? JSON.parse(String(form.get("cfgJson"))) : {};
                const selectedHeaders = form.get("selectedHeadersJson")
                    ? JSON.parse(String(form.get("selectedHeadersJson")))
                    : null;

                const ab = await file.arrayBuffer();
                const fileBase64 = btoa(String.fromCharCode(...new Uint8Array(ab)));
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

                const zipName = "bang_cong_png.zip";
                const { zip, response } = streamZipResponse(zipName, output_url, outputSpreadsheetId);

                ctx.waitUntil((async () => {
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

                            for (const p of (pageRes.pages || [])) {
                                const entry = new ZipPassThrough(p.name || `page_${offset}.png`);
                                zip.add(entry);
                                const png = await renderPng(browser, p.html, viewport, waitMs);
                                const bytes = (png instanceof ArrayBuffer) ? new Uint8Array(png) : new Uint8Array(png);
                                entry.push(bytes, true);
                            }

                            if (pageRes.done) break;
                            offset = pageRes.nextOffset;
                        }
                    } finally {
                        try { await browser.close(); } catch { }
                        zip.end();
                    }
                })());

                return response;
            }

            return bad("Not Found", 404);
        } catch (err) {
            return bad(err?.message || String(err), 500);
        }
    },
};
