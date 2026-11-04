const GITHUB_OWNER = 'thering999';
const GITHUB_REPO = 'mchmuk';
const GITHUB_BRANCH = 'main';
const ALLOWED_ORIGIN = 'https://thering999.github.io';

const GH_API = `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/contents`;

function corsHeaders(origin) {
    const allowed = origin === ALLOWED_ORIGIN || origin?.startsWith('http://127.0.0.1') || origin?.startsWith('http://localhost');
    return {
        'Access-Control-Allow-Origin': allowed ? origin : ALLOWED_ORIGIN,
        'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type',
    };
}

export default {
    async fetch(request, env) {
        const origin = request.headers.get('Origin') || '';
        const cors = corsHeaders(origin);

        if (request.method === 'OPTIONS') {
            return new Response(null, { status: 204, headers: cors });
        }

        const url = new URL(request.url);
        const pat = env.GITHUB_PAT;
        if (!pat) return new Response(JSON.stringify({ error: 'GITHUB_PAT not configured' }), { status: 500, headers: { ...cors, 'Content-Type': 'application/json' } });

        const ghHeaders = {
            'Authorization': `Bearer ${pat}`,
            'Accept': 'application/vnd.github+json',
            'X-GitHub-Api-Version': '2022-11-28',
            'Content-Type': 'application/json',
            'User-Agent': 'mchmuk-worker/1.0',
        };

        // GET /sha?path=<filepath> — get current SHA
        if (request.method === 'GET' && url.pathname === '/sha') {
            const filePath = url.searchParams.get('path') || 'tmp_exchange_data.xlsx';
            const res = await fetch(`${GH_API}/${filePath}`, { headers: ghHeaders });
            if (res.status === 404) return new Response(JSON.stringify({ sha: null }), { headers: { ...cors, 'Content-Type': 'application/json' } });
            if (!res.ok) return new Response(JSON.stringify({ error: await res.text() }), { status: res.status, headers: { ...cors, 'Content-Type': 'application/json' } });
            const json = await res.json();
            return new Response(JSON.stringify({ sha: json.sha }), { headers: { ...cors, 'Content-Type': 'application/json' } });
        }

        // POST /upload — push xlsx + meta to GitHub
        if (request.method === 'POST' && url.pathname === '/upload') {
            const body = await request.json();
            const { content, sha, commitMsg, metaContent, metaSha } = body;
            if (!content) return new Response(JSON.stringify({ error: 'missing content' }), { status: 400, headers: { ...cors, 'Content-Type': 'application/json' } });

            // Push xlsx
            const xlsxBody = { message: commitMsg || '📊 อัปเดตข้อมูล HDC', content, branch: GITHUB_BRANCH };
            if (sha) xlsxBody.sha = sha;
            const putRes = await fetch(`${GH_API}/tmp_exchange_data.xlsx`, {
                method: 'PUT', headers: ghHeaders, body: JSON.stringify(xlsxBody),
            });

            if (!putRes.ok) {
                const err = await putRes.json().catch(() => ({}));
                return new Response(JSON.stringify({ ok: false, status: putRes.status, message: err.message || putRes.statusText }), {
                    status: putRes.status, headers: { ...cors, 'Content-Type': 'application/json' },
                });
            }

            const putJson = await putRes.json();
            const commitSha = putJson?.commit?.sha?.slice(0, 7) || '';

            // Push metadata JSON — always fetch fresh SHA inside Worker to avoid stale-SHA 409
            let metaWarning = null;
            if (metaContent) {
                let freshMetaSha = null;
                try {
                    const metaGetRes = await fetch(`${GH_API}/tmp_exchange_data_meta.json`, { headers: ghHeaders });
                    if (metaGetRes.ok) {
                        const metaGetJson = await metaGetRes.json();
                        freshMetaSha = metaGetJson.sha || null;
                    }
                } catch (_) {}

                const metaBody = { message: commitMsg || '📊 อัปเดตข้อมูล HDC', content: metaContent, branch: GITHUB_BRANCH };
                if (freshMetaSha) metaBody.sha = freshMetaSha;

                const metaPutRes = await fetch(`${GH_API}/tmp_exchange_data_meta.json`, {
                    method: 'PUT', headers: ghHeaders, body: JSON.stringify(metaBody),
                });
                if (!metaPutRes.ok) {
                    const metaErr = await metaPutRes.json().catch(() => ({}));
                    metaWarning = `meta push failed ${metaPutRes.status}: ${metaErr.message || metaPutRes.statusText}`;
                }
            }

            return new Response(JSON.stringify({ ok: true, commitSha, ...(metaWarning ? { metaWarning } : {}) }), { headers: { ...cors, 'Content-Type': 'application/json' } });
        }

        return new Response('Not Found', { status: 404, headers: cors });
    },
};
