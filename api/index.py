"""
api/index.py — Vercel entrypoint for BRIMI Flask app.
All HTML inlined — no templates folder (Vercel only deploys api/ files).
"""
import json
import os
import sys
import uuid
import requests
from datetime import datetime
from flask import Flask, request, jsonify, send_file, abort

# Add project root to sys.path for pipeline modules
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024

# Use /tmp for uploads — /var/task is read-only on Vercel
TMP_DIR = "/tmp"

# ─── Inlined HTML (replaces templates/index.html) ──────────────────────────
def serve_page(today):
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>BRIMI BRIMI! HAHAHA!</title>
    <link href="https://fonts.googleapis.com/css2?family=Google+Sans:wght@400;500;700&family=Roboto:wght@400;500&display=swap" rel="stylesheet">
    <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
    <style>
        :root {{
            --md-sys-color-primary: #FFB4C2;
            --md-sys-color-primary-dark: #FF8FA3;
            --md-sys-color-on-primary: #000000;
            --md-sys-color-surface: #FFFFFF;
            --md-sys-color-surface-variant: #F5F5F5;
            --md-sys-color-outline: #E0E0E0;
            --md-sys-color-on-surface: #1C1B1F;
            --md-sys-color-success: #4CAF50;
            --md-sys-color-error: #F44336;
            --md-ref-typeface-brand: 'Google Sans', sans-serif;
            --md-ref-typeface-plain: 'Roboto', sans-serif;
        }}
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: var(--md-ref-typeface-plain);
            background: var(--md-sys-color-surface);
            color: var(--md-sys-color-on-surface);
            min-height: 100vh;
            display: flex;
            flex-direction: column;
            align-items: center;
        }}
        .container {{
            max-width: 640px;
            width: 100%;
            padding: 24px 16px;
        }}
        header {{ text-align: center; margin-bottom: 32px; }}
        header h1 {{
            font-family: var(--md-ref-typeface-brand);
            font-size: 28px;
            font-weight: 700;
            margin-bottom: 4px;
        }}
        header .subtitle {{ font-size: 14px; color: #666; }}
        header .date {{ font-size: 13px; color: #888; margin-top: 8px; font-weight: 500; }}
        header .admin-link {{
            position: absolute;
            top: 24px;
            right: 24px;
            font-size: 13px;
            color: #999;
            text-decoration: none;
        }}
        header .admin-link:hover {{ color: var(--md-sys-color-primary-dark); }}
        .card {{
            background: var(--md-sys-color-surface);
            border: 1px solid var(--md-sys-color-outline);
            border-radius: 12px;
            padding: 24px;
            margin-bottom: 24px;
        }}
        .upload-field {{ margin-bottom: 20px; }}
        .upload-field label {{ display: block; font-weight: 500; font-size: 14px; margin-bottom: 8px; }}
        .upload-field .hint {{ font-size: 12px; color: #888; margin-bottom: 6px; }}
        .upload-field input[type="file"] {{
            width: 100%;
            padding: 12px;
            border: 2px dashed var(--md-sys-color-outline);
            border-radius: 8px;
            background: var(--md-sys-color-surface-variant);
            cursor: pointer;
            font-size: 14px;
        }}
        .upload-field input[type="file"]:hover {{ border-color: var(--md-sys-color-primary); }}
        .process-btn {{
            width: 100%;
            padding: 14px 24px;
            background: var(--md-sys-color-primary);
            color: var(--md-sys-color-on-primary);
            border: none;
            border-radius: 8px;
            font-family: var(--md-ref-typeface-brand);
            font-size: 16px;
            font-weight: 500;
            cursor: pointer;
        }}
        .process-btn:hover {{ background: var(--md-sys-color-primary-dark); }}
        .process-btn:disabled {{ background: #E0E0E0; color: #999; cursor: not-allowed; }}
        .loading {{ display: none; text-align: center; padding: 24px; }}
        .loading.active {{ display: block; }}
        .spinner {{
            display: inline-block;
            width: 40px;
            height: 40px;
            border: 4px solid var(--md-sys-color-outline);
            border-top-color: var(--md-sys-color-primary);
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
            margin-bottom: 12px;
        }}
        @keyframes spin {{ to {{ transform: rotate(360deg); }} }}
        .loading p {{ font-size: 14px; color: #666; }}
        .status {{ display: none; margin-top: 24px; }}
        .status.active {{ display: block; }}
        .status.success {{ padding: 16px; background: #E8F5E9; border: 1px solid #4CAF50; border-radius: 8px; }}
        .status.error {{ padding: 16px; background: #FFEBEE; border: 1px solid #F44336; border-radius: 8px; }}
        .status .title {{ font-weight: 500; font-size: 15px; margin-bottom: 8px; }}
        .status.success .title {{ color: #2E7D32; }}
        .status.error .title {{ color: #C62828; }}
        .download-btn {{
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 12px 24px;
            background: #4CAF50;
            color: white;
            border: none;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            text-decoration: none;
        }}
        .download-btn:hover {{ background: #43A047; }}
        details {{ margin-top: 12px; border: 1px solid var(--md-sys-color-outline); border-radius: 8px; }}
        details summary {{ padding: 12px 16px; cursor: pointer; font-size: 13px; font-weight: 500; color: #666; }}
        .log-content {{
            padding: 12px 16px;
            font-family: monospace;
            font-size: 12px;
            line-height: 1.6;
            max-height: 300px;
            overflow-y: auto;
            white-space: pre-wrap;
            word-break: break-all;
        }}
        footer {{ text-align: center; padding: 24px 16px; font-size: 12px; color: #999; margin-top: auto; }}
        @media (max-width: 480px) {{
            .container {{ padding: 16px 12px; }}
            header h1 {{ font-size: 24px; }}
            .card {{ padding: 16px; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>BRIMI BRIMI! HAHAHA!</h1>
            <p class="subtitle">MBG: Measuring BRIMI's Growth</p>
            <p class="date">{today}</p>
            <a class="admin-link" href="/admin">Manage Peers</a>
        </header>
        <div class="card">
            <form id="uploadForm" enctype="multipart/form-data">
                <div class="upload-field">
                    <label for="historicalnav_t1">HistoricalNAV (T-1)</label>
                    <p class="hint">BRI daily AUM — most recent date</p>
                    <input type="file" id="historicalnav_t1" name="historicalnav_t1" accept=".xlsx" required>
                </div>
                <div class="upload-field">
                    <label for="historicalnav_t2">HistoricalNAV (T-2)</label>
                    <p class="hint">BRI daily AUM — previous date</p>
                    <input type="file" id="historicalnav_t2" name="historicalnav_t2" accept=".xlsx" required>
                </div>
                <div class="upload-field">
                    <label for="bloomberg">INDEKS Bloomberg</label>
                    <p class="hint">Bloomberg global indexes</p>
                    <input type="file" id="bloomberg" name="bloomberg" accept=".xlsx" required>
                </div>
                <button type="submit" class="process-btn" id="processBtn">Process Files</button>
            </form>
            <div class="loading" id="loading">
                <div class="spinner"></div>
                <p>Processing files... (this may take 5–15 seconds)</p>
            </div>
            <div class="status" id="status">
                <div class="title" id="statusTitle"></div>
                <div id="statusBody"></div>
            </div>
        </div>
    </div>
    <footer>&copy; 2026 Saya akan LAWANNN!. All rights reserved.</footer>
    <script>
        const form = document.getElementById('uploadForm');
        const loading = document.getElementById('loading');
        const status = document.getElementById('status');
        const statusTitle = document.getElementById('statusTitle');
        const statusBody = document.getElementById('statusBody');
        const processBtn = document.getElementById('processBtn');

        form.addEventListener('submit', async (e) => {{
            e.preventDefault();
            const t1 = document.getElementById('historicalnav_t1').files[0];
            const t2 = document.getElementById('historicalnav_t2').files[0];
            const bb = document.getElementById('bloomberg').files[0];
            if (!t1 || !t2 || !bb) {{ alert('Please select all 3 files.'); return; }}

            loading.classList.add('active');
            status.classList.remove('active', 'success', 'error');
            processBtn.disabled = true;
            processBtn.textContent = 'Processing...';

            const formData = new FormData();
            formData.append('historicalnav_t1', t1);
            formData.append('historicalnav_t2', t2);
            formData.append('bloomberg', bb);

            try {{
                const response = await fetch('/process', {{ method: 'POST', body: formData }});
                let data;
                try {{
                    data = await response.json();
                }} catch (jsonErr) {{
                    const text = await response.text().catch(() => '');
                    data = {{
                        status: 'error',
                        message: `Server returned status ${{response.status}} ${{response.statusText}}: ${{text.substring(0, 300)}}`
                    }};
                }}
                handleResponse(data);
            }} catch (err) {{
                loading.classList.remove('active');
                status.classList.add('active', 'error');
                statusTitle.textContent = 'Network error';
                statusBody.textContent = (err.message === 'Failed to fetch') 
                    ? 'Koneksi terputus atau server timeout saat memproses file. Pastikan koneksi internet stabil dan coba klik Process Files sekali lagi.' 
                    : err.message;
            }}
            processBtn.disabled = false;
            processBtn.textContent = 'Process Files';
        }});

        function handleResponse(data) {{
            loading.classList.remove('active');
            status.classList.add('active');
            processBtn.disabled = false;
            processBtn.textContent = 'Process Files';
            if (data.status === 'ok') {{
                status.classList.add('success');
                statusTitle.textContent = 'Success!';
                statusBody.innerHTML = '<a href="' + data.download_url + '" class="download-btn"><span class="material-icons">download</span> Download Output</a><details><summary>Processing Log</summary><div class="log-content">' + escapeHtml(data.logs.join('\\\\n')) + '</div></details>';
            }} else {{
                status.classList.add('error');
                statusTitle.textContent = 'Error: ' + escapeHtml(data.message);
                statusBody.innerHTML = '<details><summary>Full Log</summary><div class="log-content">' + escapeHtml(data.logs.join('\\\\n')) + '</div></details>';
            }}
        }}
        function escapeHtml(s) {{ const d = document.createElement('div'); d.textContent = s; return d.innerHTML; }}
    </script>
</body>
</html>"""


@app.route("/")
def index():
    today = datetime.now().strftime("%d %B %Y")
    return serve_page(today)


@app.route("/process", methods=["POST"])
def process():
    request_id = str(uuid.uuid4())[:8]
    logs = []

    def log_callback(msg):
        logs.append(msg)

    try:
        files = request.files
        missing = []
        for key, label in [
            ("historicalnav_t1", "HistoricalNAV (T-1)"),
            ("historicalnav_t2", "HistoricalNAV (T-2)"),
            ("bloomberg", "INDEKS Bloomberg"),
        ]:
            if key not in files or files[key].filename == "":
                missing.append(label)

        if missing:
            return jsonify({
                "status": "error",
                "message": f"Missing files: {', '.join(missing)}"
            }), 400

        saved_paths = {}
        for key in ["historicalnav_t1", "historicalnav_t2", "bloomberg"]:
            f = files[key]
            ext = os.path.splitext(f.filename)[1] or ".xlsx"
            tmp_path = os.path.join(TMP_DIR, f"upload_{request_id}_{key}{ext}")
            f.save(tmp_path)
            saved_paths[key] = tmp_path

        from brimi_engine import run_pipeline
        output_path, nav_date = run_pipeline(
            saved_paths["historicalnav_t1"],
            saved_paths["historicalnav_t2"],
            saved_paths["bloomberg"],
            log_callback=log_callback,
        )

        download_name = f"BRIMI_Output_{nav_date.replace(' ', '_')}.xlsx"
        download_path = os.path.join(TMP_DIR, f"output_{request_id}_{download_name}")
        import shutil
        shutil.copy2(output_path, download_path)

        for p in saved_paths.values():
            try:
                os.remove(p)
            except OSError:
                pass

        return jsonify({
            "status": "ok",
            "download_url": f"/download?file={os.path.basename(download_path)}",
            "logs": logs,
        })

    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        logs.append(f"\nERROR: {e}")
        logs.append(tb)
        return jsonify({
            "status": "error",
            "message": str(e),
            "logs": logs,
        }), 500


@app.route("/download")
def download():
    filename = request.args.get("file")
    if not filename:
        abort(400)
    filepath = os.path.join(TMP_DIR, filename)
    if not os.path.exists(filepath):
        abort(404)
    return send_file(
        filepath,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


# ─── Admin Page ──────────────────────────────────────────────────────────────

def serve_admin_page():
    return """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>BRIMI Peer Manager</title>
    <link href="https://fonts.googleapis.com/css2?family=Google+Sans:wght@400;500;700&family=Roboto:wght@400;500&display=swap" rel="stylesheet">
    <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
    <style>
        :root {
            --md-sys-color-primary: #FFB4C2;
            --md-sys-color-primary-dark: #FF8FA3;
            --md-sys-color-on-primary: #000000;
            --md-sys-color-surface: #FFFFFF;
            --md-sys-color-surface-variant: #F5F5F5;
            --md-sys-color-outline: #E0E0E0;
            --md-sys-color-on-surface: #1C1B1F;
            --md-ref-typeface-brand: 'Google Sans', sans-serif;
            --md-ref-typeface-plain: 'Roboto', sans-serif;
        }
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: var(--md-ref-typeface-plain); background: var(--md-sys-color-surface); color: var(--md-sys-color-on-surface); min-height: 100vh; }
        .container { max-width: 900px; margin: 0 auto; padding: 24px 16px; }
        header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 24px; }
        header h1 { font-family: var(--md-ref-typeface-brand); font-size: 22px; }
        header a { color: var(--md-sys-color-primary); text-decoration: none; font-size: 14px; }
        .section-card { border: 1px solid var(--md-sys-color-outline); border-radius: 12px; margin-bottom: 12px; overflow: hidden; }
        .section-header { padding: 16px 20px; background: var(--md-sys-color-surface-variant); cursor: pointer; display: flex; justify-content: space-between; align-items: center; font-weight: 500; }
        .section-header:hover { background: #e8e8e8; }
        .section-header .badge { background: var(--md-sys-color-primary); border-radius: 12px; padding: 2px 10px; font-size: 12px; }
        .section-body { display: none; padding: 16px 20px; }
        .section-body.open { display: block; }
        .peer-row { display: flex; align-items: center; gap: 12px; padding: 8px 0; border-bottom: 1px solid var(--md-sys-color-outline); }
        .peer-row:last-child { border-bottom: none; }
        .peer-row input[type="checkbox"] { width: 18px; height: 18px; }
        .peer-row span { flex: 1; font-size: 14px; }
        .peer-row .idx { color: #999; font-size: 12px; min-width: 24px; }
        .peer-row .remove-btn { background: none; border: none; color: #F44336; cursor: pointer; font-size: 18px; padding: 4px; }
        .peer-row .remove-btn:hover { color: #C62828; }
        .peer-row.inactive span { text-decoration: line-through; color: #999; }
        .add-row { display: flex; gap: 8px; margin-top: 12px; }
        .add-row input { flex: 1; padding: 10px; border: 1px solid var(--md-sys-color-outline); border-radius: 8px; font-size: 14px; }
        .add-row button { padding: 10px 16px; background: var(--md-sys-color-primary); border: none; border-radius: 8px; font-family: var(--md-ref-typeface-brand); font-size: 14px; cursor: pointer; }
        .add-row button:hover { background: var(--md-sys-color-primary-dark); }
        .save-btn { width: 100%; padding: 14px; background: #4CAF50; color: white; border: none; border-radius: 8px; font-family: var(--md-ref-typeface-brand); font-size: 16px; cursor: pointer; margin-top: 20px; }
        .save-btn:hover { background: #43A047; }
        .save-btn:disabled { background: #ccc; cursor: not-allowed; }
        .toast { position: fixed; bottom: 24px; left: 50%; transform: translateX(-50%); padding: 12px 24px; background: #333; color: white; border-radius: 8px; font-size: 14px; display: none; }
        .toast.show { display: block; }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>Peer Group Manager</h1>
            <a href="/">&#8592; Back to BRIMI</a>
        </header>
        <div id="sections"></div>
        <button class="save-btn" id="saveBtn" onclick="saveChanges()">Save All Changes</button>
    </div>
    <div class="toast" id="toast"></div>
    <script>
        const toast = (msg) => { const t = document.getElementById('toast'); t.textContent = msg; t.classList.add('show'); setTimeout(() => t.classList.remove('show'), 3000); };
        let config = null;

        async function load() {
            try {
                const r = await fetch('/admin/config');
                if (!r.ok) {
                    const text = await r.text();
                    toast(`Failed to load (${r.status}): ${text.substring(0, 200)}`);
                    return;
                }
                const data = await r.json();
                if (data.status !== 'ok') {
                    toast('Failed to load: ' + (data.message || 'Unknown error'));
                    return;
                }
                config = data;
                render();
            } catch (e) { toast('Failed to load: ' + e.message); }
        }

        function render() {
            const el = document.getElementById('sections');
            el.innerHTML = '';
            config.sections.forEach((sec, si) => {
                const card = document.createElement('div');
                card.className = 'section-card';
                const funds = sec.funds || [];
                const lead = funds.find(f => !f.is_index);
                const idxCount = funds.filter(f => f.is_index).length;

                let headerHtml = `<div class="section-header" onclick="toggle(${si})">
                    <span>${sec.section}${lead ? ' — ' + lead.display_name : ''}</span>
                    <span class="badge">${funds.length}</span>
                </div>`;

                let bodyHtml = `<div class="section-body" id="body${si}">`;
                funds.forEach((f, fi) => {
                    const inactive = !f.active ? ' inactive' : '';
                    const type = f.is_index ? '📊 ' : '';
                    bodyHtml += `<div class="peer-row${inactive}" id="peer-${si}-${fi}">
                        <span class="idx">${fi + 1}</span>
                        <input type="checkbox" ${f.active !== false ? 'checked' : ''} onchange="toggleActive(${si},${fi},this.checked)">
                        <span>${type}${f.display_name}</span>
                        ${!f.is_index ? `<button class="remove-btn" onclick="removePeer(${si},${fi})" title="Remove">✕</button>` : ''}
                    </div>`;
                });

                bodyHtml += `<div class="add-row">
                    <input type="text" id="add-${si}" placeholder="Fund name (exact match with fund universe)..." onkeydown="if(event.key==='Enter')addPeer(${si})">
                    <button onclick="addPeer(${si})">Add</button>
                </div></div>`;

                card.innerHTML = headerHtml + bodyHtml;
                el.appendChild(card);
            });
        }

        function toggle(si) {
            document.getElementById('body' + si).classList.toggle('open');
        }

        function toggleActive(si, fi, checked) {
            config.sections[si].funds[fi].active = checked;
            const row = document.getElementById('peer-' + si + '-' + fi);
            row.classList.toggle('inactive', !checked);
        }

        function removePeer(si, fi) {
            config.sections[si].funds.splice(fi, 1);
            render();
            document.getElementById('body' + si).classList.add('open');
        }

        function addPeer(si) {
            const input = document.getElementById('add-' + si);
            const name = input.value.trim();
            if (!name) return;
            config.sections[si].funds.splice(
                config.sections[si].funds.findIndex(f => f.is_index) !== -1
                    ? config.sections[si].funds.findIndex(f => f.is_index)
                    : config.sections[si].funds.length,
                0,
                { display_name: name, alias: null, is_index: false, active: true }
            );
            input.value = '';
            render();
            document.getElementById('body' + si).classList.add('open');
        }

        async function saveChanges() {
            const btn = document.getElementById('saveBtn');
            btn.disabled = true;
            btn.textContent = 'Saving...';
            try {
                const r = await fetch('/admin/save', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(config)
                });
                const data = await r.json();
                if (data.status === 'ok') {
                    toast('Saved! Committed to git.');
                } else {
                    toast('Error: ' + (data.message || 'Unknown'));
                }
            } catch (e) {
                toast('Save failed: ' + e.message);
            }
            btn.disabled = false;
            btn.textContent = 'Save All Changes';
        }

        load();
    </script>
</body>
</html>"""


@app.route("/admin")
def admin_page():
    return serve_admin_page()


@app.route("/admin/config", methods=["GET"])
def admin_config():
    """Return current page table config (universe + any existing overrides)."""
    try:
        PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        universe_path = os.path.join(PROJECT_ROOT, "fund_universe.json")
        with open(universe_path) as f:
            universe = json.load(f)

        # Apply any existing overrides (matched by section index)
        overrides_path = os.path.join(PROJECT_ROOT, "peer_overrides.json")
        if os.path.exists(overrides_path):
            with open(overrides_path) as f:
                overrides = json.load(f)
            for idx, section in enumerate(universe["page_table"]["sections"]):
                if str(idx) in overrides.get("sections", {}):
                    section["funds"] = overrides["sections"][str(idx)]

        # Build response: one entry per section (including duplicates) with its funds
        result = []
        for section in universe["page_table"]["sections"]:
            result.append({
                "section": section["section"],
                "funds": section["funds"]
            })

        return jsonify({"status": "ok", "sections": result})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/admin/save", methods=["POST"])
def admin_save():
    """Save peer config — only diffs vs fund_universe.json, committed via GitHub API.
    Overrides stored as {section_index: [funds]} dict to handle duplicate section names."""
    try:
        import base64
        data = request.get_json()
        if not data or "sections" not in data:
            return jsonify({"status": "error", "message": "No sections provided"}), 400

        PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        universe_path = os.path.join(PROJECT_ROOT, "fund_universe.json")
        with open(universe_path) as f:
            base_universe = json.load(f)

        base_sections = base_universe["page_table"]["sections"]
        incoming = data["sections"]

        # Compare each section by index (handles duplicate section names correctly)
        # Store overrides as {"0": [funds], "5": [funds]} — dict keyed by section index
        overrides = {}
        for i in range(len(base_sections)):
            if i >= len(incoming):
                break
            sec_in = incoming[i]
            sec_base = base_sections[i]
            if sec_in["section"] != sec_base["section"]:
                continue

            # Normalise for comparison: strip 'active' field
            def norm(funds):
                return [{k: f[k] for k in ("display_name", "alias", "is_index")
                         if k in f} for f in funds]

            if norm(sec_in.get("funds", [])) != norm(sec_base.get("funds", [])):
                # Store just the fund dicts list, keyed by index
                overrides[str(i)] = sec_in["funds"]

        # If no overrides, delete peer_overrides.json from GitHub
        if not overrides:
            github_token = os.environ.get("GITHUB_TOKEN")
            if not github_token:
                return jsonify({"status": "ok", "message": "No overrides — all sections match base"})

            owner = os.environ.get("GITHUB_REPO_OWNER", "cevinkidambi")
            repo = os.environ.get("GITHUB_REPO_NAME", "brimi-webapp")
            branch = os.environ.get("GITHUB_BRANCH", "main")
            path = "peer_overrides.json"
            url = f"https://api.github.com/repos/{owner}/{repo}/contents/{path}"
            headers = {"Authorization": f"Bearer {github_token}", "Accept": "application/vnd.github.v3+json"}

            resp = requests.get(url, headers=headers, params={"ref": branch})
            if resp.status_code == 200:
                sha = resp.json().get("sha")
                del_resp = requests.delete(url, headers=headers, json={
                    "message": "Clear peer overrides — all sections match base",
                    "sha": sha,
                    "branch": branch,
                })
                status_msg = "Overrides cleared — all sections match base" if del_resp.status_code in (200, 201) else f"Delete error: {del_resp.status_code}"
                return jsonify({"status": "ok", "message": status_msg})

            return jsonify({"status": "ok", "message": "No overrides file to clear"})

        content = json.dumps(overrides, indent=2, ensure_ascii=False)

        # Commit to GitHub
        github_token = os.environ.get("GITHUB_TOKEN")
        if not github_token:
            # Fallback: save locally for testing
            overrides_path = os.path.join(PROJECT_ROOT, "peer_overrides.json")
            with open(overrides_path, "w") as f:
                f.write(content)
            return jsonify({"status": "ok", "message": f"Saved locally ({len(overrides)} section(s) changed)"})

        owner = os.environ.get("GITHUB_REPO_OWNER", "cevinkidambi")
        repo = os.environ.get("GITHUB_REPO_NAME", "brimi-webapp")
        branch = os.environ.get("GITHUB_BRANCH", "main")
        path = "peer_overrides.json"

        url = f"https://api.github.com/repos/{owner}/{repo}/contents/{path}"
        headers = {
            "Authorization": f"Bearer {github_token}",
            "Accept": "application/vnd.github.v3+json",
        }

        # Check if file exists (need sha for update)
        resp = requests.get(url, headers=headers, params={"ref": branch})
        sha = None
        if resp.status_code == 200:
            sha = resp.json().get("sha")

        body = {
            "message": f"Update peer overrides ({len(overrides)} section(s) changed)",
            "content": base64.b64encode(content.encode()).decode(),
            "branch": branch,
        }
        if sha:
            body["sha"] = sha

        resp = requests.put(url, headers=headers, json=body)
        if resp.status_code in (200, 201):
            return jsonify({"status": "ok", "message": f"Committed to git ({len(overrides)} section(s))"})
        else:
            return jsonify({
                "status": "error",
                "message": f"GitHub API error: {resp.status_code} {resp.text}"
            }), 500
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500
