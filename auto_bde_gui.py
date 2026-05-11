#!/usr/bin/env python3
"""
BDE Assignment Tool — Interfaccia web
======================================
Si apre nel browser. Seleziona i 3 file Excel, premi Esegui.
Funziona su Mac, Windows, Linux.

Requisiti: pip install pandas openpyxl
"""

import os
import sys
import json
import base64
import threading
import webbrowser
import tempfile
import shutil
from http.server import HTTPServer, BaseHTTPRequestHandler
from datetime import datetime

import pandas as pd
import re

# ===================== LOGICA BDE =====================

BDE_TAGS = {"ndg", "national id", "piva"}
TABLE_KEY_COLS = ["RGP", "CON_ID", "TBL_SCH", "TBL_NM"]


def make_table_key(df):
    return (
        df["RGP"].astype(str) + "|" +
        df["CON_ID"].astype(str) + "|" +
        df["TBL_SCH"].astype(str) + "|" +
        df["TBL_NM"].astype(str)
    )


def classify_pattern(fields, bde_code):
    fields_upper = [f.upper() for f in fields]
    fields_str = " ".join(fields_upper)

    has_new = any(k in f for f in fields_upper for k in ["_NEW", "_NUOV", "NEW_"])
    has_old = any(k in f for f in fields_upper for k in ["_OLD", "_VECCH", "OLD_", "_PREC"])
    if has_new and has_old:
        return "Pattern 3A: NEW/OLD"

    if any(re.search(r"[\d]", f) for f in fields_upper):
        return "Pattern 3B: Numerati"

    role_kw = [
        "TITOL", "CONTRAE", "DELEG", "RICHIED", "BENEFIC", "ESECUTOR",
        "CAPO", "GARANT", "COINTEST", "FIDEIUSS", "ORDINANT", "ASSIC",
        "LOCATOR", "FORNIT", "OPERANTE", "MITTENT", "DESTINAT",
    ]
    if any(k in fields_str for k in role_kw):
        return "Pattern 4: Ruoli"

    pure_keywords = {
        "NDG": ["NDG", "COD_NDG", "CODNDG", "C_NDG"],
        "CF": ["CF", "CFIS", "CF_PIVA", "COD_FISC", "CODICE_FISC", "CFISC",
               "COD_CF", "CFPIVA", "CODFISC", "CODFISCALE", "COD_FISCALE", "CODICE_FISCALE"],
        "PIVA": ["PIVA", "P_IVA", "PARTITA_IVA", "COD_PIVA", "PARTITAIVA"],
    }
    kw = pure_keywords.get(bde_code, [])
    pure = [f for f in fields if any(f.upper().replace(" ", "") == k for k in kw)]
    suffixed = [f for f in fields if f not in pure]
    if len(pure) >= 1 and len(suffixed) >= 1:
        return "Pattern 1: Puro vs Suffisso"

    return "Pattern 2: Pareggio/Altro"


def run_bde_assignment(file_discovery, file_perimeter, file_mapping, output_dir):
    logs = []
    def log(msg):
        logs.append(msg)

    log("Caricamento file...")
    df_main = pd.read_excel(file_discovery)
    df_per = pd.read_excel(file_perimeter)
    df_map = pd.read_excel(file_mapping, sheet_name="MappingBizTerm")

    log(f"  Discovery: {len(df_main)} righe")
    log(f"  Perimetro: {len(df_per)} righe")
    log(f"  Mapping: {len(df_map)} righe")

    for col in TABLE_KEY_COLS + ["FLD_NM"]:
        if not (df_main[col] == df_per[col]).all():
            raise ValueError(f"Disallineamento colonna {col} tra discovery e perimetro")

    tag_to_bde = {
        str(row["PDP_TAG"]).strip().lower(): str(row["BizTerm"]).strip()
        for _, row in df_map.iterrows()
    }

    log("Assegnazione BDE...")
    df = df_main.copy()
    df["_TBL_KEY"] = make_table_key(df)
    df["_PDP_LOWER"] = df["PDP_TAG"].str.strip().str.lower()
    counts = df.groupby(["_TBL_KEY", "_PDP_LOWER"]).size().to_dict()

    bde_col = []
    for _, row in df.iterrows():
        existing = row["BDE"]
        if pd.notna(existing) and str(existing).strip() != "":
            bde_col.append(existing)
            continue
        pdp = row["_PDP_LOWER"]
        if pdp not in BDE_TAGS:
            bde_col.append("")
            continue
        bde_name = tag_to_bde.get(pdp, "")
        if not bde_name:
            bde_col.append("")
            continue
        n = counts.get((row["_TBL_KEY"], pdp), 0)
        bde_col.append(bde_name if n == 1 else f"{bde_name} [DA VERIFICARE]")

    df["BDE"] = bde_col
    df.insert(1, "RGP_PERIMETER", df_per["RGP_PERIMETER"].values)
    df.drop(columns=["_TBL_KEY", "_PDP_LOWER"], inplace=True)

    filled = sum(1 for v in bde_col if v)
    verificare = sum(1 for v in bde_col if "DA VERIFICARE" in str(v))
    log(f"  BDE assegnati: {filled}")
    log(f"  DA VERIFICARE: {verificare}")

    log("Generazione report pattern...")
    df_tmp = df.copy()
    df_tmp["_TBL_KEY"] = make_table_key(df_tmp)

    bde_configs = [
        ("Codice NDG [DA VERIFICARE]", "Codice NDG", "NDG"),
        ("Codice Fiscale [DA VERIFICARE]", "Codice Fiscale", "CF"),
        ("Partita iva [DA VERIFICARE]", "Partita iva", "PIVA"),
    ]
    all_rows = []
    for bde_val, bde_tipo, bde_code in bde_configs:
        subset = df_tmp[df_tmp["BDE"] == bde_val]
        tables = subset.groupby("_TBL_KEY").apply(
            lambda g: g[["RGP", "RGP_PERIMETER", "CON_ID", "TBL_SCH",
                         "TBL_NM", "FLD_NM", "AP_BDE"]].to_dict("records")
        ).to_dict()
        for tbl_key, records in tables.items():
            fields = [r["FLD_NM"] for r in records]
            pattern = classify_pattern(fields, bde_code)
            for r in records:
                all_rows.append({
                    "BDE_TIPO": bde_tipo, "PATTERN": pattern,
                    "RGP": r["RGP"], "RGP_PERIMETER": r["RGP_PERIMETER"],
                    "CON_ID": r["CON_ID"], "TBL_SCH": r["TBL_SCH"],
                    "TBL_NM": r["TBL_NM"], "FLD_NM": r["FLD_NM"],
                    "AP_BDE": r["AP_BDE"], "N_CAMPI_TABELLA": len(fields),
                    "TUTTI_CAMPI": " | ".join(fields),
                })

    df_report = pd.DataFrame(all_rows).sort_values(["BDE_TIPO", "PATTERN", "TBL_NM", "FLD_NM"])

    log("Salvataggio...")
    out_bde = os.path.join(output_dir, "output_bde_v10_smart_tag.xlsx")
    out_rep = os.path.join(output_dir, "REPORT_COMPLETO_DA_VERIFICARE.xlsx")
    df.to_excel(out_bde, index=False, sheet_name="output_bde_v10_smart_tag")
    df_report.to_excel(out_rep, index=False, sheet_name="Tutti i DA VERIFICARE")

    log(f"\nCompletato!")
    log(f"  -> {out_bde}")
    log(f"  -> {out_rep}")
    log(f"  BDE totali: {filled} | DA VERIFICARE: {verificare} | Report: {len(df_report)} righe")

    return "\n".join(logs)


# ===================== WEB SERVER =====================

HTML_PAGE = r"""<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="UTF-8">
<title>BDE Assignment Tool</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: -apple-system, 'Segoe UI', Arial, sans-serif; background: #f4f6f9; color: #333; }
  .container { max-width: 640px; margin: 40px auto; background: #fff; border-radius: 12px;
               box-shadow: 0 2px 12px rgba(0,0,0,0.08); padding: 32px; }
  h1 { font-size: 22px; margin-bottom: 4px; }
  .subtitle { color: #888; font-size: 13px; margin-bottom: 24px; }
  .file-group { margin-bottom: 16px; }
  .file-group label { display: block; font-weight: 600; font-size: 14px; margin-bottom: 6px; }
  .file-row { display: flex; align-items: center; gap: 8px; }
  .file-name { flex: 1; background: #f0f0f0; border: 1px solid #ddd; border-radius: 6px;
               padding: 10px 12px; font-size: 13px; color: #888; min-height: 40px;
               display: flex; align-items: center; overflow: hidden; text-overflow: ellipsis; }
  .file-name.has-file { color: #2d7d2d; background: #e8f5e8; border-color: #b5d8b5; }
  input[type="file"] { display: none; }
  .btn-browse { background: #e8e8e8; border: 1px solid #ccc; border-radius: 6px;
                padding: 10px 16px; cursor: pointer; font-size: 13px; white-space: nowrap; }
  .btn-browse:hover { background: #ddd; }
  .output-group { margin: 20px 0; }
  .output-group input[type="text"] { width: 100%; padding: 10px 12px; border: 1px solid #ddd;
                                      border-radius: 6px; font-size: 13px; background: #f9f9f9; }
  .btn-run { display: block; width: 100%; padding: 14px; background: #2d7d2d; color: #fff;
             border: none; border-radius: 8px; font-size: 16px; font-weight: 600;
             cursor: pointer; margin: 20px 0 16px; }
  .btn-run:hover { background: #256b25; }
  .btn-run:disabled { background: #999; cursor: wait; }
  .log { background: #1e1e1e; color: #d4d4d4; font-family: 'Courier New', monospace;
         font-size: 12px; border-radius: 8px; padding: 16px; min-height: 120px;
         max-height: 300px; overflow-y: auto; white-space: pre-wrap; display: none; }
  .log.visible { display: block; }
  .spinner { display: inline-block; width: 16px; height: 16px; border: 2px solid #fff;
             border-top-color: transparent; border-radius: 50%; animation: spin 0.8s linear infinite;
             vertical-align: middle; margin-right: 8px; }
  @keyframes spin { to { transform: rotate(360deg); } }
</style>
</head>
<body>
<div class="container">
  <h1>BDE Assignment Tool</h1>
  <p class="subtitle">Assegnazione automatica BDE per NDG, Codice Fiscale, Partita IVA</p>

  <div class="file-group">
    <label>1. File Discovery (output_bde_v10_smart_tag)</label>
    <div class="file-row">
      <div class="file-name" id="name1">Nessun file selezionato</div>
      <button type="button" class="btn-browse" onclick="document.getElementById('file1').click()">Sfoglia</button>
      <input type="file" id="file1" accept=".xlsx,.xls" onchange="onFile(this, 'name1')">
    </div>
  </div>

  <div class="file-group">
    <label>2. File Perimetro (bde_CON_RGP_PERIMETER)</label>
    <div class="file-row">
      <div class="file-name" id="name2">Nessun file selezionato</div>
      <button type="button" class="btn-browse" onclick="document.getElementById('file2').click()">Sfoglia</button>
      <input type="file" id="file2" accept=".xlsx,.xls" onchange="onFile(this, 'name2')">
    </div>
  </div>

  <div class="file-group">
    <label>3. File Mapping (DM_Estrazione...Masking)</label>
    <div class="file-row">
      <div class="file-name" id="name3">Nessun file selezionato</div>
      <button type="button" class="btn-browse" onclick="document.getElementById('file3').click()">Sfoglia</button>
      <input type="file" id="file3" accept=".xlsx,.xls" onchange="onFile(this, 'name3')">
    </div>
  </div>

  <div class="output-group">
    <label style="font-weight:600; font-size:14px; display:block; margin-bottom:6px;">Cartella output</label>
    <input type="text" id="outdir" value="OUTPUT_DIR_PLACEHOLDER">
  </div>

  <button type="button" class="btn-run" id="btnRun" onclick="runProcess()">Esegui</button>
  <div class="log" id="log"></div>
</div>

<script>
function onFile(input, nameId) {
  var el = document.getElementById(nameId);
  if (input.files.length > 0) {
    el.textContent = input.files[0].name;
    el.classList.add('has-file');
  }
}

function fileToBase64(file) {
  return new Promise(function(resolve) {
    var reader = new FileReader();
    reader.onload = function() {
      resolve(reader.result.split(',')[1]);
    };
    reader.readAsDataURL(file);
  });
}

async function runProcess() {
  var f1 = document.getElementById('file1').files[0];
  var f2 = document.getElementById('file2').files[0];
  var f3 = document.getElementById('file3').files[0];
  var outdir = document.getElementById('outdir').value;

  if (!f1 || !f2 || !f3) { alert('Seleziona tutti e 3 i file.'); return; }

  var btn = document.getElementById('btnRun');
  var log = document.getElementById('log');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span>In esecuzione...';
  log.classList.add('visible');
  log.textContent = 'Upload file in corso...\n';

  var payload = {
    file_discovery: { name: f1.name, data: await fileToBase64(f1) },
    file_perimeter: { name: f2.name, data: await fileToBase64(f2) },
    file_mapping:   { name: f3.name, data: await fileToBase64(f3) },
    output_dir: outdir
  };

  fetch('/run', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  })
  .then(function(r) { return r.json(); })
  .then(function(data) {
    log.textContent = data.log;
    if (data.error) log.textContent += '\nERRORE: ' + data.error;
    btn.disabled = false;
    btn.textContent = 'Esegui';
  })
  .catch(function(err) {
    log.textContent += '\nErrore: ' + err;
    btn.disabled = false;
    btn.textContent = 'Esegui';
  });
}
</script>
</body>
</html>"""


class Handler(BaseHTTPRequestHandler):
    output_dir = "."

    def log_message(self, format, *args):
        pass

    def do_GET(self):
        html = HTML_PAGE.replace("OUTPUT_DIR_PLACEHOLDER", self.output_dir)
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write(html.encode("utf-8"))

    def do_POST(self):
        if self.path != "/run":
            self.send_response(404)
            self.end_headers()
            return

        length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(length)
        payload = json.loads(body)

        tmpdir = tempfile.mkdtemp()
        result = {"log": "", "error": None}

        try:
            paths = {}
            for key in ["file_discovery", "file_perimeter", "file_mapping"]:
                info = payload[key]
                fpath = os.path.join(tmpdir, info["name"])
                with open(fpath, "wb") as f:
                    f.write(base64.b64decode(info["data"]))
                paths[key] = fpath

            out_dir = payload.get("output_dir", self.output_dir)
            if not os.path.isdir(out_dir):
                os.makedirs(out_dir, exist_ok=True)

            start = datetime.now()
            log_text = run_bde_assignment(
                paths["file_discovery"],
                paths["file_perimeter"],
                paths["file_mapping"],
                out_dir,
            )
            elapsed = (datetime.now() - start).total_seconds()
            result["log"] = log_text + f"\n\nTempo: {elapsed:.1f} secondi"

        except Exception as e:
            result["error"] = str(e)
            result["log"] = result.get("log", "") or "Errore durante l'elaborazione."
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.end_headers()
        self.wfile.write(json.dumps(result).encode("utf-8"))


def main():
    port = 8787
    script_dir = os.path.dirname(os.path.abspath(__file__))
    Handler.output_dir = script_dir

    server = HTTPServer(("127.0.0.1", port), Handler)
    url = f"http://127.0.0.1:{port}"

    print(f"BDE Assignment Tool")
    print(f"Apri nel browser: {url}")
    print(f"Per chiudere: Ctrl+C")
    print()

    threading.Timer(1.0, lambda: webbrowser.open(url)).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nChiuso.")
        server.server_close()


if __name__ == "__main__":
    main()
