#!/usr/bin/env python3
"""
Script automatico per assegnazione BDE (Business Data Element) — Data Masking GDPR
==================================================================================
Input:
  - File discovery con campi e PDP_TAG
  - File perimetro con RGP_PERIMETER
  - File mapping (MappingBizTerm da DM_Estrazione)

Output:
  - Excel con BDE assegnati + RGP_PERIMETER
  - Report DA VERIFICARE con pattern classificati

Uso:
  python3 auto_bde_assignment.py

Configurare i path nella sezione CONFIG prima di lanciare.
"""

import pandas as pd
import re
import os
from datetime import datetime


# ===================== CONFIG =====================
# Modificare questi path per puntare ai file corretti

FILE_DISCOVERY = "../output_bde_v10_smart_tag_FINAL_v3.xlsx"
FILE_PERIMETER = "../bde_CON_RGP_PERIMETER.xlsx"
FILE_MAPPING = "../output_bde_v10_smart_tag.xlsx/DM_Estrazione_Discovery_e_Associazione_Regole_Masking_V6.xlsx"
SHEET_MAPPING_BIZTERM = "MappingBizTerm"

OUTPUT_DIR = "."
OUTPUT_BDE = "output_bde_v10_smart_tag.xlsx"
OUTPUT_REPORT = "REPORT_COMPLETO_DA_VERIFICARE.xlsx"

# PDP_TAG (lowercase) che ricevono un BDE
BDE_TAGS = {"ndg", "national id", "piva"}

# Colonne che identificano univocamente una tabella
TABLE_KEY_COLS = ["RGP", "CON_ID", "TBL_SCH", "TBL_NM"]

# ===================== FUNZIONI =====================


def load_data():
    """Carica tutti i file sorgente."""
    print("[1/5] Caricamento file...")
    df_main = pd.read_excel(FILE_DISCOVERY)
    df_per = pd.read_excel(FILE_PERIMETER)
    df_map = pd.read_excel(FILE_MAPPING, sheet_name=SHEET_MAPPING_BIZTERM)

    # Verifica allineamento discovery <-> perimetro
    for col in TABLE_KEY_COLS + ["FLD_NM"]:
        if not (df_main[col] == df_per[col]).all():
            raise ValueError(f"Disallineamento tra discovery e perimetro sulla colonna {col}")

    print(f"    Discovery: {len(df_main)} righe")
    print(f"    Perimetro: {len(df_per)} righe")
    print(f"    Mapping BizTerm: {len(df_map)} righe")
    return df_main, df_per, df_map


def build_mapping(df_map):
    """Costruisce dizionario case-insensitive PDP_TAG -> BizTerm."""
    return {
        str(row["PDP_TAG"]).strip().lower(): str(row["BizTerm"]).strip()
        for _, row in df_map.iterrows()
    }


def make_table_key(df):
    """Crea chiave unica per tabella."""
    return (
        df["RGP"].astype(str) + "|" +
        df["CON_ID"].astype(str) + "|" +
        df["TBL_SCH"].astype(str) + "|" +
        df["TBL_NM"].astype(str)
    )


def assign_bde(df_main, df_per, tag_to_bde):
    """Assegna BDE: singolo -> diretto, multiplo -> DA VERIFICARE."""
    print("[2/5] Assegnazione BDE...")

    df = df_main.copy()
    df["_TBL_KEY"] = make_table_key(df)
    df["_PDP_LOWER"] = df["PDP_TAG"].str.strip().str.lower()

    # Count per (tabella, pdp_tag)
    counts = df.groupby(["_TBL_KEY", "_PDP_LOWER"]).size().to_dict()

    # Assegna BDE
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
        if n == 1:
            bde_col.append(bde_name)
        else:
            bde_col.append(f"{bde_name} [DA VERIFICARE]")

    df["BDE"] = bde_col

    # Aggiungi RGP_PERIMETER
    df.insert(1, "RGP_PERIMETER", df_per["RGP_PERIMETER"].values)

    # Pulizia colonne temporanee
    df.drop(columns=["_TBL_KEY", "_PDP_LOWER"], inplace=True)

    # Stats
    filled = sum(1 for v in bde_col if v != "")
    verificare = sum(1 for v in bde_col if "DA VERIFICARE" in str(v))
    print(f"    BDE assegnati: {filled}")
    print(f"    Di cui DA VERIFICARE: {verificare}")
    return df


def classify_pattern(fields, bde_code):
    """Classifica il pattern dei campi multipli per una tabella."""
    fields_upper = [f.upper() for f in fields]
    fields_str = " ".join(fields_upper)

    # NEW/OLD
    has_new = any(k in f for f in fields_upper for k in ["_NEW", "_NUOV", "NEW_"])
    has_old = any(k in f for f in fields_upper for k in ["_OLD", "_VECCH", "OLD_", "_PREC"])
    if has_new and has_old:
        return "Pattern 3A: NEW/OLD"

    # NUMERATI
    if any(re.search(r"[\d]", f) for f in fields_upper):
        return "Pattern 3B: Numerati"

    # RUOLI
    role_kw = [
        "TITOL", "CONTRAE", "DELEG", "RICHIED", "BENEFIC", "ESECUTOR",
        "CAPO", "GARANT", "COINTEST", "FIDEIUSS", "ORDINANT", "ASSIC",
        "LOCATOR", "FORNIT", "OPERANTE", "MITTENT", "DESTINAT",
    ]
    if any(k in fields_str for k in role_kw):
        return "Pattern 4: Ruoli"

    # PURO vs SUFFISSO
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


def build_report(df_bde):
    """Genera report dettagliato dei casi DA VERIFICARE."""
    print("[3/5] Generazione report pattern...")

    df = df_bde.copy()
    df["_TBL_KEY"] = make_table_key(df)

    bde_configs = [
        ("Codice NDG [DA VERIFICARE]", "Codice NDG", "NDG"),
        ("Codice Fiscale [DA VERIFICARE]", "Codice Fiscale", "CF"),
        ("Partita iva [DA VERIFICARE]", "Partita iva", "PIVA"),
    ]

    all_rows = []
    for bde_val, bde_tipo, bde_code in bde_configs:
        subset = df[df["BDE"] == bde_val]
        tables = subset.groupby("_TBL_KEY").apply(
            lambda g: g[["RGP", "RGP_PERIMETER", "CON_ID", "TBL_SCH",
                         "TBL_NM", "FLD_NM", "AP_BDE"]].to_dict("records")
        ).to_dict()

        for tbl_key, records in tables.items():
            fields = [r["FLD_NM"] for r in records]
            pattern = classify_pattern(fields, bde_code)
            for r in records:
                all_rows.append({
                    "BDE_TIPO": bde_tipo,
                    "PATTERN": pattern,
                    "RGP": r["RGP"],
                    "RGP_PERIMETER": r["RGP_PERIMETER"],
                    "CON_ID": r["CON_ID"],
                    "TBL_SCH": r["TBL_SCH"],
                    "TBL_NM": r["TBL_NM"],
                    "FLD_NM": r["FLD_NM"],
                    "AP_BDE": r["AP_BDE"],
                    "N_CAMPI_TABELLA": len(fields),
                    "TUTTI_CAMPI": " | ".join(fields),
                })

    df_report = pd.DataFrame(all_rows)
    df_report = df_report.sort_values(["BDE_TIPO", "PATTERN", "TBL_NM", "FLD_NM"])
    print(f"    Righe report: {len(df_report)}")
    return df_report


def print_summary(df_bde, df_report):
    """Stampa riepilogo finale."""
    print("\n" + "=" * 60)
    print("RIEPILOGO")
    print("=" * 60)

    # BDE
    print("\n--- Assegnazione BDE ---")
    vals = df_bde["BDE"].replace("", pd.NA).dropna().value_counts()
    print(vals.to_string())

    # Pattern per tipo
    print("\n--- Pattern DA VERIFICARE ---")
    for bde in ["Codice NDG", "Codice Fiscale", "Partita iva"]:
        sub = df_report[df_report["BDE_TIPO"] == bde].copy()
        if len(sub) == 0:
            continue
        sub["_TK"] = make_table_key(sub)
        pat = sub.groupby("PATTERN").agg(
            tabelle=("_TK", "nunique")
        ).reset_index()
        total = sub["_TK"].nunique()
        auto = pat[pat["PATTERN"].str.contains("Puro|NEW|Numerati")]["tabelle"].sum()
        print(f"\n  {bde}: {total} tabelle DA VERIFICARE | Automatizzabili: {auto} ({auto/total*100:.0f}%)")
        for _, row in pat.iterrows():
            print(f"    {row['PATTERN']}: {row['tabelle']}")


def save_output(df_bde, df_report):
    """Salva i file Excel."""
    print("[4/5] Salvataggio file...")
    out_bde = os.path.join(OUTPUT_DIR, OUTPUT_BDE)
    out_rep = os.path.join(OUTPUT_DIR, OUTPUT_REPORT)

    df_bde.to_excel(out_bde, index=False, sheet_name="output_bde_v10_smart_tag")
    df_report.to_excel(out_rep, index=False, sheet_name="Tutti i DA VERIFICARE")

    print(f"    {out_bde}")
    print(f"    {out_rep}")


# ===================== MAIN =====================

def main():
    start = datetime.now()
    print(f"Avvio: {start.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"BDE target: {BDE_TAGS}\n")

    df_main, df_per, df_map = load_data()
    tag_to_bde = build_mapping(df_map)
    df_bde = assign_bde(df_main, df_per, tag_to_bde)
    df_report = build_report(df_bde)
    save_output(df_bde, df_report)
    print_summary(df_bde, df_report)

    elapsed = (datetime.now() - start).total_seconds()
    print(f"\n[5/5] Completato in {elapsed:.1f} secondi")


if __name__ == "__main__":
    main()
