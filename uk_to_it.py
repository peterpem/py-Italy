# -*- coding: utf-8 -*-
import os, re, sys, csv, random
import numpy as np
import pandas as pd
from collections import defaultdict
from openpyxl import load_workbook

# ==================== BUSINESS RULES ====================

ITEM_TAIL = " - Rear Sides & Rear Window - Custom Fit, UV Protection, Heat & Glare Reduction, High Performance"

def normalize_item_name(name: str) -> str:
    if not isinstance(name, str) or "TINTCOM" not in name:
        return name
    m = re.search(r"for\s+(.+?)\s+-\s+(?:\((\d+%[^)]*)\)\s+-\s+)?", name, flags=re.I)
    if not m:
        return name
    car = m.group(1).strip()
    shade = (m.group(2) or "").strip()
    return (f"TINTCOM Pre-Cut Window Tint Film for {car} - ({shade}){ITEM_TAIL}"
            if shade else f"TINTCOM Pre-Cut Window Tint Film for {car}{ITEM_TAIL}")

# Ако искаш италиански етикети – смени стойностите вдясно.
TONE_MAP = {
    "05% Limo Black":   "05% Limo Black",
    "20% Dark Smoke":   "20% Dark Smoke",
    "35% Medium Smoke": "35% Medium Smoke",
    "50% Light Smoke":  "50% Light Smoke",
    "70% Ultra Light":  "70% Ultra Light",
}

IT_KEYWORD_SETS = [
    "pellicola oscurante vetri auto, pellicola vetri auto, oscuramento vetri auto, oscurare vetri auto",
    "pellicola vetri auto, oscuramento vetri auto, oscurare vetri auto, vetri oscurati",
    "oscuramento vetri auto, oscurare vetri auto, vetri oscurati, pellicola oscurante",
    "oscurare vetri auto, vetri oscurati, pellicola oscurante, oscuramento vetri",
    "vetri oscurati, pellicola oscurante, oscuramento vetri, pellicola oscurante vetri auto pretagliata",
    "pellicola oscurante, oscuramento vetri, pellicola oscurante vetri auto pretagliata, kit pellicola oscurante vetri auto",
    "oscuramento vetri, pellicola oscurante vetri auto pretagliata, kit pellicola oscurante vetri auto, oscura vetri auto",
    "pellicola oscurante vetri auto pretagliata, kit pellicola oscurante vetri auto, oscura vetri auto, pellicola oscurante auto",
    "kit pellicola oscurante vetri auto, oscura vetri auto, pellicola oscurante auto, pellicola oscurante vetri auto pre-tagliata",
    "oscura vetri auto, pellicola oscurante auto, pellicola oscurante vetri auto pre-tagliata, vetri auto oscurati",
    "pellicola oscurante auto, pellicola oscurante vetri auto pre-tagliata, vetri auto oscurati, oscuramento",
    "pellicola oscurante vetri auto pre-tagliata, vetri auto oscurati, oscuramento, pellicole oscuranti vetri auto",
    "vetri auto oscurati, oscuramento, pellicole oscuranti vetri auto, pellicola vetro oscurante",
    "oscuramento, pellicole oscuranti vetri auto, pellicola vetro oscurante, oscurante vetri auto",
    "pellicole oscuranti vetri auto, pellicola vetro oscurante, oscurante vetri auto, pellicola oscuramento vetri",
    "pellicola vetro oscurante, oscurante vetri auto, pellicola oscuramento vetri, pellicola vetri oscurati auto",
    "oscurante vetri auto, pellicola oscuramento vetri, pellicola vetri oscurati auto, pellicola auto vetri",
    "pellicola oscuramento vetri, pellicola vetri oscurati auto, pellicola auto vetri, pellicola per oscurare vetri auto",
]
DEFAULT_GENERIC_KEYWORDS = IT_KEYWORD_SETS[0]

COLUMNS_TO_COPY = []  # ако е празно -> копира всички съвпадащи по име
PRICE_COLS_HINTS = ["price","our_price","standard_price","list_price",
                    "minimum_seller_allowed_price","maximum_seller_allowed_price"]

# ==================== ЛОГ/КОНСТАНТИ ====================

APPLIED_LOG = defaultdict(int)

# колони само с ТОЧНА замяна (никакъв regex)
EXACT_ONLY_FIELDS = {
    "variation_theme","package_level","material_type","color_map","unit_count_type",
    "length_longer_edge_unit_of_measure","width_shorter_edge_unit_of_measure",
    "package_length_unit_of_measure","fulfillment_center_id","is_fragile",
    "package_weight_unit_of_measure","country_of_origin",
    "compliance_media_content_type1","gpsr_safety_attestation",
    "gpsr_manufacturer_reference_email_address",
    "compliance_media_source_location1","compliance_media_content_language1",
    "product_tax_code","condition_type"
}

# дълги текстови колони – позволяваме regex
LONG_TEXT_FIELDS = {
    "item_name","product_description","bullet_point1","bullet_point2",
    "bullet_point3","bullet_point4","bullet_point5","generic_keywords","color_name"
}

# ==================== ХЕЛПЕРИ ====================

def translate_uk_description_to_it_html(uk_description: str, name: str = "", color: str = "") -> str:
    if not isinstance(uk_description, str):
        return ""
    def norm(s):
        if not isinstance(s, str): return ""
        s = (s.replace("\u2019","'").replace("\u2018","'")
               .replace("\u201c",'"').replace("\u201d",'"'))
        s = s.replace("<br><br>","\n").replace("<br>","\n")
        s = re.sub(r"[ \t]+"," ", s)
        return s.strip()
    text = norm(uk_description)
    if not text: return ""

    # извади име/нюанс ако не са подадени
    def extract_name(src: str) -> str:
        m = re.search(r"for\s+(.+?)\s+-\s+", src, flags=re.I)
        return m.group(1).strip() if m else ""
    def extract_color(src: str) -> str:
        m = re.search(r"\{([^}]+)\}", src)
        if m: return m.group(1).strip()
        m = re.search(r"\(([^)]+)\)\s+-\s+Rear", src, flags=re.I)
        return m.group(1).strip() if m else ""

    vehicle_name = name.strip() or extract_name(text)
    shade = color.strip() or extract_color(text)

    # ако няма шаблон – върни нормализирания текст като HTML
    key = "enhance your vehicle"
    if key not in text.lower():
        return text.replace("\n","<br>")

    color_it_map = {
        "05% Limo Black": "05% Super Nero",
        "20% Dark Smoke": "20% Fumo Scuro",
        "35% Medium Smoke": "35% Fumo Medio",
        "50% Light Smoke": "50% Fumo Chiaro",
        "70% Ultra Light": "70% Ultra Leggera",
    }
    color_it = color_it_map.get(shade, shade)
    name_disp = vehicle_name or "[NAME]"
    color_frag = f" - ({color_it})" if color_it else ""

    paragraphs = [
        ("<b>TINTCOM Pellicola Oscurante Vetri Auto Pre-Tagliata per "
         f"{name_disp}{color_frag}</b>"),
        ("Migliora comfort, estetica e sicurezza della tua auto con il nostro kit "
         "<b>Pre-Tagliato su Misura</b>. Realizzato per adattarsi perfettamente al modello "
         "della tua vettura, il kit include pellicole per <b>Vetri Posteriori e Lunotto</b>."),
        ("La pellicola blocca fino al 99% dei raggi UV, proteggendo interni, sedili e plancia "
         "da scolorimento e crepe. Riduce l'abbagliamento solare e il calore interno, rendendo "
         "la guida più confortevole durante le giornate calde."),
        ("Il materiale <b>High Performance Series</b> offre visibilità ottimale, mantenendo "
         "l'interno dell'auto più fresco e privato. Grazie alla finitura antigraffio e alla "
         "tecnologia anti-bolle, la pellicola rimane liscia e trasparente nel tempo."),
        ("<b>Facile da installare</b> – non serve smontare pannelli o tagliare la pellicola. "
         "Tutto è pronto per un'installazione rapida fai-da-te con le istruzioni incluse."),
        ("<b>Tonalità disponibili:</b><br>"
         "5% (Super Nero) – massimo oscuramento, massima privacy.<br>"
         "20% (Fumo Scuro) – equilibrio perfetto tra privacy e visibilità.<br>"
         "35% (Fumo Medio) – aspetto elegante e buona visibilità notturna.<br>"
         "50% (Fumo Chiaro) – visibilità eccellente, effetto sobrio e moderno."),
        ("<b>TINTCOM</b> – Soluzioni di protezione auto su misura, con materiali di alta qualità e innovazione.")
    ]
    return "<br><br>".join(paragraphs)

def normalize_html_text(s: str) -> str:
    if not isinstance(s, str):
        return s
    s = (s.replace("\u2019", "'").replace("\u2018", "'")
           .replace("\u201c", '"').replace("\u201d", '"'))
    s = s.replace("<br><br>", "\n").replace("<br>", "\n")
    s = re.sub(r"[ \t]+", " ", s)
    return s.strip()

def load_map_sheet(xls: pd.ExcelFile, sheet_name: str) -> pd.DataFrame:
    if sheet_name not in xls.sheet_names:
        return pd.DataFrame(columns=["field","find","replace"])
    df = xls.parse(sheet_name).fillna("")
    cols = {str(c).strip().lower(): c for c in df.columns}
    df = df.rename(columns={cols.get("field","field"): "field",
                            cols.get("find","find"): "find",
                            cols.get("replace","replace"): "replace"})
    for k in ["field","find","replace"]:
        if k not in df.columns: df[k] = ""
    return df[["field","find","replace"]]

def expand_fields(df_columns, field_cell: str):
    return [c.strip() for c in str(field_cell).split("|")
            if c.strip() and c.strip() in df_columns]

def apply_exact_rules(df, rules: pd.DataFrame):
    for _, r in rules.iterrows():
        cols = expand_fields(df.columns, r.get("field",""))
        find = str(r.get("find",""))
        rep  = str(r.get("replace",""))
        if not cols or find == "":
            continue
        for col in cols:
            s = df[col].astype(str)
            mask = s == find
            cnt = int(mask.sum())
            if cnt:
                df.loc[mask, col] = rep
                APPLIED_LOG[f"{col}::EXACT {find}->{rep}"] += cnt
    return df

def apply_text_rules(df, rules: pd.DataFrame, case_insensitive=True):
    flags = re.IGNORECASE if case_insensitive else 0
    for _, r in rules.iterrows():
        cols = expand_fields(df.columns, r.get("field",""))
        find = str(r.get("find","")).strip()
        rep  = str(r.get("replace",""))
        if not cols or find == "":
            continue
        wordish = re.match(r"^[A-Za-z0-9% .\-]+$", find) is not None
        pattern = r"\b{}\b".format(re.escape(find)) if wordish else re.escape(find)
        for col in cols:
            if col in EXACT_ONLY_FIELDS:
                continue  # никога regex върху системните полета
            if col not in LONG_TEXT_FIELDS:
                continue
            before = df[col].astype(str).map(normalize_html_text)
            after  = before.str.replace(pattern, rep, regex=True, flags=flags)
            changed = int((before != after).sum())
            if changed:
                df[col] = after
                APPLIED_LOG[f"{col}::REGEX {find}->{rep}"] += changed
    return df

def apply_value_map(df, cols, value_map):
    for c in cols:
        if c in df.columns:
            df[c] = df[c].astype(str).map(lambda v: value_map.get(v, v))
    return df

def guess_price_cols(df):
    cols = []
    for c in df.columns:
        lc = str(c).lower()
        if any(h in lc for h in PRICE_COLS_HINTS):
            cols.append(c)
    return cols

def file_out_name(uk_path):
    base = os.path.basename(uk_path)
    if base.lower().startswith("uk-"):
        return os.path.join(os.path.dirname(uk_path), "IT-" + base[3:])
    name, ext = os.path.splitext(base)
    return os.path.join(os.path.dirname(uk_path), f"IT-{name}{ext}")

def cleanup_nulls(df: pd.DataFrame) -> pd.DataFrame:
    return df.map(lambda v: "" if (
        (isinstance(v, float) and pd.isna(v)) or
        (isinstance(v, str) and v.strip().lower() in ("nan","none","nat"))
    ) else v)

def write_into_it_template(template_path: str, df: pd.DataFrame, out_path: str):
    df = df.replace({np.nan: ""}).fillna("")
    wb = load_workbook(template_path)
    ws = wb.active
    col_index_by_name = {ws.cell(row=3, column=c).value: c
                         for c in range(1, ws.max_column + 1)}
    start_row = 4
    if ws.max_row >= start_row:
        for r in range(start_row, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                ws.cell(row=r, column=c, value=None)
    r = start_row
    for _, row in df.iterrows():
        for col_name, val in row.items():
            if col_name in col_index_by_name:
                c = col_index_by_name[col_name]
                ws.cell(row=r, column=c, value=(None if val == "" else val))
        r += 1
    wb.save(out_path)

def write_rules_log(out_path):
    out_dir = os.path.dirname(out_path) or "."
    path = os.path.join(out_dir, "applied_rules_log.csv")
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f); w.writerow(["field","rule","count"])
        for k, v in APPLIED_LOG.items():
            field, rule = (k.split("::",1)+[""])[:2]
            w.writerow([field, rule, v])
    return path

# ==================== ОСНОВНА ФУНКЦИЯ ====================

def uk_to_it(uk_file, it_template_file, out_path=None, find_replace_xlsx=None):
    # 1) UK таблица (ред 3 = header)
    uk = pd.read_excel(uk_file, header=2, dtype=str).fillna("")

    # 2) IT шаблон (взимаме имената от ред 3)
    it_raw = pd.read_excel(it_template_file, header=None, dtype=str).fillna("")
    it_header_names = list(it_raw.iloc[2])
    it_out = pd.DataFrame(columns=it_header_names)

    # 3) Копиране по име
    cols = COLUMNS_TO_COPY or [c for c in uk.columns if c in it_out.columns]
    for c in cols:
        it_out[c] = uk[c]

    # 4) Правила
    if find_replace_xlsx:
        xls = pd.ExcelFile(find_replace_xlsx)
        text_map  = load_map_sheet(xls, "text_map")
        words_map = load_map_sheet(xls, "words_find_replace") if "words_find_replace" in xls.sheet_names else (
                    load_map_sheet(xls, "words_find_replece") if "words_find_replece" in xls.sheet_names
                    else pd.DataFrame(columns=["field","find","replace"]))
        sku_map   = load_map_sheet(xls, "sku_map")
        price_map_df = xls.parse("price_map").fillna("") if "price_map" in xls.sheet_names else pd.DataFrame(columns=["find","replace"])

        # 4.1 точни замени
        it_out = apply_exact_rules(it_out, text_map)
        it_out = apply_exact_rules(it_out, words_map)

        # 4.2 sku точни замени
        for _, r in sku_map.iterrows():
            cols_ = expand_fields(it_out.columns, r.get("field",""))
            find = str(r.get("find","")); rep = str(r.get("replace",""))
            if not cols_ or find == "": continue
            for col in cols_:
                s = it_out[col].astype(str)
                mask = s == find
                cnt = int(mask.sum())
                if cnt:
                    it_out.loc[mask, col] = rep
                    APPLIED_LOG[f"{col}::SKU_MAP {find}->{rep}"] += cnt

        # 4.3 regex – само за дълги текстови полета
        it_out = apply_text_rules(it_out, text_map)
        it_out = apply_text_rules(it_out, words_map)

        # 4.4 цени
        price_cols = guess_price_cols(it_out)
        value_map  = dict(zip(price_map_df["find"].astype(str), price_map_df["replace"].astype(str)))
        it_out     = apply_value_map(it_out, price_cols, value_map)

    # 5) Бизнес правила
    if "update_delete" in it_out.columns:
        it_out["update_delete"] = it_out["update_delete"].replace({"Update":"Aggiorna","update":"Aggiorna"})
    if "item_name" in it_out.columns:
        it_out["item_name"] = it_out["item_name"].astype(str).map(normalize_item_name)
    if "material_type" in it_out.columns:
        it_out["material_type"] = it_out["material_type"].replace({"Polyester":"Poliestere"})
    if "color_map" in it_out.columns:
        it_out["color_map"] = it_out["color_map"].replace({"Black":"nero","black":"nero"})
    for col in ("color_name","size_name"):
        if col in it_out.columns:
            it_out[col] = it_out[col].replace(TONE_MAP)

    # „Heat Shrink“ -> празно
    if "installation_type" in it_out.columns:
        mask = it_out["installation_type"].astype(str).str.strip().str.lower().eq("heat shrink")
        it_out.loc[mask, "installation_type"] = ""

    # product_description: генерирай италиански HTML по UK шаблона (ако има)
    if "product_description" in it_out.columns:
        it_out["product_description"] = it_out.apply(
            lambda row: translate_uk_description_to_it_html(
                row.get("product_description",""),
                row.get("item_name",""),
                row.get("color_name","")
            ),
            axis=1
        )

    # keywords – попълни правилната колона
    if "generic_keywords" in it_out.columns:
        it_out["generic_keywords"] = DEFAULT_GENERIC_KEYWORDS
    elif "generic_keyword" in it_out.columns:
        it_out["generic_keyword"] = DEFAULT_GENERIC_KEYWORDS

    # 6) Почистване и запис
    it_out = cleanup_nulls(it_out)
    out_file = out_path or file_out_name(uk_file)
    write_into_it_template(it_template_file, it_out, out_file)
    log_path = write_rules_log(out_file)
    print(f"Done → {out_file}")
    print(f"Rules log → {log_path}")

# ==================== RUN ====================
if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python uk_to_it.py <UK.xlsx> <IT_template.xlsx> [Find_and_Replace.xlsx] [OUT.xlsx]")
        sys.exit(1)
    uk   = sys.argv[1]
    it_t = sys.argv[2]
    fr   = sys.argv[3] if len(sys.argv) >= 4 else None
    out  = sys.argv[4] if len(sys.argv) >= 5 else None
    uk_to_it(uk, it_t, out, fr)
