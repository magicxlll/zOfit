# -*- coding: utf-8 -*-
"""
ETL script for zOfit QS Pro Add-in (Milestone 1).
Extracts, cleans, validates, and merges material reference tables from:
  1) Z-Bang tra thep hinh - inox.xlsx (REBAR, STEEL BOX, INOX, STEEL MESH)
  2) TKT.xlsm (Bang tra: V, L, I, U, H, Hop ma kem, Ong ma kem, Ong den, Xa go C, Xa go Z)

Outputs clean UTF-8-SIG CSV files into src/data/:
  - catalog_rebar.csv
  - catalog_materials.csv
"""
import csv
import math
import os
import re
import zipfile
import xml.etree.ElementTree as ET

NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
DATA_DIR = os.path.join(ROOT_DIR, "src", "data")


def col_to_num(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def read_sheet_dict(xlsx_path: str, sheet_name: str):
    with zipfile.ZipFile(xlsx_path) as z:
        ss = []
        if "xl/sharedStrings.xml" in z.namelist():
            r = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in r.findall("main:si", NS):
                ss.append("".join(t.text or "" for t in si.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")))
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rmap = {e.attrib["Id"]: e.attrib["Target"] for e in rels}
        target = None
        for s in wb.findall(".//main:sheet", NS):
            if s.attrib["name"] == sheet_name:
                rid = s.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
                target = "xl/" + rmap[rid].lstrip("/")
                break
        if not target:
            raise ValueError(f"Sheet {sheet_name} not found in {xlsx_path}")
        root = ET.fromstring(z.read(target))
        rows = {}
        for row_el in root.findall(".//main:sheetData/main:row", NS):
            rn = int(row_el.attrib["r"])
            rc = {}
            for c in row_el.findall("main:c", NS):
                m = re.match(r"([A-Z]+)(\d+)", c.attrib["r"])
                if not m:
                    continue
                col = m.group(1)
                t = c.attrib.get("t", "")
                v = c.find("main:v", NS)
                if t == "s" and v is not None and v.text is not None:
                    val = ss[int(v.text)]
                elif v is not None and v.text is not None:
                    val = v.text
                else:
                    val = None
                if val is not None:
                    rc[col] = val
            if rc:
                rows[rn] = rc
        return rows


def to_float(val, default=0.0):
    if val is None:
        return default
    s = str(val).strip().replace("\xa0", "").replace(" ", "")
    if not s or s == "-":
        return default
    # Handle Vietnamese decimal comma like "7,11"
    if "," in s and "." not in s:
        s = s.replace(",", ".")
    elif "," in s and "." in s:
        s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return default


def clean_str(val) -> str:
    if val is None:
        return ""
    return re.sub(r"\s+", " ", str(val).replace("\xa0", " ")).strip()


def build_rebar_table():
    """Standard rebar table Ø6..Ø40 per TCVN 1651 / BS 4449 (density 7850 kg/m3)."""
    diameters = [6, 8, 10, 12, 14, 16, 18, 20, 22, 25, 28, 30, 32, 36, 40]
    rows = []
    for d in diameters:
        area_mm2 = round(math.pi * (d ** 2) / 4.0, 2)
        wt_m = round(math.pi * ((d / 1000.0) ** 2) / 4.0 * 7850.0, 3)
        wt_117 = round(wt_m * 11.7, 2)
        wt_120 = round(wt_m * 12.0, 2)
        bar_type = "Thép cuộn (Coil)" if d <= 8 else "Thép cây (Bar)"
        rows.append({
            "Diameter_mm": d,
            "Code": f"D{d}",
            "Area_mm2": area_mm2,
            "Weight_kg_m": wt_m,
            "Weight_11m7_kg": wt_117,
            "Weight_12m0_kg": wt_120,
            "Type": bar_type,
        })
    return rows


def extract_all_materials():
    z_file = os.path.join(ROOT_DIR, "Z-Bang tra thep hinh - inox.xlsx")
    tkt_file = os.path.join(ROOT_DIR, "TKT.xlsm")
    items = []
    seen_keys = set()

    def add_item(cat, subcat, spec, unit, std_len, wt_unit, wt_bar, price, source):
        spec_clean = clean_str(spec)
        if not spec_clean or wt_unit <= 0:
            return
        dedup_key = (cat.lower(), spec_clean.lower(), round(std_len, 2))
        if dedup_key in seen_keys:
            return
        seen_keys.add(dedup_key)
        items.append({
            "Category": cat,
            "SubCategory": subcat,
            "Spec": spec_clean,
            "Unit": unit,
            "StdLen_m": round(std_len, 2) if std_len > 0 else "",
            "Weight_PerUnit_kg": round(wt_unit, 4),
            "Weight_PerBar_kg": round(wt_bar, 3) if wt_bar > 0 else "",
            "RefPrice_VND_kg": int(round(price)) if price > 0 else "",
            "Source": source,
        })

    # 1. Rebar Ø6..Ø40
    for rb in build_rebar_table():
        std_l = 11.7 if rb["Diameter_mm"] >= 10 else 0.0
        add_item(
            "Thép cốt bê tông",
            rb["Type"],
            f"Thép tròn Ø{rb['Diameter_mm']} (D{rb['Diameter_mm']})",
            "m",
            std_l,
            rb["Weight_kg_m"],
            rb["Weight_11m7_kg"] if std_l > 0 else 0.0,
            0,
            "TCVN 1651",
        )

    # 2. Z-Bang tra: STEEL BOX (~880 items)
    sb = read_sheet_dict(z_file, "STEEL BOX")
    cur_sub = "Thép hình"
    cat_map = {
        "Thép hình I": ("Thép hình I/H/U", "Thép hình I"),
        "THÉP HÌNH H": ("Thép hình I/H/U", "Thép hình H"),
        "Thép hình U, V, I, H mạ kẽm nhúng nóng": ("Thép hình U/V/L", "Hình U/V/I/H mạ kẽm nhúng nóng"),
        "Thép hộp mạ kẽm Hòa Phát": ("Thép hộp", "Hộp mạ kẽm Hòa Phát"),
        "Thép hộp đen Hòa Phát": ("Thép hộp", "Hộp đen Hòa Phát"),
        "Thép hộp đen cỡ lớn": ("Thép hộp", "Hộp đen cỡ lớn"),
        "Hộp mạ kẽm nhúng nóng": ("Thép hộp", "Hộp mạ kẽm nhúng nóng"),
        "Thép ống mạ kẽm Hòa Phát": ("Thép ống", "Ống mạ kẽm Hòa Phát"),
        "Thép ống đen Hòa Phát": ("Thép ống", "Ống đen Hòa Phát"),
        "Thép ống đen cỡ lớn": ("Thép ống", "Ống đen cỡ lớn"),
        "Thép ống đen siêu dày": ("Thép ống", "Ống đen siêu dày"),
        "Ống mạ kẽm nhúng nóng": ("Thép ống", "Ống mạ kẽm nhúng nóng"),
        "Bản mã các loại": ("Thép tấm & Bản mã", "Bản mã"),
    }
    cur_cat = "Thép hình I/H/U"
    for rn in sorted(sb.keys()):
        if rn <= 5:
            continue
        r = sb[rn]
        b = clean_str(r.get("B", ""))
        c = clean_str(r.get("C", ""))
        if b and not b.isdigit() and not c:
            if b in cat_map:
                cur_cat, cur_sub = cat_map[b]
            else:
                cur_sub = b
            continue
        if c:
            std_len = to_float(r.get("D", 0))
            wt_bar = to_float(r.get("E", 0))
            price = to_float(r.get("F", 0))
            if wt_bar > 0:
                wt_m = wt_bar / std_len if std_len > 0 else wt_bar
                # Fine-tune category for mixed galvanized shapes section
                cat = cur_cat
                if cur_sub == "Hình U/V/I/H mạ kẽm nhúng nóng":
                    cu = c.upper()
                    if "THÉP I" in cu or "THÉP H" in cu or "THÉP U" in cu:
                        cat = "Thép hình I/H/U"
                    else:
                        cat = "Thép hình V/L"
                add_item(cat, cur_sub, c, "m", std_len, wt_m, wt_bar, price, "Z-Bang tra")

    # 3. TKT.xlsm: Bang tra (V, L, I, U, H, C/Z Purlins, Box, Pipe)
    bt = read_sheet_dict(tkt_file, "Bang tra")
    for rn in sorted(bt.keys()):
        r = bt[rn]
        # Thép V đều cạnh (A, C) rows >= 7
        if rn >= 7 and "A" in r and "C" in r:
            wt = to_float(r["C"])
            if wt > 0:
                add_item("Thép hình V/L", "Thép V đều cạnh", r["A"], "m", 6.0, wt, wt * 6.0, 0, "TKT Bang tra")
        # Thép L không đều cạnh (E, G) rows >= 7
        if rn >= 7 and "E" in r and "G" in r:
            wt = to_float(r["G"])
            if wt > 0:
                add_item("Thép hình V/L", "Thép L không đều cạnh", r["E"], "m", 6.0, wt, wt * 6.0, 0, "TKT Bang tra")
        # Thép I (I, O) rows >= 5
        if rn >= 5 and "I" in r and "O" in r:
            wt = to_float(r["O"])
            if wt > 0:
                add_item("Thép hình I/H/U", "Thép hình I (TCVN)", r["I"], "m", 12.0, wt, wt * 12.0, 0, "TKT Bang tra")
        # Thép U (Q, W) rows >= 5
        if rn >= 5 and "Q" in r and "W" in r:
            wt = to_float(r["W"])
            spec = clean_str(r["Q"]).replace("[", "U")
            if wt > 0:
                add_item("Thép hình I/H/U", "Thép hình U (TCVN)", spec, "m", 12.0, wt, wt * 12.0, 0, "TKT Bang tra")
        # Thép H (AK, AL) rows >= 5
        if rn >= 5 and "AK" in r and "AL" in r:
            wt = to_float(r["AL"])
            if wt > 0:
                add_item("Thép hình I/H/U", "Thép hình H (TCVN)", r["AK"], "m", 12.0, wt, wt * 12.0, 0, "TKT Bang tra")
        # Xà gồ C (AV, BC) rows >= 6
        if rn >= 6 and "AV" in r and "BC" in r:
            wt = to_float(r["BC"])
            if wt > 0:
                add_item("Xà gồ C/Z", "Xà gồ C", r["AV"], "m", 6.0, wt, wt * 6.0, 0, "TKT Bang tra")
        # Xà gồ Z (BL, BU) rows >= 6
        if rn >= 6 and "BL" in r and "BU" in r:
            wt = to_float(r["BU"])
            if wt > 0:
                add_item("Xà gồ C/Z", "Xà gồ Z", r["BL"], "m", 6.0, wt, wt * 6.0, 0, "TKT Bang tra")

    # 4. Z-Bang tra: INOX (Round pipes R5..R26, Box tubes R33..R51, Solid rods R62..R63)
    inox = read_sheet_dict(z_file, "INOX")
    # Round inox pipes (kg per 6m bar)
    if 5 in inox:
        thick_cols = {col: to_float(v) for col, v in inox[5].items() if col != "B" and to_float(v) > 0}
        for rn in range(6, 27):
            if rn not in inox or "B" not in inox[rn]:
                continue
            dia = clean_str(inox[rn]["B"])
            for col, thick in sorted(thick_cols.items(), key=lambda kv: col_to_num(kv[0])):
                if col not in inox[rn]:
                    continue
                wt6 = to_float(inox[rn][col])
                # Fix verified typo in source sheet: INOX!C6 = 3.35 for D8x0.3 (theoretical = 0.35 kg/6m)
                if rn == 6 and col == "C" and abs(wt6 - 3.35) < 0.01:
                    wt6 = 0.35
                if wt6 > 0:
                    add_item("Inox", "Ống tròn Inox", f"Inox tròn Ø{dia} x {thick}mm", "m", 6.0, wt6 / 6.0, wt6, 0, "Z-Bang tra (cleaned)")
    # Rectangular / square inox tubes (R33..R55)
    if 33 in inox:
        thick_cols2 = {col: to_float(v) for col, v in inox[33].items() if col != "B" and to_float(v) > 0}
        for rn in range(34, 56):
            if rn not in inox or "B" not in inox[rn]:
                continue
            dim = clean_str(inox[rn]["B"])
            if "x" not in dim.lower():
                continue
            for col, thick in sorted(thick_cols2.items(), key=lambda kv: col_to_num(kv[0])):
                if col not in inox[rn]:
                    continue
                wt6 = to_float(inox[rn][col])
                if wt6 > 0:
                    add_item("Inox", "Hộp Inox", f"Inox hộp {dim} x {thick}mm", "m", 6.0, wt6 / 6.0, wt6, 0, "Z-Bang tra")
    # Solid inox rods (R62..R63)
    if 62 in inox and 63 in inox:
        for col, dval in sorted(inox[62].items(), key=lambda kv: col_to_num(kv[0])):
            if col == "B":
                continue
            d = to_float(dval)
            wt6 = to_float(inox[63].get(col, 0))
            if d > 0 and wt6 > 0:
                add_item("Inox", "Cây đặc Inox", f"Cây đặc Inox Ø{d:g}mm", "m", 6.0, wt6 / 6.0, wt6, 0, "Z-Bang tra")

    # 5. Z-Bang tra: STEEL MESH (kg/m2)
    mesh = read_sheet_dict(z_file, "STEEL MESH")
    cur_mtype = "Lưới thép hàn ô vuông"
    for rn in sorted(mesh.keys()):
        if rn < 8:
            continue
        r = mesh[rn]
        b = clean_str(r.get("B", ""))
        if "chữ nhật" in b.lower():
            cur_mtype = "Lưới thép hàn ô chữ nhật"
            continue
        if b.startswith("D") and "F" in r:
            desc = clean_str(r.get("C", b))
            grid = clean_str(r.get("D", ""))
            wt_m2 = to_float(r["F"])
            if wt_m2 > 0:
                add_item("Lưới thép hàn", cur_mtype, f"{b} ({grid}mm) - {desc}", "m2", 0.0, wt_m2, 0.0, 0, "Z-Bang tra")

    # 6. Standard Steel Plate thicknesses (Thép tấm 1mm..30mm, kg/m2 = t * 7.85)
    for t in [1.0, 1.2, 1.5, 2.0, 2.5, 3.0, 4.0, 5.0, 6.0, 8.0, 10.0, 12.0, 14.0, 16.0, 18.0, 20.0, 22.0, 25.0, 30.0]:
        wt_m2 = round(t * 7.85, 3)
        add_item("Thép tấm & Bản mã", "Thép tấm phẳng", f"Thép tấm dày {t:g}mm (SS400/Q235)", "m2", 0.0, wt_m2, 0.0, 0, "TCVN 7850 kg/m3")

    return items


def write_csv(path, fieldnames, rows):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)


def main():
    rebar = build_rebar_table()
    write_csv(
        os.path.join(DATA_DIR, "catalog_rebar.csv"),
        ["Diameter_mm", "Code", "Area_mm2", "Weight_kg_m", "Weight_11m7_kg", "Weight_12m0_kg", "Type"],
        rebar,
    )
    materials = extract_all_materials()
    write_csv(
        os.path.join(DATA_DIR, "catalog_materials.csv"),
        ["Category", "SubCategory", "Spec", "Unit", "StdLen_m", "Weight_PerUnit_kg", "Weight_PerBar_kg", "RefPrice_VND_kg", "Source"],
        materials,
    )
    from collections import Counter
    counts = Counter(m["Category"] for m in materials)
    print(f"Extracted {len(rebar)} rebar diameters -> src/data/catalog_rebar.csv")
    print(f"Extracted {len(materials)} total material specs -> src/data/catalog_materials.csv")
    for k, v in counts.most_common():
        print(f"  - {k}: {v}")


if __name__ == "__main__":
    main()
