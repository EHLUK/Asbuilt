"""
As-Built Drawing Compiler — HPC HK2794
Core logic, extracted from the tested Claude skill.
"""

import os
import re
import json
import shutil
import subprocess
import tempfile
import zipfile
from collections import Counter
from pathlib import Path

import pdfplumber
from PIL import Image


# ── ECS / TRN ────────────────────────────────────────────────────────────────

ECS_PATTERN   = re.compile(r'^\d[A-Z]{2,4}\d{4}-\d{2}-\d+$')
DB_PATTERN    = re.compile(r'ENG-GSC-WS08-\d+')
DELIVERY_PAT  = re.compile(r'S\.S Welded.*?DO[\d,\s&]+\d+')
TC_REF_PAT    = re.compile(r'TC Ref[:\s]+(\d{6,})')
AREA_PAT      = re.compile(r'[A-Z]{3}-\d{2}')

DUCTBOOK_INFO = {
    'ENG-GSC-WS08-02385': ('5.6.2 Ductwork Book HKX-08 EL ]+14,80m;+19,50m]', '01'),
    'ENG-GSC-WS08-02393': ('5.6.2 Ductwork Book HKX-08 EL ]+14,80m;+19,50m]', '01'),
    'ENG-GSC-WS08-02978': ('5.6.2 Ductwork Book HWX-05', '02'),
    'ENG-GSC-WS08-03005': ('5.6.2 Ductwork Book HVL-03', '02'),
    'ENG-GSC-WS08-03046': ('5.6.2 Ductwork Book HWX-04', '02'),
}

# Image sizing — calibrated from LibreOffice (15.24 cm × 21.56 cm)
FIXED_CX = 5_486_400
FIXED_CY = 7_761_600


def extract_trn(trn_path: str, progress=None) -> dict:
    """Extract ECS codes, ductbook groups and header fields from TRN PDF."""
    ecs_codes = []
    ecs_ductbook = {}
    seen = set()
    delivery_ref = ""
    tc_ref = ""
    area_codes = []

    if progress:
        progress("Reading TRN PDF…")

    with pdfplumber.open(trn_path) as pdf:
        for page in pdf.pages:
            # Extract header fields from raw text on first page
            text = page.extract_text() or ""
            if not delivery_ref:
                m = DELIVERY_PAT.search(text)
                if m:
                    delivery_ref = m.group(0).strip()
            if not tc_ref:
                m = TC_REF_PAT.search(text)
                if m:
                    tc_ref = m.group(1)
            if not area_codes:
                found = AREA_PAT.findall(text)
                area_codes = list(dict.fromkeys(found))  # deduplicated, ordered

            # Extract ECS codes from tables
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row:
                        continue
                    row_ecs, row_db = [], None
                    for cell in row:
                        if not cell:
                            continue
                        cell = str(cell).strip()
                        if ECS_PATTERN.match(cell) and cell not in seen:
                            row_ecs.append(cell)
                            seen.add(cell)
                        db_m = DB_PATTERN.search(cell)
                        if db_m:
                            row_db = db_m.group(0)
                    for ecs in row_ecs:
                        ecs_codes.append(ecs)
                        if row_db:
                            ecs_ductbook[ecs] = row_db

    return {
        "ecs_codes": ecs_codes,
        "ecs_ductbook": ecs_ductbook,
        "delivery_ref": delivery_ref,
        "tc_ref": tc_ref,
        "area_codes": area_codes,
    }


# ── Drawing matching ──────────────────────────────────────────────────────────

def _scan_pdf(pdf_path: str, ecs_set: set) -> dict:
    """Scan one drawing PDF and return {ecs: [(page_num, count), ...]}."""
    results = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            words = page.extract_words()
            ecs_on_page = []
            for w in words:
                txt = w['text'].strip()
                if ECS_PATTERN.match(txt) and txt in ecs_set:
                    ecs_on_page.append(txt)
                # Title block format: -DWKxxxx -> 1DWKxxxx
                if re.match(r'^-DW[A-Z]\d{4}-\d{2}-\d+$', txt):
                    candidate = '1' + txt[1:]
                    if candidate in ecs_set:
                        ecs_on_page.append(candidate)
            unique = list(set(ecs_on_page))
            for ecs in unique:
                results.setdefault(ecs, []).append((page_num, len(unique)))
    return results


def match_drawings(ecs_codes: list, ecs_ductbook: dict,
                   drawing_pdf_paths: list, progress=None) -> tuple[dict, list]:
    """
    Match each ECS code to its individual drawing page.
    Returns (matches dict, not_found list).
    matches = {ecs: {"pdf": path, "page": int}}
    """
    ecs_set = set(ecs_codes)
    ecs_page_candidates = {}

    for i, pdf_path in enumerate(drawing_pdf_paths):
        if progress:
            progress(f"Scanning {Path(pdf_path).name} ({i+1}/{len(drawing_pdf_paths)})…")
        pdf_results = _scan_pdf(pdf_path, ecs_set)
        for ecs, hits in pdf_results.items():
            for page_num, count in hits:
                ecs_page_candidates.setdefault(ecs, []).append((page_num, count, pdf_path))

    matches = {}
    for ecs, candidates in ecs_page_candidates.items():
        best = min(candidates, key=lambda x: x[1])  # min ECS count = individual page
        matches[ecs] = {"pdf": best[2], "page": best[0]}

    not_found = [c for c in ecs_codes if c not in matches]
    return matches, not_found


# ── Render + stamp ────────────────────────────────────────────────────────────

def render_and_stamp(matches: dict, stamp_path: str,
                     work_dir: str, progress=None) -> dict:
    """
    Render each matched page with pdftoppm, rotate 90° CW, apply stamp.
    Returns {ecs: stamped_png_path}
    """
    images_dir = os.path.join(work_dir, "images")
    stamped_dir = os.path.join(work_dir, "stamped")
    os.makedirs(images_dir, exist_ok=True)
    os.makedirs(stamped_dir, exist_ok=True)

    stamp_src = Image.open(stamp_path).convert("RGBA")
    STAMP_W = 170
    STAMP_H = int(stamp_src.height * STAMP_W / stamp_src.width)
    stamp_fit = stamp_src.resize((STAMP_W, STAMP_H), Image.LANCZOS)

    # Calibrated stamp position (portrait 1404×1985 px at 120 DPI)
    CELL_X1, CELL_Y1 = 350, 1283
    CELL_X2 = 750
    CELL_W = CELL_X2 - CELL_X1
    STAMP_X = CELL_X1 + (CELL_W - STAMP_W) // 2 + 71
    STAMP_Y = CELL_Y1 - 37

    stamped_paths = {}
    total = len(matches)

    for i, (ecs, info) in enumerate(matches.items()):
        if progress:
            progress(f"Rendering & stamping {i+1}/{total}: {ecs}…")

        pdf_path = info["pdf"]
        page_num = info["page"]
        out_prefix = os.path.join(images_dir, ecs)

        subprocess.run([
            "pdftoppm", "-r", "120", "-jpeg",
            "-f", str(page_num + 1), "-l", str(page_num + 1),
            pdf_path, out_prefix
        ], check=True, capture_output=True)

        candidates = sorted([
            f for f in os.listdir(images_dir)
            if f.startswith(os.path.basename(out_prefix)) and f.endswith(".jpg")
        ])
        if not candidates:
            continue

        img_path = os.path.join(images_dir, candidates[-1])
        img = Image.open(img_path).rotate(-90, expand=True).convert("RGBA")
        img.paste(stamp_fit, (STAMP_X, STAMP_Y), stamp_fit)
        out = os.path.join(stamped_dir, f"{ecs}.png")
        img.convert("RGB").save(out, "PNG")
        stamped_paths[ecs] = out

    return stamped_paths


# ── Word document build ───────────────────────────────────────────────────────

def _fit_image(img_w, img_h):
    cx = FIXED_CX
    cy = int(cx * img_h / img_w)
    if cy > FIXED_CY:
        cy = FIXED_CY
        cx = int(cy * img_w / img_h)
    return cx, cy


def _img_para(rid, ecs, img_path):
    img = Image.open(img_path)
    cx, cy = _fit_image(img.size[0], img.size[1])
    uid = abs(hash(ecs)) % 999990 + 1
    return f"""    <w:p>
      <w:pPr><w:pageBreakBefore w:val="1"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>
      <w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"
        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
        <wp:extent cx="{cx}" cy="{cy}"/>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        <wp:docPr id="{uid}" name="{ecs}"/>
        <wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:nvPicPr><pic:cNvPr id="0" name="{ecs}"/><pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1"/></pic:cNvPicPr></pic:nvPicPr>
              <pic:blipFill><a:blip r:embed="{rid}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
              <pic:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline></w:drawing></w:r>
    </w:p>"""


def build_docx(template_path: str, trn_data: dict, matches: dict,
               stamped_paths: dict, output_path: str,
               work_dir: str, progress=None) -> str:
    """
    Build the completed As-Built Word document.
    Returns output_path on success.
    """
    ecs_codes   = trn_data["ecs_codes"]
    ecs_ductbook = trn_data["ecs_ductbook"]
    delivery_ref = trn_data["delivery_ref"]
    tc_ref       = trn_data["tc_ref"]
    area_codes   = trn_data["area_codes"]

    docx_work = os.path.join(work_dir, "docx_work")
    if os.path.exists(docx_work):
        shutil.rmtree(docx_work)
    os.makedirs(docx_work)

    if progress:
        progress("Unpacking Word template…")

    # Unpack docx (it's a zip)
    with zipfile.ZipFile(template_path, 'r') as z:
        z.extractall(docx_work)

    doc_xml_path = os.path.join(docx_work, "word", "document.xml")
    rels_path    = os.path.join(docx_work, "word", "_rels", "document.xml.rels")
    ct_path      = os.path.join(docx_work, "[Content_Types].xml")
    media_dir    = os.path.join(docx_work, "word", "media")
    os.makedirs(media_dir, exist_ok=True)

    with open(doc_xml_path, 'r', encoding='utf-8') as f:
        xml = f.read()

    # ── 5a Text replacements ─────────────────────────────────────────────────
    if progress:
        progress("Filling in header fields…")

    area_str = " / ".join(dict.fromkeys(area_codes)) if area_codes else "HKX-08"
    do_str   = delivery_ref.replace("S.S Welded - ", "").replace("&", "&amp;") if delivery_ref else "DO??"

    matched_dbs = sorted(set(ecs_ductbook[e] for e in stamped_paths))
    first_db    = matched_dbs[0] if matched_dbs else "ENG-GSC-WS08-XXXXX"
    _, first_rev = DUCTBOOK_INFO.get(first_db, (first_db, "01"))

    replacements = [
        ('HK2794 As-Built Drawings Area XXXXXX',  area_str),
        ('HK2794 As-Built Drawings Area XXXXXXX', area_str),
        ('Ductwork Book Delivery Order XX',        f'DO {do_str}'),
        ('XXXXXXXX',                               tc_ref or "????????"),
        ('Stainless Steel Welded Duct Delivery Order XX associated',
         f'Stainless Steel Welded Duct Delivery Order {do_str} associated'),
        ('HK2794 ITP for Stainless Steel Welded Duct D0XX',
         f'HK2794 ITP for Stainless Steel Welded Duct {do_str}'),
        (' (TC number XXXXXXX). ', f' (TC number: {tc_ref or "___________"}). '),
        ('APPENDIX 1 – DUCTWORK BOOK: ENG-GSC-WS08-0XXX (Rev 0X)',
         f'APPENDIX 1 – DUCTWORK BOOK: {first_db} (Rev {first_rev})'),
    ]
    for old, new in replacements:
        xml = xml.replace(old, new)

    # ── 5b Ductwork Books table row ──────────────────────────────────────────
    def make_db_row(db, idx):
        desc, rev = DUCTBOOK_INFO.get(db, (db, '01'))
        pids = [f'0A{idx+j:06X}' for j in range(4)]
        return (
            f'<w:tr w14:paraId="{pids[0]}" w14:textId="FFFFFFFF" '
            f'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
            f'<w:tc><w:tcPr><w:tcW w:w="1418" w:type="dxa"/></w:tcPr>'
            f'<w:p w14:paraId="{pids[1]}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
            f'<w:r><w:t>{db}</w:t></w:r></w:p></w:tc>'
            f'<w:tc><w:tcPr><w:tcW w:w="4513" w:type="dxa"/></w:tcPr>'
            f'<w:p w14:paraId="{pids[2]}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
            f'<w:r><w:t>{desc}</w:t></w:r></w:p></w:tc>'
            f'<w:tc><w:tcPr><w:tcW w:w="851" w:type="dxa"/></w:tcPr>'
            f'<w:p w14:paraId="{pids[3]}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
            f'<w:r><w:t>{rev}</w:t></w:r></w:p></w:tc>'
            f'</w:tr>'
        )

    new_rows = ''.join(make_db_row(db, 0x100 + i * 10) for i, db in enumerate(matched_dbs))
    m = re.search(r'w14:paraId="30481C10"', xml)
    if m:
        tr_start = xml.rfind('<w:tr ', 0, m.start())
        tr_end   = xml.find('</w:tr>', m.start()) + len('</w:tr>')
        xml = xml[:tr_start] + new_rows + xml[tr_end:]

    # ── 5c ECS code table ────────────────────────────────────────────────────
    para_idx = 0x200

    def make_ecs_tables():
        nonlocal para_idx
        tables = []
        for db in matched_dbs:
            db_codes = [e for e in ecs_codes if ecs_ductbook.get(e) == db and e in stamped_paths]
            if not db_codes:
                continue
            hpid = f'0B{para_idx:06X}'; para_idx += 1
            rows_xml = (
                f'<w:tr><w:tc><w:tcPr><w:gridSpan w:val="4"/>'
                f'<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:tcPr>'
                f'<w:p w14:paraId="{hpid}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
                f'<w:pPr><w:jc w:val="center"/></w:pPr>'
                f'<w:r><w:rPr><w:b/></w:rPr><w:t>{db}</w:t></w:r></w:p></w:tc></w:tr>'
            )
            for i in range(0, len(db_codes), 4):
                chunk = db_codes[i:i+4]
                while len(chunk) < 4:
                    chunk.append('')
                rpid = f'0B{para_idx:06X}'; para_idx += 1
                cells = ''
                for code in chunk:
                    cp = f'0B{para_idx:06X}'; para_idx += 1
                    cells += (
                        f'<w:tc><w:tcPr><w:tcW w:w="2550" w:type="dxa"/></w:tcPr>'
                        f'<w:p w14:paraId="{cp}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
                        f'<w:r><w:t>{code}</w:t></w:r></w:p></w:tc>'
                    )
                rows_xml += f'<w:tr w14:paraId="{rpid}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">{cells}</w:tr>'
            tables.append(
                f'<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/>'
                f'<w:tblW w:w="0" w:type="auto"/></w:tblPr>'
                f'<w:tblGrid><w:gridCol w:w="2550"/><w:gridCol w:w="2550"/>'
                f'<w:gridCol w:w="2550"/><w:gridCol w:w="2551"/></w:tblGrid>'
                f'{rows_xml}</w:tbl>'
            )
        return '\n'.join(tables)

    m2 = re.search(
        r'<w:tblGrid>\s*<w:gridCol w:w="2550"/>\s*<w:gridCol w:w="2550"/>'
        r'\s*<w:gridCol w:w="2550"/>\s*<w:gridCol w:w="2551"/>', xml)
    if m2:
        tbl_start = xml.rfind('<w:tbl>', 0, m2.start())
        tbl_end   = xml.find('</w:tbl>', m2.start()) + len('</w:tbl>')
        xml = xml[:tbl_start] + make_ecs_tables() + xml[tbl_end:]

    # ── 5d Register images + build appendices ────────────────────────────────
    if progress:
        progress("Embedding drawings into document…")

    with open(rels_path, 'r', encoding='utf-8') as f:
        rels = f.read()

    rId_map = {}
    rId_num = 100
    for ecs, img_path in stamped_paths.items():
        ecs_safe = ecs.replace('-', '_')
        dst = os.path.join(media_dir, f"drawing_{ecs_safe}.png")
        shutil.copy2(img_path, dst)
        rid = f'rId{rId_num}'
        rId_map[ecs] = rid
        target = f'media/drawing_{ecs_safe}.png'
        if rid not in rels:
            rels = rels.replace('</Relationships>',
                f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{target}"/>\n</Relationships>')
        rId_num += 1

    with open(rels_path, 'w', encoding='utf-8') as f:
        f.write(rels)

    # Ensure PNG content type registered
    with open(ct_path, 'r', encoding='utf-8') as f:
        ct = f.read()
    if 'Extension="png"' not in ct:
        ct = ct.replace('</Types>', '<Default Extension="png" ContentType="image/png"/>\n</Types>')
        with open(ct_path, 'w', encoding='utf-8') as f:
            f.write(ct)

    # Build appendix XML
    appendix_xml = ''
    for i, db in enumerate(matched_dbs):
        db_codes = [e for e in ecs_codes if ecs_ductbook.get(e) == db and e in stamped_paths]
        if not db_codes:
            continue
        _, rev = DUCTBOOK_INFO.get(db, (db, '01'))
        appendix_xml += (
            f'<w:p><w:pPr><w:pStyle w:val="Heading1"/>'
            f'<w:pageBreakBefore w:val="1"/></w:pPr>'
            f'<w:r><w:t>APPENDIX {i+1} – DUCTWORK BOOK: {db} (Rev {rev})</w:t></w:r></w:p>'
        )
        for ecs in db_codes:
            appendix_xml += _img_para(rId_map[ecs], ecs, stamped_paths[ecs])

    # Inject before </w:body> — sectPr must stay last
    inject_point = xml.rfind('</w:body>')
    xml = xml[:inject_point] + appendix_xml + xml[inject_point:]

    # Ensure sectPr is last child and has header/footer references
    sp_match = re.search(r'<w:sectPr[ >].*?</w:sectPr>', xml, re.DOTALL)
    if sp_match:
        sect_pr = sp_match.group(0)
        xml = xml[:sp_match.start()] + xml[sp_match.end():]
        # Ensure headerReference and footerReference are present
        if 'headerReference' not in sect_pr:
            sect_pr = sect_pr.replace('<w:pgSz',
                '<w:headerReference w:type="default" r:id="rId8"/>\n'
                '  <w:footerReference w:type="default" r:id="rId9"/>\n  <w:pgSz')
        body_end = xml.rfind('</w:body>')
        xml = xml[:body_end] + '\n  ' + sect_pr + '\n' + xml[body_end:]

    with open(doc_xml_path, 'w', encoding='utf-8') as f:
        f.write(xml)

    # ── Repack as docx ───────────────────────────────────────────────────────
    if progress:
        progress("Packing final Word document…")

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(docx_work):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, docx_work)
                zf.write(file_path, arcname)

    return output_path
