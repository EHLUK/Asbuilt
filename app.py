"""
As-Built Drawing Compiler — HPC HK2794
Streamlit web app
"""

import os
import tempfile
import shutil
from collections import Counter
from pathlib import Path

import streamlit as st

from compiler import extract_trn, match_drawings, render_and_stamp, build_docx

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="As-Built Compiler — HK2794",
    page_icon="📐",
    layout="centered",
)

st.title("📐 As-Built Drawing Compiler")
st.caption("HPC HK2794 · Exentec Hargreaves · Ductwork")

st.markdown("---")

# ── File uploads ──────────────────────────────────────────────────────────────
st.subheader("1 · Upload files")

col1, col2 = st.columns(2)

with col1:
    trn_file = st.file_uploader(
        "TRN PDF",
        type=["pdf"],
        help="Technical Release Note — filename should contain 'TRN' and a TC reference number",
    )
    template_file = st.file_uploader(
        "As-Built Word template (.docx)",
        type=["docx"],
        help="E21369-EHL-XX-ZZ-RP-MM-000xxx.docx",
    )

with col2:
    drawing_files = st.file_uploader(
        "Ductwork Drawing PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        help="One or more drawing books e.g. ENG-GSC-WS08-02393.pdf",
    )
    stamp_file = st.file_uploader(
        "Conformance Stamp PNG",
        type=["png"],
        help="Portrait red-border stamp image",
    )

st.markdown("---")

# ── Run button ────────────────────────────────────────────────────────────────
st.subheader("2 · Compile")

all_ready = trn_file and template_file and drawing_files and stamp_file

if not all_ready:
    missing = []
    if not trn_file:        missing.append("TRN PDF")
    if not template_file:   missing.append("Word template")
    if not drawing_files:   missing.append("at least one drawing PDF")
    if not stamp_file:      missing.append("conformance stamp PNG")
    st.info(f"Still needed: {', '.join(missing)}")

run_btn = st.button("▶ Build As-Built Document", disabled=not all_ready, type="primary")

if run_btn and all_ready:
    progress_box = st.empty()
    status_log   = st.empty()
    log_lines    = []

    def progress(msg):
        progress_box.info(f"⏳ {msg}")
        log_lines.append(msg)
        status_log.code('\n'.join(log_lines[-8:]))

    with tempfile.TemporaryDirectory() as tmp:
        try:
            # ── Save uploads to temp dir ──────────────────────────────────────
            trn_path      = os.path.join(tmp, "trn.pdf")
            template_path = os.path.join(tmp, "template.docx")
            stamp_path    = os.path.join(tmp, "stamp.png")

            with open(trn_path,      "wb") as f: f.write(trn_file.read())
            with open(template_path, "wb") as f: f.write(template_file.read())
            with open(stamp_path,    "wb") as f: f.write(stamp_file.read())

            drawing_paths = []
            for df in drawing_files:
                dp = os.path.join(tmp, df.name)
                with open(dp, "wb") as f: f.write(df.read())
                drawing_paths.append(dp)

            # ── Step 1: Extract TRN ───────────────────────────────────────────
            progress("Extracting ECS codes from TRN…")
            trn_data = extract_trn(trn_path, progress=progress)
            ecs_codes    = trn_data["ecs_codes"]
            ecs_ductbook = trn_data["ecs_ductbook"]
            delivery_ref = trn_data["delivery_ref"]
            tc_ref       = trn_data["tc_ref"]

            if not ecs_codes:
                st.error("❌ No ECS codes found in the TRN PDF. Check the file and try again.")
                st.stop()

            db_counts = Counter(ecs_ductbook.values())
            progress(f"Found {len(ecs_codes)} ECS codes across {len(db_counts)} ductbook(s)")

            # ── Step 2: Match drawings ────────────────────────────────────────
            matches, not_found = match_drawings(
                ecs_codes, ecs_ductbook, drawing_paths, progress=progress
            )
            progress(f"Matched {len(matches)}/{len(ecs_codes)} drawings")

            # ── Step 3+4: Render + stamp ──────────────────────────────────────
            stamped_paths = render_and_stamp(matches, stamp_path, tmp, progress=progress)
            progress(f"Stamped {len(stamped_paths)} drawings")

            # ── Step 5: Build Word doc ────────────────────────────────────────
            safe_ref = (delivery_ref or "AsBuilt").replace(" ", "_").replace(",", "").replace("&", "and")
            output_path = os.path.join(tmp, f"AsBuilt_{safe_ref}.docx")
            build_docx(template_path, trn_data, matches, stamped_paths,
                       output_path, tmp, progress=progress)

            progress("Done! ✅")
            progress_box.success("✅ As-Built document compiled successfully")
            status_log.empty()

            # ── Results summary ───────────────────────────────────────────────
            st.markdown("---")
            st.subheader("3 · Results")

            col_a, col_b, col_c = st.columns(3)
            col_a.metric("ECS codes in TRN", len(ecs_codes))
            col_b.metric("Drawings matched", len(matches))
            col_c.metric("Drawings embedded", len(stamped_paths))

            if not_found:
                missing_db = Counter(ecs_ductbook.get(c, "unknown") for c in not_found)
                st.warning(
                    f"⚠️ {len(not_found)} drawings not found — "
                    f"upload these PDFs to get the rest:\n"
                    + "\n".join(f"• `{db}.pdf` ({cnt} drawings)" for db, cnt in sorted(missing_db.items()))
                )

            matched_dbs = sorted(set(ecs_ductbook[e] for e in stamped_paths))
            if matched_dbs:
                st.markdown("**Appendices built:**")
                for i, db in enumerate(matched_dbs):
                    db_codes = [e for e in ecs_codes if ecs_ductbook.get(e) == db and e in stamped_paths]
                    st.markdown(f"- Appendix {i+1}: `{db}` — {len(db_codes)} drawings")

            # ── Download ──────────────────────────────────────────────────────
            with open(output_path, "rb") as f:
                docx_bytes = f.read()

            filename = f"AsBuilt_{safe_ref}.docx"
            st.download_button(
                label="⬇️ Download As-Built Document",
                data=docx_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
            )

        except Exception as e:
            progress_box.error(f"❌ Error: {e}")
            st.exception(e)

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption("Exentec Hargreaves · HPC HK2794 · As-Built Compiler")
