import io
import os
import streamlit as st
from PIL import Image

from core.pptxtagparser import extract_tags_from_pptx, TemplateTagReport
from core.exceltemplate import build_excel_template_bytes
from core.mergesheet import read_merge_sheet
from core.pptxmerge import build_merged_presentation, save_presentation_to_bytes
from core.photoprocessor import processphoto
from core.namematcher import matchphotostospeakers

APP_DIR = os.path.dirname(os.path.abspath(__file__))

st.set_page_config(page_title="Soft Slide Generator", layout="wide")

with st.sidebar:
    st.header("Template tags")
    st.markdown("""
Type these exact tags inside textboxes in your PPTX template:

- <<SESSION_NAME>>
- <<HALL_NAME>>
- <<DATE>>
- <<MAIN_SESSION_DETAILS>>
- <<SPEAKER_SESSION_DETAILS>>
- <<PLACEHOLDER_1>>
- <<PLACEHOLDER_2>>
- <<SPEAKER_NAME_1>>
- <<SPEAKER_TITLE_1>>
- <<SPEAKER_COMPANY_1>>
- <<SPEAKER_PHOTO_1>>
""")

if "tag_report" not in st.session_state:
    st.session_state.tag_report = None
if "template_bytes" not in st.session_state:
    st.session_state.template_bytes = None
if "final_pptx_bytes" not in st.session_state:
    st.session_state.final_pptx_bytes = None

st.title("Conference Soft Slide Generator")
st.caption("Upload a tagged PPTX template, fill Excel, and generate a multi-slide deck.")

st.header("1. Upload Tagged PPTX Template")
template_pptx_file = st.file_uploader("Template .pptx", type="pptx")

slide_groups = None
processed_photo_lookup = {}

if template_pptx_file is not None:
    template_bytes = template_pptx_file.getvalue()
    st.session_state.template_bytes = template_bytes
    try:
        report = extract_tags_from_pptx(io.BytesIO(template_bytes))
        st.session_state.tag_report = report
        st.success(f"Template parsed. Detected {report.max_speaker_slot} speaker slot(s).")
        st.markdown("**Found:** " + (", ".join(f"`<<{t}>>`" for t in report.found_single_tags) or "none"))
        st.markdown("**Missing (optional):** " + (", ".join(f"`<<{t}>>`" for t in report.missing_single_tags) or "none"))
        for w in report.warnings:
            st.warning(w)
    except Exception as e:
        st.error(f"Could not parse template: {e}")
        st.session_state.tag_report = None

st.divider()
st.header("2. Download Data Template, Fill It, Upload It Back")

if st.session_state.tag_report is None:
    st.info("Upload a template first to generate the matching Excel sheet.")
else:
    excel_bytes = build_excel_template_bytes(st.session_state.tag_report)
    st.download_button(
        "Download Excel data template",
        data=excel_bytes,
        file_name="slide_data_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.caption("Use the same Require Single Slide label, e.g. Panel 1, Panel 2, to group rows onto one slide.")
    merge_excel_file = st.file_uploader("Upload filled data sheet (.xlsx)", type="xlsx", key="mergeexcel")
    photo_files = st.file_uploader(
        "Upload speaker photos",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
        key="mergephotos",
    )

    if merge_excel_file is not None:
        try:
            slide_groups, overflow_warnings = read_merge_sheet(
                merge_excel_file, st.session_state.tag_report.max_speaker_slot
            )
            st.success(f"Loaded {len(slide_groups)} slide(s).")
            for w in overflow_warnings:
                st.warning(w)
        except Exception as e:
            st.error(f"Could not read data sheet: {e}")
            slide_groups = None

    if slide_groups and photo_files:
        try:
            photo_filenames = [pf.name for pf in photo_files]
            photo_file_lookup = {pf.name: pf for pf in photo_files}
            wanted_names = []
            for g in slide_groups:
                for sp in g["speakers"]:
                    if sp.get("photo_key"):
                        wanted_names.append(str(sp["photo_key"]))
            match_results = matchphotostospeakers(wanted_names, photo_filenames)
            wanted_to_chosen = {}
            idx = 0
            for g in slide_groups:
                for sp in g["speakers"]:
                    if sp.get("photo_key"):
                        wanted_to_chosen[str(sp["photo_key"])] = match_results[idx].matchedfilename
                        idx += 1
            for wanted_key, matched_filename in wanted_to_chosen.items():
                if matched_filename and matched_filename in photo_file_lookup:
                    raw_img = Image.open(photo_file_lookup[matched_filename])
                    result = processphoto(raw_img)
                    processed_photo_lookup[wanted_key] = result.image
        except Exception as e:
            st.error(f"Photo processing failed: {e}")

st.divider()
st.header("3. Generate Final Deck")

ready = st.session_state.template_bytes is not None and slide_groups is not None and len(slide_groups) > 0
if not ready:
    st.info("Complete the template and data sheet steps above to generate the deck.")

if st.button("Generate Deck", type="primary", disabled=not ready):
    try:
        prs = build_merged_presentation(
            io.BytesIO(st.session_state.template_bytes),
            slide_groups,
            processed_photo_lookup=processed_photo_lookup,
        )
        st.session_state.final_pptx_bytes = save_presentation_to_bytes(prs)
        st.success(f"Generated {len(slide_groups)} slide(s).")
    except Exception as e:
        st.error(f"Could not generate deck: {e}")

if st.session_state.final_pptx_bytes is not None:
    st.download_button(
        "Download Generated PPTX",
        data=st.session_state.final_pptx_bytes,
        file_name="generated_slides.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
