import copy
import io
import re
from typing import List, Dict, Optional
from PIL import Image
from pptx import Presentation
from core.pptxtagparser import TAG_PATTERN


def _duplicate_slide(prs: Presentation, source_slide):
    new_slide = prs.slides.add_slide(source_slide.slide_layout)
    for shape in list(new_slide.shapes):
        try:
            shape._element.getparent().remove(shape._element)
        except Exception:
            pass
    for shape in source_slide.shapes:
        new_el = copy.deepcopy(shape._element)
        new_slide.shapes._spTree.append(new_el)
    return new_slide


def _replace_text_in_shape(shape, tag_values: Dict[str, str]):
    if not getattr(shape, "has_text_frame", False):
        return
    full_text = "\n".join(p.text for p in shape.text_frame.paragraphs)
    if "<<" not in full_text or ">>" not in full_text:
        return
    new_text = TAG_PATTERN.sub(lambda m: tag_values.get(m.group(1), ""), full_text)
    shape.text_frame.clear()
    parts = new_text.split("\n")
    for i, part in enumerate(parts):
        p = shape.text_frame.paragraphs[0] if i == 0 else shape.text_frame.add_paragraph()
        p.text = part


def _find_photo_shape(slide, tag_name: str):
    pattern = re.compile(r"<<\s*" + re.escape(tag_name) + r"\s*>>")
    for shape in slide.shapes:
        if getattr(shape, "has_text_frame", False):
            full_text = "\n".join(p.text for p in shape.text_frame.paragraphs)
            if pattern.search(full_text):
                return shape
    return None


def _insert_photo_into_shape(slide, placeholder_shape, pil_image: Image.Image):
    left, top, width, height = placeholder_shape.left, placeholder_shape.top, placeholder_shape.width, placeholder_shape.height
    buf = io.BytesIO()
    pil_image.convert("RGB").save(buf, format="PNG")
    buf.seek(0)
    slide.shapes.add_picture(buf, left, top, width=width, height=height)
    try:
        placeholder_shape._element.getparent().remove(placeholder_shape._element)
    except Exception:
        pass


def build_merged_presentation(template_pptx_path_or_file, slide_groups: List[Dict], processed_photo_lookup: Optional[Dict[str, Image.Image]] = None):
    processed_photo_lookup = processed_photo_lookup or {}
    prs = Presentation(template_pptx_path_or_file)
    if len(prs.slides) == 0:
        raise ValueError("Template has no slides.")
    template_slide = prs.slides[0]
    for group in slide_groups:
        new_slide = _duplicate_slide(prs, template_slide)
        tag_values = {k: str(group.get(k, "") or "") for k in ["SESSION_NAME", "HALL_NAME", "DATE", "MAIN_SESSION_DETAILS", "SPEAKER_SESSION_DETAILS", "PLACEHOLDER_1", "PLACEHOLDER_2"]}
        speakers = group.get("speakers", [])
        for idx, sp in enumerate(speakers, start=1):
            tag_values[f"SPEAKER_NAME_{idx}"] = str(sp.get("name", "") or "")
            tag_values[f"SPEAKER_TITLE_{idx}"] = str(sp.get("title", "") or "")
            tag_values[f"SPEAKER_COMPANY_{idx}"] = str(sp.get("company", "") or "")
        for shape in list(new_slide.shapes):
            _replace_text_in_shape(shape, tag_values)
        for idx, sp in enumerate(speakers, start=1):
            photo_key = sp.get("photo_key")
            if not photo_key or photo_key not in processed_photo_lookup:
                continue
            placeholder_shape = _find_photo_shape(new_slide, f"SPEAKER_PHOTO_{idx}")
            if placeholder_shape is not None:
                _insert_photo_into_shape(new_slide, placeholder_shape, processed_photo_lookup[photo_key])
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    if slides:
        xml_slides.remove(slides[0])
    return prs


def save_presentation_to_bytes(prs: Presentation) -> bytes:
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()
