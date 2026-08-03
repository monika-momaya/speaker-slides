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


def _replace_text_in_runs_preserve_style(shape, tag_values: Dict[str, str], skip_tags=None):
    if not getattr(shape, 'has_text_frame', False):
        return
    skip_tags = set(skip_tags or [])
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            text = run.text or ''
            if '<<' not in text or '>>' not in text:
                continue
            def repl(m):
                key = m.group(1)
                if key in skip_tags:
                    return m.group(0)
                return tag_values.get(key, '')
            run.text = TAG_PATTERN.sub(repl, text)


def _shape_full_text(shape):
    if not getattr(shape, 'has_text_frame', False):
        return ''
    out = []
    for p in shape.text_frame.paragraphs:
        out.append(''.join(r.text for r in p.runs) if p.runs else p.text)
    return '\n'.join(out)


def _find_photo_shape(slide, tag_name: str):
    pattern = re.compile(r"<<\s*" + re.escape(tag_name) + r"\s*>>")
    for shape in slide.shapes:
        if pattern.search(_shape_full_text(shape)):
            return shape
    return None


def _insert_picture_cover(slide, left, top, width, height, pil_image: Image.Image):
    img = pil_image.convert('RGB')
    target_ratio = width / height
    img_ratio = img.width / img.height if img.height else target_ratio
    if img_ratio > target_ratio:
        new_w = int(img.height * target_ratio)
        x0 = max((img.width - new_w) // 2, 0)
        img = img.crop((x0, 0, x0 + new_w, img.height))
    else:
        new_h = int(img.width / target_ratio) if target_ratio else img.height
        y0 = max((img.height - new_h) // 2, 0)
        img = img.crop((0, y0, img.width, y0 + new_h))
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    slide.shapes.add_picture(buf, left, top, width=width, height=height)


def _replace_photo_placeholders(slide, speakers, processed_photo_lookup):
    inserted = []
    for idx, sp in enumerate(speakers, start=1):
        photo_key = sp.get('photo_key')
        if not photo_key or photo_key not in processed_photo_lookup:
            continue
        tag_name = f'SPEAKER_PHOTO_{idx}'
        placeholder_shape = _find_photo_shape(slide, tag_name)
        if placeholder_shape is None:
            continue
        left, top, width, height = placeholder_shape.left, placeholder_shape.top, placeholder_shape.width, placeholder_shape.height
        _insert_picture_cover(slide, left, top, width, height, processed_photo_lookup[photo_key])
        try:
            placeholder_shape._element.getparent().remove(placeholder_shape._element)
        except Exception:
            pass
        inserted.append(tag_name)
    return inserted


def build_merged_presentation(template_pptx_path_or_file, slide_groups: List[Dict], processed_photo_lookup: Optional[Dict[str, Image.Image]] = None):
    processed_photo_lookup = processed_photo_lookup or {}
    prs = Presentation(template_pptx_path_or_file)
    if len(prs.slides) == 0:
        raise ValueError('Template has no slides.')
    template_slide = prs.slides[0]
    for group in slide_groups:
        new_slide = _duplicate_slide(prs, template_slide)
        speakers = group.get('speakers', [])
        _replace_photo_placeholders(new_slide, speakers, processed_photo_lookup)
        tag_values = {k: str(group.get(k, '') or '') for k in ['SESSION_NAME', 'HALL_NAME', 'DATE', 'MAIN_SESSION_DETAILS', 'SPEAKER_SESSION_DETAILS', 'PLACEHOLDER_1', 'PLACEHOLDER_2']}
        for idx, sp in enumerate(speakers, start=1):
            tag_values[f'SPEAKER_NAME_{idx}'] = str(sp.get('name', '') or '')
            tag_values[f'SPEAKER_TITLE_{idx}'] = str(sp.get('title', '') or '')
            tag_values[f'SPEAKER_COMPANY_{idx}'] = str(sp.get('company', '') or '')
        skip_tags = {f'SPEAKER_PHOTO_{i}' for i in range(1, max(1, len(speakers)) + 5)}
        for shape in list(new_slide.shapes):
            _replace_text_in_runs_preserve_style(shape, tag_values, skip_tags=skip_tags)
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    if slides:
        xml_slides.remove(slides[0])
    return prs


def save_presentation_to_bytes(prs: Presentation) -> bytes:
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()
