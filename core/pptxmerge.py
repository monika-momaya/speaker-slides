import copy
import io
import re
from typing import List, Dict, Optional
from PIL import Image
from pptx import Presentation

TAG_RE = re.compile(r'<<\s*([A-Z0-9_]+)\s*>>')


def _duplicate_slide(prs: Presentation, source_slide):
    blank_layout = prs.slide_layouts[len(prs.slide_layouts) - 1]
    new_slide = prs.slides.add_slide(blank_layout)
    for shp in source_slide.shapes:
        newel = copy.deepcopy(shp.element)
        new_slide.shapes._spTree.insert_element_before(newel, 'p:extLst')
    return new_slide


def _shape_text(shape):
    if not getattr(shape, 'has_text_frame', False):
        return ''
    return '\n'.join(''.join(r.text for r in p.runs) if p.runs else p.text for p in shape.text_frame.paragraphs)


def _replace_tag_text_preserve_runs(shape, values: Dict[str, str]):
    if not getattr(shape, 'has_text_frame', False):
        return False
    changed = False
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            txt = run.text or ''
            if '<<' not in txt:
                continue
            def repl(m):
                key = m.group(1)
                return str(values.get(key, m.group(0)) or '')
            new_txt = TAG_RE.sub(repl, txt)
            if new_txt != txt:
                run.text = new_txt
                changed = True
    return changed


def _find_photo_placeholder(slide, tag='SPEAKER_PHOTO_1'):
    pattern = re.compile(r'<<\s*' + re.escape(tag) + r'\s*>>')
    for shape in slide.shapes:
        if getattr(shape, 'has_text_frame', False):
            if pattern.search(_shape_text(shape)):
                return shape
    return None


def _cover_crop(pil_image: Image.Image, width: int, height: int):
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
    return img


def _replace_placeholder_with_photo(slide, placeholder_shape, pil_image: Image.Image):
    left, top, width, height = placeholder_shape.left, placeholder_shape.top, placeholder_shape.width, placeholder_shape.height
    img = _cover_crop(pil_image, width, height)
    buf = io.BytesIO()
    img.save(buf, format='PNG')
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
        raise ValueError('Template has no slides.')
    template_slide = prs.slides[0]
    for group in slide_groups:
        new_slide = _duplicate_slide(prs, template_slide)
        speakers = group.get('speakers', [])
        values = {
            'SESSION_NAME': str(group.get('SESSION_NAME', '') or ''),
            'HALL_NAME': str(group.get('HALL_NAME', '') or ''),
            'DATE': str(group.get('DATE', '') or ''),
            'MAIN_SESSION_DETAILS': str(group.get('MAIN_SESSION_DETAILS', '') or ''),
            'SPEAKER_SESSION_DETAILS': str(group.get('SPEAKER_SESSION_DETAILS', '') or ''),
            'PLACEHOLDER_1': str(group.get('PLACEHOLDER_1', '') or ''),
            'PLACEHOLDER_2': str(group.get('PLACEHOLDER_2', '') or ''),
        }
        for idx, sp in enumerate(speakers, start=1):
            values[f'SPEAKER_NAME_{idx}'] = str(sp.get('name', '') or '')
            values[f'SPEAKER_TITLE_{idx}'] = str(sp.get('title', '') or '')
            values[f'SPEAKER_COMPANY_{idx}'] = str(sp.get('company', '') or '')
        for shape in list(new_slide.shapes):
            _replace_tag_text_preserve_runs(shape, values)
        if speakers:
            photo_key = speakers[0].get('photo_key')
            if photo_key and photo_key in processed_photo_lookup:
                photo_placeholder = _find_photo_placeholder(new_slide, 'SPEAKER_PHOTO_1')
                if photo_placeholder is not None:
                    _replace_placeholder_with_photo(new_slide, photo_placeholder, processed_photo_lookup[photo_key])
    return prs


def save_presentation_to_bytes(prs: Presentation) -> bytes:
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()
