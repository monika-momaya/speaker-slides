import copy
import io
from typing import List, Dict, Optional
from PIL import Image
from pptx import Presentation

FIELD_ORDER = [
    'MAIN_SESSION_DETAILS',
    'SPEAKER_SESSION_DETAILS',
    'SPEAKER_NAME_1',
    'SPEAKER_TITLE_1',
    'SPEAKER_COMPANY_1',
    'DATE',
    'HALL_NAME',
]


def _duplicate_slide(prs: Presentation, source_slide):
    blank_slide_layout = prs.slide_layouts[len(prs.slide_layouts) - 1]
    new_slide = prs.slides.add_slide(blank_slide_layout)
    for shp in source_slide.shapes:
        newel = copy.deepcopy(shp.element)
        new_slide.shapes._spTree.insert_element_before(newel, 'p:extLst')
    for key, value in source_slide.part.rels.items():
        if "notesSlide" not in value.reltype:
            try:
                if value.is_external:
                    new_slide.part.rels.add_relationship(value.reltype, value.target_ref, value.rId, value.is_external)
                else:
                    new_slide.part.rels.add_relationship(value.reltype, value.target_part, value.rId)
            except Exception:
                pass
    return new_slide


def _shape_text(shape):
    if not getattr(shape, 'has_text_frame', False):
        return ''
    return '\n'.join(''.join(r.text for r in p.runs) if p.runs else p.text for p in shape.text_frame.paragraphs)


def _replace_text_preserve_shape(shape, new_text: str):
    if not getattr(shape, 'has_text_frame', False):
        return
    tf = shape.text_frame
    if not tf.paragraphs:
        return
    p = tf.paragraphs[0]
    if p.runs:
        p.runs[0].text = new_text
        for r in list(p.runs)[1:]:
            r.text = ''
    else:
        p.text = new_text


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


def _insert_photo_into_placeholder(slide, photo_shape, pil_image: Image.Image):
    left, top, width, height = photo_shape.left, photo_shape.top, photo_shape.width, photo_shape.height
    img = _cover_crop(pil_image, width, height)
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    slide.shapes.add_picture(buf, left, top, width=width, height=height)
    try:
        photo_shape._element.getparent().remove(photo_shape._element)
    except Exception:
        pass


def _find_photo_shape(slide):
    candidates = []
    for shape in slide.shapes:
        if getattr(shape, 'shape_type', None) == 13 and shape.width > 0 and shape.height > 0:
            candidates.append(shape)
    if not candidates:
        return None
    candidates.sort(key=lambda s: (s.width * s.height), reverse=True)
    return candidates[0]


def build_merged_presentation(template_pptx_path_or_file, slide_groups: List[Dict], processed_photo_lookup: Optional[Dict[str, Image.Image]] = None):
    processed_photo_lookup = processed_photo_lookup or {}
    prs = Presentation(template_pptx_path_or_file)
    if len(prs.slides) == 0:
        raise ValueError('Template has no slides.')
    template_slide = prs.slides[0]
    for group in slide_groups:
        new_slide = _duplicate_slide(prs, template_slide)
        speakers = group.get('speakers', [])
        if speakers:
            photo_key = speakers[0].get('photo_key')
            if photo_key and photo_key in processed_photo_lookup:
                photo_shape = _find_photo_shape(new_slide)
                if photo_shape is not None:
                    _insert_photo_into_placeholder(new_slide, photo_shape, processed_photo_lookup[photo_key])
        values = {
            'MAIN_SESSION_DETAILS': str(group.get('MAIN_SESSION_DETAILS', '') or ''),
            'SPEAKER_SESSION_DETAILS': str(group.get('SPEAKER_SESSION_DETAILS', '') or ''),
            'DATE': str(group.get('DATE', '') or ''),
            'HALL_NAME': str(group.get('HALL_NAME', '') or ''),
        }
        if speakers:
            values['SPEAKER_NAME_1'] = str(speakers[0].get('name', '') or '')
            values['SPEAKER_TITLE_1'] = str(speakers[0].get('title', '') or '')
            values['SPEAKER_COMPANY_1'] = str(speakers[0].get('company', '') or '')
        text_shapes = [s for s in new_slide.shapes if getattr(s, 'has_text_frame', False)]
        text_shapes.sort(key=lambda s: (s.top, s.left))
        for shape, key in zip(text_shapes, FIELD_ORDER):
            _replace_text_preserve_shape(shape, values.get(key, ''))
    return prs


def save_presentation_to_bytes(prs: Presentation) -> bytes:
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()
