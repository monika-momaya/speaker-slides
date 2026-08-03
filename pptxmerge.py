import copy
import io
from typing import List, Dict, Optional
from PIL import Image
from pptx import Presentation

FIELD_MAP = {
    'SESSION_NAME': 'SESSION_NAME',
    'HALL_NAME': 'HALL_NAME',
    'DATE': 'DATE',
    'MAIN_SESSION_DETAILS': 'MAIN_SESSION_DETAILS',
    'SPEAKER_SESSION_DETAILS': 'SPEAKER_SESSION_DETAILS',
    'PLACEHOLDER_1': 'PLACEHOLDER_1',
    'PLACEHOLDER_2': 'PLACEHOLDER_2',
}


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


def _shape_text(shape):
    if not getattr(shape, 'has_text_frame', False):
        return ''
    return '\n'.join(''.join(r.text for r in p.runs) if p.runs else p.text for p in shape.text_frame.paragraphs)


def _replace_entire_shape_text_preserve_first_run(shape, new_text: str):
    if not getattr(shape, 'has_text_frame', False):
        return False
    tf = shape.text_frame
    if not tf.paragraphs:
        return False
    p = tf.paragraphs[0]
    if p.runs:
        p.runs[0].text = new_text
        for extra in list(p.runs)[1:]:
            extra.text = ''
    else:
        p.text = new_text
    for extra_p in list(tf.paragraphs)[1:]:
        for r in extra_p.runs:
            r.text = ''
        if not extra_p.runs:
            extra_p.text = ''
    return True


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


def _replace_picture_in_picture_shape(slide, picture_shape, pil_image: Image.Image):
    left, top, width, height = picture_shape.left, picture_shape.top, picture_shape.width, picture_shape.height
    img = _cover_crop(pil_image, width, height)
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    slide.shapes.add_picture(buf, left, top, width=width, height=height)
    try:
        picture_shape._element.getparent().remove(picture_shape._element)
    except Exception:
        pass


def _picture_candidates(slide):
    pics = []
    for shape in slide.shapes:
        if getattr(shape, 'shape_type', None) == 13:
            pics.append(shape)
    return pics


def _largest_picture_near_center(slide):
    pics = _picture_candidates(slide)
    if not pics:
        return None
    slide_w = slide.part.slide_layout.part.package.presentation_part.presentation.slide_width
    slide_h = slide.part.slide_layout.part.package.presentation_part.presentation.slide_height
    cx = slide_w / 2
    cy = slide_h / 2
    def score(s):
        area = s.width * s.height
        scx = s.left + s.width / 2
        scy = s.top + s.height / 2
        dist = abs(scx - cx) + abs(scy - cy)
        return (area, -dist)
    pics.sort(key=score, reverse=True)
    return pics[0]


def _replace_text_shapes_by_order(slide, tag_values: Dict[str, str], speakers: List[Dict]):
    text_shapes = [s for s in slide.shapes if getattr(s, 'has_text_frame', False)]
    text_shapes.sort(key=lambda s: (s.top, s.left))
    ordered_values = []
    ordered_values.append(tag_values.get('MAIN_SESSION_DETAILS', ''))
    ordered_values.append(tag_values.get('SPEAKER_SESSION_DETAILS', ''))
    for idx, sp in enumerate(speakers, start=1):
        ordered_values.extend([
            sp.get('name', ''),
            sp.get('title', ''),
            sp.get('company', ''),
        ])
    ordered_values.append(tag_values.get('DATE', ''))
    ordered_values.append(tag_values.get('HALL_NAME', ''))
    targets = []
    for s in text_shapes:
        txt = _shape_text(s).strip()
        if txt:
            targets.append(s)
    for shape, value in zip(targets, ordered_values):
        _replace_entire_shape_text_preserve_first_run(shape, str(value or ''))


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
                pic_shape = _largest_picture_near_center(new_slide)
                if pic_shape is not None:
                    _replace_picture_in_picture_shape(new_slide, pic_shape, processed_photo_lookup[photo_key])
        tag_values = {k: str(group.get(k, '') or '') for k in FIELD_MAP.keys()}
        _replace_text_shapes_by_order(new_slide, tag_values, speakers)
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    if slides:
        xml_slides.remove(slides[0])
    return prs


def save_presentation_to_bytes(prs: Presentation) -> bytes:
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()
