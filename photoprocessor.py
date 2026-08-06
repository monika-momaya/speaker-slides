from dataclasses import dataclass
from PIL import Image, ImageOps

@dataclass
class PhotoProcessResult:
    image: Image.Image
    facedetected: bool
    note: str

def processphoto(rawimg: Image.Image) -> PhotoProcessResult:
    img = ImageOps.exif_transpose(rawimg).convert("RGB")
    return PhotoProcessResult(image=img, facedetected=True, note="Processed")
