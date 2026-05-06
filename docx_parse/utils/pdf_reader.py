import base64
from io import BytesIO

from PIL import Image


def image_to_bytes(image: Image.Image, image_format: str = "JPEG") -> bytes:
    with BytesIO() as image_buffer:
        image.save(image_buffer, format=image_format)
        return image_buffer.getvalue()


def image_to_b64str(image: Image.Image, image_format: str = "JPEG") -> str:
    image_bytes = image_to_bytes(image, image_format)
    return f"data:image/{image_format.lower()};base64,{base64.b64encode(image_bytes).decode('utf-8')}"
