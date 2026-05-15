import os
import tempfile
import uuid

from PIL import Image as PILImage
from openpyxl.drawing.image import Image


class ImageInserter:

    def __init__(self, worksheet):
        self.ws = worksheet

    def insert_image(
        self,
        image_folder,
        image_name,
        cell,
        width=150,
        height=150,
    ):

        image_path = os.path.join(image_folder, image_name)

        ext = os.path.splitext(image_path)[1].lower()

        # webp → png
        if ext == '.webp':

            converted_path = os.path.join(
                tempfile.gettempdir(),
                f'{uuid.uuid4()}.png'
            )

            pil_img = PILImage.open(image_path)
            pil_img.save(converted_path, 'PNG')

            image_path = converted_path

        # 매번 새 객체 생성
        img = Image(image_path)

        img.width = width
        img.height = height

        row = int(''.join(filter(str.isdigit, cell)))
        col = ''.join(filter(str.isalpha, cell))

        self.ws.row_dimensions[row].height = height * 0.75
        self.ws.column_dimensions[col].width = width / 7

        self.ws.add_image(img, cell)