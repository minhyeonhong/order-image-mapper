import os

from PIL import Image as PILImage
from openpyxl.drawing.image import Image


class ImageInserter:
    def __init__(self, worksheet):
        self.ws = worksheet

    def insert_image(self, image_folder, image_name, cell):

        image_path = os.path.join(image_folder, image_name)

        ext = os.path.splitext(image_path)[1].lower()

        # webp 변환
        if ext == '.webp':

            converted_path = image_path.replace('.webp', '.png')

            pil_img = PILImage.open(image_path)
            pil_img.save(converted_path, 'PNG')

            image_path = converted_path

        img = Image(image_path)

        img.width = 60
        img.height = 60

        row = int(cell[1:])

        # 행 높이 조절
        self.ws.row_dimensions[row].height = 50

        self.ws.add_image(img, cell)