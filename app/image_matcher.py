import os
import re
from rapidfuzz import fuzz


class ImageMatcher:
    def __init__(self, image_folder):
        self.image_folder = image_folder

        valid_ext = ['.png', '.jpg', '.jpeg', '.webp']

        self.images = [
            f for f in os.listdir(image_folder)
            if (
                os.path.isfile(os.path.join(image_folder, f))
                and not f.startswith('.')
                and not f.startswith('._')
                and os.path.splitext(f)[1].lower() in valid_ext
            )
        ]

    def normalize(self, text):

        if not text:
            return ''

        text = str(text).lower()

        # 괄호 제거
        text = re.sub(r'\(.*?\)|（.*?）', '', text)

        # 공백류 제거
        text = re.sub(r'[\s　]+', '', text)

        # 특수문자 제거
        text = re.sub(r'[!"#$%&\'()*+,\-./:;<=>?@[\\\]^_`{|}~]', '', text)

        # 일본 특수기호 제거
        replace_list = [
            '「',
            '」',
            '『',
            '』',
            '【',
            '】',
            '～',
            '・',
            '!',
            '！',
        ]

        for r in replace_list:
            text = text.replace(r, '')

        return text

    def find_image(self, jan, name):
        # 1순위 JAN
        for image in self.images:
            if jan and jan in image:
                return image

        normalized_name = self.normalize(name)

        # 2순위 이름 포함
        for image in self.images:
            filename = os.path.splitext(image)[0]
            normalized_image = self.normalize(filename)

            if normalized_name in normalized_image:
                return image

        # 3순위 fuzzy
        best_score = 0
        best_image = None

        for image in self.images:
            score = fuzz.partial_ratio(normalized_name, self.normalize(image))

            if score > best_score:
                best_score = score
                best_image = image

        if best_score >= 55:
            return best_image

        return None