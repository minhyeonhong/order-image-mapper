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
        text = text.lower()
        text = re.sub(r'[^\w가-힣ぁ-んァ-ン一-龥]', '', text)
        return text

    def find_image(self, jan, name):
        # 1순위 JAN
        for image in self.images:
            if jan and jan in image:
                return image

        normalized_name = self.normalize(name)

        # 2순위 이름 포함
        for image in self.images:
            normalized_image = self.normalize(image)

            if normalized_name in normalized_image:
                return image

        # 3순위 fuzzy
        best_score = 0
        best_image = None

        for image in self.images:
            score = fuzz.ratio(normalized_name, self.normalize(image))

            if score > best_score:
                best_score = score
                best_image = image

        if best_score >= 70:
            return best_image

        return None