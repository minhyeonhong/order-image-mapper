import os
import re

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

        # normalize 캐싱
        self.normalized_images = {}

        for image in self.images:

            filename = os.path.splitext(image)[0]

            self.normalized_images[image] = self.normalize(filename)

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

        jan = str(jan).strip() if jan else ''

        if not jan:
            return None

        matched = []

        for image in self.images:

            filename = os.path.splitext(image)[0]

            # 파일명 앞 숫자 추출
            match = re.match(r'^(\d+)', filename)

            if not match:
                continue

            image_jan = match.group(1)

            if image_jan == jan:
                matched.append(image)

        if matched:

            matched.sort(key=lambda x: (
                0 if os.path.splitext(x)[0] == jan else 1,
                x
            ))

            return matched[0]

        return None
