from pptx import Presentation
from pptx.util import Pt
from pptx.util import Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN


class ChangeLyrics:
    def __init__(self, text_file_path, save_path, ppt_title, ppt_path=None):
        self.text_file_path = text_file_path
        self.ppt_path = ppt_path
        self.save_path = save_path
        self.ppt_title = ppt_title

    def read_text_file(self):
        # 텍스트 파일 열기
        with open(self.text_file_path, 'r', encoding='utf-8') as f:
            text = f.read()
        return text

    def sep_lyrics(self, text, i, j):
        # 텍스트를 줄 단위로 분리하고 빈 줄 제거
        lines = [line.strip() for line in text.strip().split('\n') if line.strip()]

        # 한글과 영어 라인 분리 (홀수 인덱스: 한글, 짝수 인덱스: 영어)
        lines_pair = lines[i::j]
        return lines_pair

    # 두 줄씩 그룹화
    def group_pairs(self, lines):
        return [lines[i:i + 2] for i in range(0, len(lines), 2)]

    def setting_slide_size_16_9(self, prs):
        prs.slide_width = Inches(13.3333)  # 약 12288000 EMU
        prs.slide_height = Inches(7.5)

    def setting_slide_layout(self, prs):
        # 슬라이드 레이아웃 선택(빈 슬라이드)
        slide_layout = prs.slide_layouts[6]
        background = slide_layout.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGB(0,0,0)
        return slide_layout

    def create_text_box(self, left, top, width, height):
        left = Inches(float(left))
        top = Inches(float(top))
        width = Inches(float(width))
        height = Inches(float(height))

        # txBox = slide.shapes.add_textbox(left, top, width, height)
        return left, top, width, height

    def text_exchange(self, pair, slide_index):
        return "\n".join(pair[slide_index])

    def input_text(self, pair):
        return "\n".join(pair)

    def font_settings(self, paragraph, font, size):
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.name = font  # 글씨체
            run.font.size = Pt(int(size))  # 글자 크기
            run.font.bold = True  # 굵게 설정
            run.font.color.rgb = RGBColor(255, 255, 255)  # 글자 색

    def save(self, prs):
        save_path = f"{self.save_path}/{self.ppt_title}.pptx"
        prs.save(save_path)
        return print("파일 저장됨.")

    def change_lyrics(self):
        text = self.read_text_file()

        korean_lines = self.sep_lyrics(text, 0, 2)
        english_lines = self.sep_lyrics(text, 1, 2)
        korean_pairs = self.group_pairs(korean_lines)
        english_pairs = self.group_pairs(english_lines)

        # PPT 파일 열기
        prs = Presentation(self.ppt_path)
        self.setting_slide_size_16_9(prs)

        # 슬라이드별로 텍스트 상자를 순회하며 처리
        for slide_index, slide in enumerate(prs.slides):
            if slide_index >= len(korean_pairs) or slide_index >= len(english_pairs):
                print(f"슬라이드 {slide_index + 1}에 텍스트가 부족합니다.")
                continue

            # 현재 슬라이드의 텍스트 상자 저장
            text_boxes = [shape.text_frame for shape in slide.shapes if shape.has_text_frame]

            # 텍스트 교환
            korean_text = self.text_exchange(korean_pairs, slide_index)
            english_text = self.text_exchange(english_pairs, slide_index)
            text_boxes[0].text = korean_text
            text_boxes[1].text = english_text
            # 텍스트 상자별로 처리
            for text_box_index, text_frame in enumerate(text_boxes):
                # 첫 번째 텍스트 상자 처리
                if text_box_index == 0:
                    for paragraph in text_frame.paragraphs:
                        self.font_settings(paragraph, "나눔고딕", "40")
                else:
                    for paragraph in text_frame.paragraphs:
                        self.font_settings(paragraph, "나눔고딕", "15")

        self.save(prs)

    def change_only_korean_lyrics(self):
        text = self.read_text_file()

        # 텍스트를 줄 단위로 분리하고 빈 줄(공백) 제거
        korean_lines = [line.strip() for line in text.strip().split('\n') if line.strip()]
        # 텍스트를 두개씩 하나의 그룹으로 형성
        korean_pairs = self.group_pairs(korean_lines)

        # PPT 생성
        prs = Presentation(self.ppt_path)
        self.setting_slide_size_16_9(prs)

        # 슬라이드별로 텍스트 상자를 순회하며 처리
        for slide_index, slide in enumerate(prs.slides):
            if slide_index >= len(korean_pairs):
                print(f"슬라이드 {slide_index + 1}에 텍스트가 부족합니다.")
                continue

            # 현재 슬라이드의 텍스트 상자 저장
            text_boxes = [shape.text_frame for shape in slide.shapes if shape.has_text_frame]

            # 텍스트 교환
            korean_text = self.text_exchange(korean_pairs, slide_index)
            text_boxes[0].text = korean_text

            # 텍스트 설정
            for text_box_index, text_frame in enumerate(text_boxes):
                if text_box_index == 0:
                    for paragraph in text_frame.paragraphs:
                        self.font_settings(paragraph, "나눔고딕", "40")

        self.save(prs)

    def change_only_english_lyrics(self):
        text = self.read_text_file()

        # 텍스트를 줄 단위로 분리하고 빈 줄(공백) 제거
        english_lines = [line.strip() for line in text.strip().split('\n') if line.strip()]
        # 텍스트를 두개씩 하나의 그룹으로 형성
        english_pairs = self.group_pairs(english_lines)

        # PPT 생성
        prs = Presentation(self.ppt_path)
        self.setting_slide_size_16_9(prs)

        # 슬라이드별로 텍스트 상자를 순회하며 처리
        for slide_index, slide in enumerate(prs.slides):
            if slide_index >= len(english_pairs):
                print(f"슬라이드 {slide_index + 1}에 텍스트가 부족합니다.")
                continue

            # 현재 슬라이드의 텍스트 상자 저장
            text_boxes = [shape.text_frame for shape in slide.shapes if shape.has_text_frame]

            # 텍스트 교환
            english_text = self.text_exchange(english_pairs, slide_index)
            text_boxes[1].text = english_text

            # 텍스트 설정
            for text_box_index, text_frame in enumerate(text_boxes):
                if text_box_index == 1:
                    for paragraph in text_frame.paragraphs:
                        self.font_settings(paragraph, "나눔고딕", "15")

        self.save(prs)

    def create_lyrics(self):
        text = self.read_text_file()

        # 한글과 영어 라인 분리 (홀수 인덱스: 한글, 짝수 인덱스: 영어)
        korean_lines = self.sep_lyrics(text, 0, 2)
        english_lines = self.sep_lyrics(text, 1, 2)
        korean_pairs = self.group_pairs(korean_lines)
        english_pairs = self.group_pairs(english_lines)
        # PPT 생성
        prs = Presentation()
        self.setting_slide_size_16_9(prs)
        slide_layout = self.setting_slide_layout(prs)

        # 슬라이드별로 텍스트 상자를 순회하며 처리
        for slide_index, (ko_pair, en_pair) in enumerate(zip(korean_pairs, english_pairs), start=1):

            # 빈 슬라이드 생성
            slide = prs.slides.add_slide(slide_layout)

            # 슬라이드 배경을 검은색으로 설정
            background = slide.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(0, 0, 0)

            # 텍스트 박스 1: 한글 텍스트
            left_ko, top_ko, width_ko, height_ko = self.create_text_box(1.6667, 2.55, 10, 1.476)
            # 텍스트 박스 2: 영어 텍스트
            left_en, top_en, width_en, height_en = self.create_text_box(1.6667, 4.572, 10, 0.933)

            # # 슬라이드에 텍스트 박스 추가
            txBox_ko = slide.shapes.add_textbox(left_ko, top_ko, width_ko, height_ko)
            txBox_en = slide.shapes.add_textbox(left_en, top_en, width_en, height_en)

            # 텍스트 입력
            korean_text = self.input_text(ko_pair)
            english_text = self.input_text(en_pair)
            txBox_ko.text = korean_text
            txBox_en.text = english_text

            # 텍스트 설정
            for paragraph in txBox_ko.text_frame.paragraphs:
                self.font_settings(paragraph, "나눔고딕", "40")

            for paragraph in txBox_en.text_frame.paragraphs:
                self.font_settings(paragraph, "나눔고딕", "15")

        self.save(prs)

    def create_only_korean_lyrics(self):
        text = self.read_text_file()

        # 텍스트를 줄 단위로 분리하고 빈 줄(공백) 제거
        korean_lines = [line.strip() for line in text.strip().split('\n') if line.strip()]
        # 텍스트를 두개씩 하나의 그룹으로 형성
        korean_pairs = self.group_pairs(korean_lines)

        # PPT 생성
        prs = Presentation()
        self.setting_slide_size_16_9(prs)
        slide_layout = self.setting_slide_layout(prs)

        # 슬라이드별로 텍스트 상자를 순회하며 처리
        for slide_index, ko_pair in enumerate(korean_pairs, start=1):

            # 빈 슬라이드 생성
            slide = prs.slides.add_slide(slide_layout)

            # 슬라이드 배경을 검은색으로 설정
            background = slide.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(0, 0, 0)

            # 텍스트 박스 1: 한글 텍스트
            left_ko, top_ko, width_ko, height_ko = self.create_text_box(1.6667, 2.55, 10, 1.476)
            # 텍스트 박스 2: 영어 텍스트
            left_en, top_en, width_en, height_en = self.create_text_box(1.6667, 4.572, 10, 0.933)

            # # 슬라이드에 텍스트 박스 추가
            txBox_ko = slide.shapes.add_textbox(left_ko, top_ko, width_ko, height_ko)
            txBox_en = slide.shapes.add_textbox(left_en, top_en, width_en, height_en)

            # 텍스트 입력
            korean_text = self.input_text(ko_pair)
            txBox_ko.text = korean_text

            # 텍스트 설정
            for paragraph in txBox_ko.text_frame.paragraphs:
                self.font_settings(paragraph, "나눔고딕", "40")

            for paragraph in txBox_en.text_frame.paragraphs:
                self.font_settings(paragraph, "나눔고딕", "15")

        self.save(prs)

    def create_only_english_lyrics(self):
        text = self.read_text_file()

        # 텍스트를 줄 단위로 분리하고 빈 줄(공백) 제거
        english_lines = [line.strip() for line in text.strip().split('\n') if line.strip()]
        # 텍스트를 두개씩 하나의 그룹으로 형성
        english_pairs = self.group_pairs(english_lines)

        # PPT 생성
        prs = Presentation()
        self.setting_slide_size_16_9(prs)
        slide_layout = self.setting_slide_layout(prs)

        # 슬라이드별로 텍스트 상자를 순회하며 처리
        for slide_index, en_pair in enumerate(english_pairs, start=1):

            # 빈 슬라이드 생성
            slide = prs.slides.add_slide(slide_layout)

            # 슬라이드 배경을 검은색으로 설정
            background = slide.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(0, 0, 0)

            # 텍스트 박스 1: 한글 텍스트
            left_ko, top_ko, width_ko, height_ko = self.create_text_box(1.6667, 2.55, 10, 1.476)
            # 텍스트 박스 2: 영어 텍스트
            left_en, top_en, width_en, height_en = self.create_text_box(1.6667, 4.572, 10, 0.933)

            # # 슬라이드에 텍스트 박스 추가
            txBox_ko = slide.shapes.add_textbox(left_ko, top_ko, width_ko, height_ko)
            txBox_en = slide.shapes.add_textbox(left_en, top_en, width_en, height_en)

            # 텍스트 입력
            english_text = self.input_text(en_pair)
            txBox_en.text = english_text

            # 텍스트 설정
            for paragraph in txBox_ko.text_frame.paragraphs:
                self.font_settings(paragraph, "나눔고딕", "40")

            for paragraph in txBox_en.text_frame.paragraphs:
                self.font_settings(paragraph, "나눔고딕", "15")

        self.save(prs)

    def create_seoul_form(self):
        text = self.read_text_file()

        # 한글과 영어 라인 분리 (홀수 인덱스: 한글, 짝수 인덱스: 영어)
        korean_lines = self.sep_lyrics(text, 0, 2)
        english_lines = self.sep_lyrics(text, 1, 2)
        korean_pairs = self.group_pairs(korean_lines)
        english_pairs = self.group_pairs(english_lines)
        # PPT 생성
        prs = Presentation()
        self.setting_slide_size_16_9(prs)
        slide_layout = self.setting_slide_layout(prs)

        # 슬라이드별로 텍스트 상자를 순회하며 처리
        for slide_index, (ko_pair, en_pair) in enumerate(zip(korean_pairs, english_pairs), start=1):

            # 빈 슬라이드 생성
            slide = prs.slides.add_slide(slide_layout)

            # 슬라이드 배경을 검은색으로 설정
            background = slide.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(0, 0, 0)

            # 텍스트 박스 1: 한글 텍스트
            left_ko, top_ko, width_ko, height_ko = self.create_text_box(0.8417, 0.5723, 11.6538, 1.4818)
            # 텍스트 박스 2: 영어 텍스트
            left_en, top_en, width_en, height_en = self.create_text_box(0.3111, 2.0536, 12.7131, 1.0432)

            # # 슬라이드에 텍스트 박스 추가
            txBox_ko = slide.shapes.add_textbox(left_ko, top_ko, width_ko, height_ko)
            txBox_en = slide.shapes.add_textbox(left_en, top_en, width_en, height_en)

            # 텍스트 입력
            korean_text = self.input_text(ko_pair)
            english_text = self.input_text(en_pair)
            txBox_ko.text = korean_text
            txBox_en.text = english_text

            # 텍스트 설정
            for paragraph in txBox_ko.text_frame.paragraphs:
                self.font_settings(paragraph, "나눔고딕OTF ExtraBold", "41")

            for paragraph in txBox_en.text_frame.paragraphs:
                self.font_settings(paragraph, "나눔고딕OTF ExtraBold", "28")

        self.save(prs)
