#pip install tensorflow
#pip install keras
#pip install git+https://github.com/haven-jeon/PyKoSpacing.git

import os
import re
from collections import defaultdict
from openpyxl import Workbook
from pykospacing import Spacing

# 한글 패턴 정의
korean_pattern = re.compile(r'[가-힣]')
# Java 파일 주석 패턴 (단일 줄 주석, 블록 주석, JavaDoc 주석 제거)
java_comment_pattern = re.compile(r'//.*|/\*[\s\S]*?\*/|/\*\*[\s\S]*?\*/')
# HTML 파일 주석 패턴 (단일 줄 주석과 블록 주석 제거)
html_comment_pattern = re.compile(r'<!--[\s\S]*?-->|//.*')
# HTML 태그와 속성 제거 패턴
html_tag_pattern = re.compile(r'<[^>]+>')
# 태그 사이의 텍스트 추출 패턴
html_content_pattern = re.compile(r'>([^<]+)<')
# 스크립트 내에서 한글이 포함된 문자열 전체를 추출하는 패턴
script_string_pattern = re.compile(r'["\']([^"\']*[가-힣]+[^"\']*)["\']')
entity_annotation_pattern = re.compile(r'@\s*Entity')

# 제외할 파일 패턴 정의
exclude_patterns = [r'Enum\.java$', r'Dto\.java$', r'Repository\.java$', r'Test\.java$', r'Code\.java$']
spacing = Spacing()

def remove_comments_and_tags(line, is_java):
    """주석과 태그 제거 함수"""
    if line.strip().startswith('*'):
        return ''  # '*'로 시작하는 줄은 건너뜀
    original_line = line
    if is_java:
        line = java_comment_pattern.sub('', line)
    else:
        line = html_comment_pattern.sub('', line)
        line = html_tag_pattern.sub('', line)  # HTML 태그 제거
    cleaned_line = line.strip()
    return cleaned_line

def extract_korean_from_html_line(line):
    """HTML 파일에서 태그 사이의 텍스트와 스크립트 내 문자열을 추출"""
    korean_texts = set()

    # 태그 사이 텍스트 추출
    for content in html_content_pattern.findall(line):
        if korean_pattern.search(content):
            korean_texts.add(content.strip())

    # 스크립트 내 문자열 추출
    if '<script>' in line and '</script>' in line:
        for script_string in script_string_pattern.findall(line):
            if korean_pattern.search(script_string):
                korean_texts.add(script_string.strip())

    return korean_texts

def should_exclude_file(file_path):
    """제외 조건 확인 함수"""
    for pattern in exclude_patterns:
        if re.search(pattern, file_path):
            return True
    return False

def extract_korean_from_file(file_path, is_java):
    """파일에서 한글 텍스트 추출 함수"""
    korean_texts = defaultdict(lambda: {'files': set(), 'count': 0})

    with open(file_path, 'r', encoding='utf-8') as file:
        if is_java:
            # Enum, Dto, Repository으로 끝나는 파일 또는 @Entity 어노테이션을 가진 파일 제외
            if should_exclude_file(file_path):
                return korean_texts

            content = file.read()
            if entity_annotation_pattern.search(content):
                return korean_texts
            file.seek(0)  # 파일 포인터를 처음으로 되돌리기

        lines = file.readlines()
        for line in lines:
            if is_java:
                clean_line = remove_comments_and_tags(line, is_java).strip()
                if korean_pattern.search(clean_line):
                    korean_text = korean_pattern.findall(clean_line)
                    korean_text_str = ''.join(korean_text)
                    if korean_text_str:
                        spaced_text = spacing(korean_text_str)
                        korean_texts[spaced_text]['files'].add(os.path.basename(file_path))
                        korean_texts[spaced_text]['count'] += 1
            else:
                # HTML 파일의 경우 태그와 스크립트 내 텍스트를 추출
                extracted_texts = extract_korean_from_html_line(line)
                for text in extracted_texts:
                    spaced_text = spacing(text)
                    korean_texts[spaced_text]['files'].add(os.path.basename(file_path))
                    korean_texts[spaced_text]['count'] += 1
    return korean_texts

def search_files(directory, extension):
    """디렉토리에서 특정 확장자를 가진 파일을 검색"""
    files = []
    for root, _, filenames in os.walk(directory):
        for filename in filenames:
            if filename.endswith(extension):
                files.append(os.path.join(root, filename))
    return files

def write_to_excel(data, output_file):
    """엑셀 파일로 데이터 저장"""
    wb = Workbook()
    ws = wb.active
    ws.append(["Text", "File Names", "Frequency"])
    for text, info in data.items():
        ws.append([text, ', '.join(sorted(info['files'])), info['count']])
    wb.save(output_file)

def main(project_dir, output_file):
    all_korean_texts = defaultdict(lambda: {'files': set(), 'count': 0})

    # .java 파일 처리
    java_files = search_files(project_dir, '.java')
    for java_file in java_files:
        file_korean_texts = extract_korean_from_file(java_file, is_java=True)
        for text, info in file_korean_texts.items():
            all_korean_texts[text]['files'].update(info['files'])
            all_korean_texts[text]['count'] += info['count']

    # .html 파일 처리
    html_files = search_files(project_dir, '.html')
    for html_file in html_files:
        file_korean_texts = extract_korean_from_file(html_file, is_java=False)
        for text, info in file_korean_texts.items():
            all_korean_texts[text]['files'].update(info['files'])
            all_korean_texts[text]['count'] += info['count']

    # 엑셀로 저장
    write_to_excel(all_korean_texts, output_file)

if __name__ == "__main__":
    project_directory = "/Users/work-space"  # 프로젝트 루트 디렉토리 경로
    output_excel_file = "korean_texts_with_frequency.xlsx"  # 결과를 저장할 엑셀 파일명
    main(project_directory, output_excel_file)