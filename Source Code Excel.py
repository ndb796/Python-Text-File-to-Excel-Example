# -*- coding: utf-8 -*-
import os
import re
import xlsxwriter

source_path = "./Source Code/"
source_files = []

# 소스코드 폴더에 존재하는 전체 소스코드를 읽어 들입니다.
for root, dirs, files in os.walk(source_path):
    for name in files:
        # 각각의 파일 이름을 가져옵니다.
        full_name = os.path.join(root, name)
        title = full_name.replace(source_path, "")
        # 정규식으로 번호만 추출합니다.
        # 예를 들어 파일의 이름이 "1. Hello World.cpp"이라면 1이 반환됩니다.
        number = int(re.findall('\d+', title[0:])[0])
        source_files.append((number, full_name))

# 소스코드의 이름(번호)을 기준으로 정렬합니다.
source_files.sort()

# 엑셀 워크 북 및 워크 시트를 생성합니다.
workbook = xlsxwriter.Workbook('Source Codes.xlsx')
worksheet = workbook.add_worksheet()

row = 0
col = 0

# 모든 소스코드를 읽어 엑셀에 출력합니다.
for source_file in source_files:
    name = source_file[1]
    # 해당 이름의 소스코드 파일을 읽습니다.
    file = open(name, 'r')
    s = file.read()
    # 각 소스코드 파일의 이름과 소스코드 내용을 엑셀에 출력합니다.
    title = name.replace(source_path, "")
    worksheet.write(row, col, title)
    worksheet.write(row, col + 1, s)
    row = row + 1

workbook.close()