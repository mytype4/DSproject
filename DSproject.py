import openpyxl
from datetime import datetime
from openpyxl.styles import Alignment, Font
from selenium import webdriver
from bs4 import BeautifulSoup
from urllib.parse import quote_plus
from selenium.webdriver.chrome.options import Options
from trafilatura import fetch_url, extract

print("<<차동성 뉴스 수집 자동화 모듈>> \n현재 뉴스 수집 중입니다. 창을 닫지 말아주세요.")

# Chrome background 실행 옵션
options = Options()
options.add_argument("headless")
driver = webdriver.Chrome(options=options)

# 시트 이름 리스트 읽기
with open('검색어 설정.txt', 'r', encoding='utf-8') as f:
    sheet_names = f.read().splitlines()

# 현재 날짜로 파일 이름 생성
today = datetime.today()
formatted_date = today.strftime("%Y. %#m. %#d")
file_name = f"{formatted_date} 뉴스 모음.xlsx"

# 워크북 생성
wb = openpyxl.Workbook()
ws = wb.active

# 시트에 데이터를 추가하는 함수
def add_sheet_template(sheet, sheet_name, sources, titles, contents, links):
    # 기본 헤더
    sheet['C3'] = '검색어'
    sheet['B5'] = '연번'
    sheet['C5'] = '출처'
    sheet['D5'] = '제목'
    sheet['E5'] = '본문 내용'
    sheet['F5'] = '링크'
    sheet.column_dimensions['C'].width = 18
    sheet.column_dimensions['D'].width = 45
    sheet.column_dimensions['E'].width = 180
    sheet['D3'] = sheet_name
    sheet.freeze_panes = "A4"

    # 모든 셀에 적용할 글자 크기
    default_font = Font(size=15)  # 글자 크기 설정

    # 각 리스트의 길이가 동일하다고 가정하고 데이터 행 추가
    for index in range(len(sources)):
        row = 6 + index
        sheet[f'B{row}'] = index + 1
        sheet[f'C{row}'] = sources[index]
        sheet[f'D{row}'] = titles[index]
        sheet[f'E{row}'] = contents[index]
        sheet[f'F{row}'] = links[index]

        # C열과 D열에 줄바꿈 적용
        sheet[f'C{row}'].alignment = Alignment(wrap_text=True)
        sheet[f'D{row}'].alignment = Alignment(wrap_text=True)

        # F열에 하이퍼링크 적용
        sheet[f'F{row}'].hyperlink = links[index]
        sheet[f'F{row}'].style = "Hyperlink"

        # 각 셀에 글자 크기 20 적용
        sheet[f'B{row}'].font = default_font
        sheet[f'C{row}'].font = default_font
        sheet[f'D{row}'].font = default_font
        sheet[f'E{row}'].font = default_font
        sheet[f'F{row}'].font = default_font

    # 헤더 셀에도 글자 크기 20 적용
    for cell in ['B5', 'C5', 'D5', 'E5', 'F5']:
        sheet[cell].font = default_font
    for cell in ['C3', 'D3']:
        sheet[cell].font = Font(size=30)

# 검색어마다 시트 생성 및 뉴스 수집
for index, keyword in enumerate(sheet_names):
    # 검색어를 사용해 뉴스 검색 URL 생성
    baseUrl = 'https://google.com/search?q='
    news = '&tbm=nws'
    url = baseUrl + quote_plus(keyword) + news

    # Chrome 드라이버로 페이지 로드
    driver.get(url)
    driver.implicitly_wait(5)

    # BeautifulSoup로 페이지 소스 파싱
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')

    # 검색 결과 컨테이너 선택
    v = soup.select('.SoaBEf, .W0kfrc')

    sources = []
    links = []
    titles = []
    contents = []

    # 결과 정보 추출
    for item in v:
        source = item.select_one('.MgUUmf.NUnG9d').text
        link = item.select_one('a').attrs['href']
        title = item.select_one('.n0jPhd.ynAwRc.MBeuO.nDgy9d').text

        sources.append(source)
        links.append(link)
        titles.append(title)

    # 각 링크로부터 본문 내용 추출
    for feed in links:
        html = fetch_url(feed)
        text = extract(html)
        contents.append(text)

    print(sources)
    # 첫 번째 검색어에 해당하는 시트 제목 설정
    if index == 0:
        ws.title = keyword
        add_sheet_template(ws, keyword, sources, titles, contents, links)
    else:  # 나머지 시트 생성 및 템플릿 적용
        new_sheet = wb.create_sheet(title=keyword)
        add_sheet_template(new_sheet, keyword, sources, titles, contents, links)

# Chrome 드라이버 닫기
driver.quit()

# 파일 저장
wb.save(file_name)
print('완료되었습니다. 오늘 하루도 화이팅입니다!')