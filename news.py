import requests
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
def search_articles(api_key, search_engine_id, keywords, max_pages=1):
    results = []

    for keyword in keywords:
        for page in range(max_pages):
            # API 요청 URL 생성
            start_index = page * 10 + 1  # Google API는 페이지당 10개의 결과를 반환
            url = f"https://www.googleapis.com/customsearch/v1?q={keyword}&key={api_key}&cx={search_engine_id}&start={start_index}"

            # API 요청
            response = requests.get(url)

            # 응답 상태 코드 확인
            if response.status_code == 200:
                json_response = response.json()
                items = json_response.get('items', [])
                for item in items:
                    # 작성일 추출
                      # 작성일 추출 및 포맷팅
                    published_date = None
                    if 'pagemap' in item and 'metatags' in item['pagemap']:
                        metatags = item['pagemap']['metatags']
                        if isinstance(metatags, list) and metatags:
                            for tag in metatags:
                                if 'article:published_time' in tag:
                                    published_date_str = tag['article:published_time']
                                    try:
                                        # ISO 8601 형식에서 datetime 객체로 변환
                                        published_date = datetime.fromisoformat(published_date_str.replace("Z", "+00:00"))
                                        # 원하는 형식으로 포맷
                                        published_date = published_date.strftime('%Y-%m-%d')
                                    except ValueError:
                                        published_date = None
                    title = item.get('title')
                    link = item.get('link')
                    results.append({'keyword':keyword,'title': title, 'link': link ,'published_date':published_date})

                # 검색 결과가 없으면 종료
                if len(items) < 10:  # 10개 미만의 결과가 반환되면 마지막 페이지로 간주
                    break
            else:
                print(f"Error: {response.status_code} - {response.text}")
                break  # 오류 발생 시 루프 종료

    return results
# 사용 예
api_key = 'AIzaSyDDRGDwskC6EQfYzQBc5twf64csPGEEd6Y'
search_engine_id = 'e53590f3b6a954c71'
keywords = ['MK', '선교사 자녀', 'Missionary kid','선교사']

articles = search_articles(api_key, search_engine_id, keywords)

# 결과를 데이터프레임으로 변환
df = pd.DataFrame(articles)

# 엑셀 파일로 저장
output_file = 'articles.xlsx'
df.to_excel(output_file, index=False)

# 엑셀 파일 열고 열 너비 조정
wb = load_workbook(output_file)
ws = wb.active

# 각 열의 너비를 조정
for column in ws.columns:
    max_length = 0
    column = [cell for cell in column]
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        except:
            pass
    adjusted_width = (max_length + 2)  # 여유 공간 추가
    ws.column_dimensions[column[0].column_letter].width = adjusted_width

# 조정된 내용을 엑셀 파일에 저장
wb.save(output_file)

print(f"Results saved to {output_file}")