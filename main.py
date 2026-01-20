import requests
import pandas as pd
import time
import re
import sys

# ==========================================
# [설정 영역]
# ==========================================
API_KEY = "devU01TX0FVVEgyMDI2MDEyMDExNTIxNjExNzQ2MTE=" 
INPUT_FILE = "input.xlsx" 
ADDRESS_COL = "주소"
# ==========================================

def advanced_clean_text(text):
    """
    [전문가용 전처리 v3.0]
    오타(점) 제거 및 강력한 띄어쓰기 분리 로직 적용
    """
    if not isinstance(text, str):
        return str(text)

    # -----------------------------------------------------------
    # [1단계] 오타 및 불필요한 특수문자 청소 (가장 먼저 실행!)
    # -----------------------------------------------------------
    
    # ★ 요청하신 기능: 주소 중간에 낀 점(.) 삭제
    # 예: "289-.2" -> "289-2" (정상 번지수로 변환됨)
    text = text.replace('.', '')
    
    # 괄호와 그 안의 내용 제거
    text = re.sub(r'\(.*?\)', ' ', text)
    
    # 물결표(~) 등 특수문자 제거
    text = text.replace('~', ' ')
    
    # -----------------------------------------------------------
    # [2단계] 붙어있는 주소 강제 분리 수술
    # -----------------------------------------------------------
    
    # A. '동/리/가/읍/면' 뒤에 숫자나 한글이 바로 오면 띄움 (독산동진도 -> 독산동 진도)
    text = re.sub(r'(동|리|가|읍|면)([0-9가-힣])', r'\1 \2', text)
    
    # B. '번지' 뒤에 뭔가 붙어있으면 띄움 (2번지성진 -> 2번지 성진)
    text = re.sub(r'(번지)([0-9가-힣])', r'\1 \2', text)

    # C. 숫자 뒤에 한글이 오면 띄움 (345-23대오 -> 345-23 대오)
    # 단, '1동', '2호', '3층' 등은 붙어있어야 하므로 제외
    text = re.sub(r'(\d)([가-힣])', r'\1 \2', text)
    # 위에서 벌어진 '1 동', '2 호' 등을 다시 붙여주는 후처리
    text = re.sub(r'(\d) (동|호|층|번지)', r'\1\2', text)

    # D. 한글 뒤에 숫자가 오면 띄움 (푸르지오104 -> 푸르지오 104)
    text = re.sub(r'([가-힣])(\d)', r'\1 \2', text)

    # 불필요한 공백 정리 (공백 2개 -> 1개)
    text = re.sub(r'\s+', ' ', text)
    
    return text.strip()

def get_road_address(keyword):
    url = "https://business.juso.go.kr/addrlink/addrLinkApi.do"
    params = {
        "confmKey": API_KEY,
        "currentPage": 1,
        "countPerPage": 1, 
        "keyword": keyword,
        "resultType": "json"
    }
    
    try:
        response = requests.get(url, params=params, timeout=5)
        if response.status_code == 200:
            data = response.json()
            if data['results']['common']['totalCount'] != '0':
                return data['results']['juso'][0]
    except Exception:
        pass
    return None

def print_progress_bar(iteration, total, prefix='', suffix='', decimals=1, length=30, fill='■'):
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filled_length = int(length * iteration // total)
    bar = fill * filled_length + '□' * (length - filled_length)
    sys.stdout.write(f'\r{prefix} [{bar}] {percent}% ({iteration}/{total}) {suffix}')
    sys.stdout.flush()

def main():
    print("\n" + "="*60)
    print(">>> 주소 변환 시스템 v3.1 (오타 제거 패치)")
    print(">>> 적용된 기술: 점(.) 제거 및 정밀 띄어쓰기")
    print("="*60)
    
    df = None
    engines = [None, 'openpyxl', 'xlrd']
    for engine in engines:
        try:
            if engine:
                df = pd.read_excel(INPUT_FILE, engine=engine)
            else:
                df = pd.read_excel(INPUT_FILE)
            break
        except Exception:
            continue
            
    if df is None:
        print(f"❌ 오류: '{INPUT_FILE}' 파일을 읽을 수 없습니다.")
        return

    print(f"✅ 파일 로드 완료! 총 {len(df)}건의 데이터를 변환합니다.\n")
    
    success_list = []
    fail_list = []
    total = len(df)

    print_progress_bar(0, total, prefix='🚀 진행 중:', suffix='준비 중...', length=30)

    for index, row in df.iterrows():
        original_addr = str(row[ADDRESS_COL])
        
        # 1. 강력해진 전처리 (점 제거 포함)
        cleaned_addr = advanced_clean_text(original_addr)
        
        # 2. 검색 (Onion Peeling)
        words = cleaned_addr.split()
        result = None
        current_words = words[:]
        
        while len(current_words) >= 2:
            search_keyword = " ".join(current_words)
            result = get_road_address(search_keyword)
            if result:
                break
            current_words.pop()

        if result:
            row_data = row.to_dict()
            row_data['도로명주소'] = result['roadAddr']
            row_data['우편번호'] = result['zipNo']
            success_list.append(row_data)
        else:
            fail_list.append(row.to_dict())
            
        # 진행률 표시
        status_msg = f"- 성공: {len(success_list)}건"
        print_progress_bar(index + 1, total, prefix='🚀 진행 중:', suffix=status_msg, length=30)
        
        # 속도 최적화
        time.sleep(0.01)

    print("\n\n" + "="*60)
    
    if success_list:
        pd.DataFrame(success_list).to_excel("변환결과_성공.xlsx", index=False)
        print(f"👍 성공: {len(success_list)}건 -> '변환결과_성공.xlsx'")
    
    if fail_list:
        pd.DataFrame(fail_list).to_excel("변환결과_실패.xlsx", index=False)
        print(f"⚠️ 실패: {len(fail_list)}건 -> '변환결과_실패.xlsx'")
        print("   (참고: '단독주택' 같이 구체적인 주소가 없는 데이터는 변환되지 않습니다.)")

if __name__ == "__main__":
    main()