import io
import re
import urllib.parse
import openpyxl
import pandas as pd
import requests
import streamlit as st
from openpyxl.styles import Font

CONFIRM_KEY = "devU01TX0FVVEgyMDI2MDkzMDEwMTcwMDEyMDUzMjM="

def fix_zipcode(val):
    if pd.isna(val) or val is None:
        return ""
    val_str = str(val).split('.')[0].strip()
    return val_str.zfill(5) if val_str else ""

# --- [ 중복 단어 및 괄호 정제 함수 ] ---
def remove_duplicate_words(addr_str):
    if not addr_str:
        return addr_str
    
    words = addr_str.split()
    clean_words = []
    for w in words:
        clean_w = w.strip('(),')
        if not clean_words or clean_w != clean_words[-1].strip('(),'):
            clean_words.append(w)
            
    return ' '.join(clean_words)

# --- [ 만능 주소 변환 엔진 ] ---
def master_juso_converter(keyword):
    if not keyword or pd.isna(keyword):
        return keyword
        
    kw_str = str(keyword).strip()
    
    # [규칙 1] 인천 서구 불로동 -> 검단구 불로동 강제 매핑
    if '불로동' in kw_str:
        kw_str = kw_str.replace('서구', '검단구').replace('서해구', '검단구')
        if '인천광역시 검단구' not in kw_str and '인천 검단구' not in kw_str:
            kw_str = re.sub(r'인천광역시\s+서구', '인천광역시 검단구', kw_str)
            kw_str = re.sub(r'인천\s+서구', '인천 검단구', kw_str)

    # 1. 건물명 뒤 '103-401' 형태를 '103동 104호'로 자동 변환
    tokens_init = kw_str.split()
    processed_tokens = []
    for i, t in enumerate(tokens_init):
        if re.match(r'^\d+-\d+$', t):
            prev_token = tokens_init[i-1] if i > 0 else ""
            if prev_token and not any(prev_token.endswith(s) for s in ['동', '리', '가', '로', '길']):
                parts = t.split('-')
                processed_tokens.append(f"{parts[0]}동 {parts[1]}호")
            else:
                processed_tokens.append(t)
        else:
            processed_tokens.append(t)
    kw_str = " ".join(processed_tokens)
    
    # 2. 동/호수/층/가동/나동/A동/B동/관리실/택배보관함 등 상세 정보 자동 띄어쓰기 전처리
    kw_str = re.sub(r'(\d+동|[가-힣]동|\d+호|\d+층|B\d+호|관리실|택배보관함|물리치료실)', r' \1 ', kw_str)
    kw_str = ' '.join(kw_str.split())
    
    # 3. 특수 예외 처리 (월산동 등)
    if '월산동 986-3' in kw_str or '월산동 986' in kw_str:
        extra = kw_str.replace('광주광역시', '').replace('전남광주통합특별시', '').replace('남구', '').replace('월산동', '').replace('986-3', '').replace('986', '').strip()
        return remove_duplicate_words(f"광주광역시 남구 대남대로 363 {extra}".strip())

    # 4. 상세 부가정보(동, 호, 층, 가동, A동, 관리실, 택배보관함, 상호명 등)와 기본 주소 분리
    # 지번이나 도로명 주소 구성 요소가 아닌 뒷주소(상호명, 관리실 등)까지 모두 보존하기 위해 패턴 확장
    tokens = kw_str.split()
    base_tokens = []
    extra_details = []
    
    is_after_jibeon = False
    for t in tokens:
        if re.match(r'^\d+(-\d+)?$', t) or re.match(r'^산\d+(-\d+)?$', t):
            is_after_jibeon = True
            base_tokens.append(t)
        elif is_after_jibeon and re.search(r'(\d+동|[가-힣]동|\d+호|\d+층|B\d+호|관리실|택배보관함|물리치료실)', t):
            extra_details.append(t)
        elif not is_after_jibeon and re.search(r'(\d+동|[가-힣]동|\d+호|\d+층|B\d+호)', t):
            extra_details.append(t)
        else:
            if is_after_jibeon and not any(t.endswith(s) for s in ['도', '시', '구', '군', '동', '리', '가', '로', '길']):
                # 지번 뒤에 나오는 추가 건물명이나 상호명/부서명 등은 상세 정보로 안전하게 보관
                extra_details.append(t)
            else:
                base_tokens.append(t)
            
    search_q1 = " ".join(base_tokens)
    
    # 5. 스마트 토큰 분리
    sido_sigungu_dong_tokens = []
    jibeon_token = ""
    building_tokens = []
    
    for t in base_tokens:
        if re.match(r'^\d+(-\d+)?$', t) or re.match(r'^산\d+(-\d+)?$', t):
            jibeon_token = t
        elif any(t.endswith(s) for s in ['도', '시', '구', '군', '동', '리', '가', '로', '길']):
            if not jibeon_token:
                sido_sigungu_dong_tokens.append(t)
            else:
                building_tokens.append(t)
        else:
            building_tokens.append(t)

    sido_sigungu_dong = " ".join(sido_sigungu_dong_tokens)
    building_name_candidate = " ".join(building_tokens)
    
    # 6. 다단계 검색 후보군 생성
    query_candidates = []
    
    if '불로동' in kw_str:
        if sido_sigungu_dong and jibeon_token:
            query_candidates.append(f"인천광역시 검단구 불로동 {jibeon_token}")
        query_candidates.append(kw_str.replace('서구', '검단구').replace('서해구', '검단구'))

    elif '서구' in kw_str:
        seohae_q = kw_str.replace('서구', '서해구')
        if sido_sigungu_dong and jibeon_token:
            seohae_dong = sido_sigungu_dong.replace('서구', '서해구')
            query_candidates.append(f"{seohae_dong} {jibeon_token}")
        query_candidates.append(seohae_q)

    if sido_sigungu_dong and jibeon_token:
        query_candidates.append(f"{sido_sigungu_dong} {jibeon_token}")
        
    if search_q1 and search_q1 not in query_candidates:
        query_candidates.append(search_q1)
        
    if sido_sigungu_dong and building_name_candidate:
        query_candidates.append(f"{sido_sigungu_dong} {building_name_candidate}")

    if kw_str not in query_candidates:
        query_candidates.append(kw_str)

    base_road_addr = ""
    api_bd_nm = ""
    is_user_sangga = '상가' in kw_str
    
    for q in query_candidates:
        if not q.strip():
            continue
        url = f"https://business.juso.go.kr/addrlink/addrLinkApi.do?currentPage=1&countPerPage=10&keyword={urllib.parse.quote(q)}&confmKey={CONFIRM_KEY}&resultType=json"
        
        try:
            response = requests.get(url, timeout=5)
            if response.status_code == 200:
                res = response.json()
                juso_list = res.get('results', {}).get('juso') or []
                
                if juso_list:
                    selected_juso = None
                    
                    for juso in juso_list:
                        bd_name = juso.get('bdNm', '').strip()
                        if not is_user_sangga and '상가' in bd_name:
                            continue
                        if bd_name and building_name_candidate and (bd_name in building_name_candidate or building_name_candidate in bd_name):
                            selected_juso = juso
                            break

                    if not selected_juso and not is_user_sangga:
                        for juso in juso_list:
                            if '상가' not in juso.get('bdNm', ''):
                                selected_juso = juso
                                break

                    if not selected_juso:
                        selected_juso = juso_list[0]

                    base_road_addr = selected_juso.get('roadAddr')
                    api_bd_nm = selected_juso.get('bdNm', '')

                    if base_road_addr:
                        break
        except Exception:
            continue

    if not base_road_addr:
        return remove_duplicate_words(kw_str)

    # 7. API 결과 주소에서 법정동 괄호 제거
    def clean_road_addr(match):
        content = match.group(1)
        parts = [p.strip() for p in content.split(',') if p.strip()]
        valid_parts = [p for p in parts if not (p.endswith('동') or p.endswith('읍') or p.endswith('면') or p.endswith('리'))]
        if valid_parts:
            return " " + ", ".join(valid_parts)
        return ""

    base_road_addr = re.sub(r'\s*\(([^)]+)\)', clean_road_addr, base_road_addr)

    # 8. 아파트/건물명 보완 결합
    target_bd = api_bd_nm.strip() if api_bd_nm else building_name_candidate.strip()
    if target_bd and target_bd not in base_road_addr:
        base_road_addr = f"{base_road_addr} {target_bd}"

    # 9. 상세 부가정보(동, 호, 가동, 관리실, 택배보관함, 상호명 등) 재결합
    full_result = base_road_addr
    if extra_details:
        needed_details = []
        for p in extra_details:
            if p not in base_road_addr:
                needed_details.append(p)
        if needed_details:
            full_result = f"{base_road_addr} {' '.join(needed_details)}"

    return remove_duplicate_words(full_result)


# --- [ Streamlit 웹 UI ] ---
st.set_page_config(page_title="자동 주소 변환기", page_icon="🚚", layout="centered")

st.title("🚚 만능 주소 변환 & 엑셀 정제 웹 앱")
st.write("엑셀 파일을 업로드하면 도로명 주소 변환, 우편번호 0 보존, 엑셀 서식을 자동으로 적용해 줍니다.")

uploaded_file = st.file_uploader("변환할 엑셀 파일(.xlsx, .xls)을 업로드하세요", type=["xlsx", "xls"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("### 📄 업로드 데이터 미리보기 (상위 5건)")
    st.dataframe(df.head())

    if st.button("🚀 주소 변환 및 서식 적용 시작"):
        with st.spinner("주소를 변환하는 중입니다... 데이터 양에 따라 시간이 걸릴 수 있습니다."):
            if '우편번호' in df.columns:
                df['우편번호'] = df['우편번호'].apply(fix_zipcode)

            if '배송지' in df.columns:
                progress_bar = st.progress(0)
                total = len(df)
                
                converted_addrs = []
                for i, row in df.iterrows():
                    res = master_juso_converter(row['배송지'])
                    converted_addrs.append(res)
                    progress_bar.progress((i + 1) / total)
                    
                df['배송지'] = converted_addrs
            else:
                st.warning("[주의] '배송지' 컬럼을 찾을 수 없습니다.")

            target_columns = ['수취인명', '전화', '우편번호', '배송지', '선택정보', '기타', '구분']
            if len(df.columns) == len(target_columns):
                df.columns = target_columns

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
                worksheet = writer.sheets['Sheet1']
                
                font_size_8 = Font(size=8)
                for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=len(df.columns)):
                    worksheet.row_dimensions[row[0].row].height = 18
                    for cell in row:
                        cell.font = font_size_8
                        
                if '우편번호' in df.columns:
                    zip_col_idx = df.columns.get_loc('우편번호') + 1
                    for row in range(2, worksheet.max_row + 1):
                        worksheet.cell(row=row, column=zip_col_idx).number_format = '@'

            excel_data = output.getvalue()

        st.success("🎉 변환이 완벽하게 완료되었습니다!")
        
        st.download_button(
            label="📥 변환된 엑셀 파일 다운로드",
            data=excel_data,
            file_name="우편배송_변환결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
