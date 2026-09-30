import io
import re
import urllib.parse
import openpyxl
import pandas as pd
import requests
import streamlit as st
from openpyxl.styles import Font

CONFIRM_KEY = "devU01TX0FVVEgyMDI2MDkzMDEwMTcwMDEyMDUzMjM="

# --- [ 주소 변환 로직 함수 ] ---
def fix_zipcode(val):
    if pd.isna(val) or val is None:
        return ""
    val_str = str(val).split('.')[0].strip()
    return val_str.zfill(5) if val_str else ""

def remove_duplicate_words(addr_str):
    if not addr_str:
        return addr_str
    
    bd_match = re.search(r'\(([^)]+)\)', addr_str)
    if bd_match:
        inner_bracket = bd_match.group(1)
        bracket_parts = [p.strip() for p in inner_bracket.split(',')]
        
        seen = set()
        unique_parts = []
        for p in bracket_parts:
            if p and p not in seen:
                seen.add(p)
                unique_parts.append(p)
        
        new_inner = ", ".join(unique_parts)
        addr_str = addr_str[:bd_match.start(1)] + new_inner + addr_str[bd_match.end(1):]

    words = addr_str.split()
    clean_words = []
    for w in words:
        clean_w = w.strip('(),')
        if not clean_words or clean_w != clean_words[-1].strip('(),'):
            clean_words.append(w)
            
    return ' '.join(clean_words)

def master_juso_converter(keyword, session):
    if not keyword or pd.isna(keyword):
        return keyword
    
    kw_str = str(keyword).strip()
    kw_str = re.sub(r'(\d+동|\d+호|\d+층|B\d+호|물리치료실)', r' \1 ', kw_str)
    kw_str = ' '.join(kw_str.split())
    
    if '월산동 986-3' in kw_str or '월산동 986' in kw_str:
        extra = kw_str.replace('광주광역시', '').replace('전남광주통합특별시', '').replace('남구', '').replace('월산동', '').replace('986-3', '').replace('986', '').strip()
        return remove_duplicate_words(f"광주광역시 남구 대남대로 363 (월산동) {extra}".strip())

    tokens = kw_str.split()
    base_tokens = []
    extra_details = []
    
    for t in tokens:
        if re.search(r'(\d+동|\d+호|\d+층|B\d+호|물리치료실)', t):
            extra_details.append(t)
        else:
            base_tokens.append(t)
            
    search_q1 = " ".join(base_tokens)
    
    pure_jibeon_tokens = []
    building_tokens = []
    found_jibeon = False
    
    for t in base_tokens:
        if not found_jibeon:
            pure_jibeon_tokens.append(t)
            if re.match(r'^\d+(-\d+)?$', t):
                found_jibeon = True
        else:
            building_tokens.append(t)

    pure_jibeon = " ".join(pure_jibeon_tokens)
    building_name_candidate = " ".join(building_tokens)
    sido_sigungu_dong = " ".join([t for t in pure_jibeon_tokens if not re.match(r'^\d+(-\d+)?$', t)])
    
    query_candidates = []
    if pure_jibeon:
        query_candidates.append(pure_jibeon)
    if search_q1 and search_q1 not in query_candidates:
        query_candidates.append(search_q1)
    if sido_sigungu_dong and building_name_candidate:
        query_candidates.append(f"{sido_sigungu_dong} {building_name_candidate}")
    if pure_jibeon and '-' in pure_jibeon:
        main_jibeon_str = re.sub(r'(-\d+)', '', pure_jibeon)
        if main_jibeon_str not in query_candidates:
            query_candidates.append(main_jibeon_str)
    if kw_str not in query_candidates:
        query_candidates.append(kw_str)

    base_road_addr = ""
    api_bd_nm = ""
    
    for q in query_candidates:
        if not q.strip():
            continue
        url = f"https://business.juso.go.kr/addrlink/addrLinkApi.do?currentPage=1&countPerPage=10&keyword={urllib.parse.quote(q)}&confmKey={CONFIRM_KEY}&resultType=json"
        
        try:
            response = session.get(url, timeout=3)
            if response.status_code == 200:
                res = response.json()
                juso_list = res.get('results', {}).get('juso') or []
                
                if juso_list:
                    for juso in juso_list:
                        bd_name = juso.get('bdNm', '')
                        if bd_name and building_name_candidate and any(part in building_name_candidate for part in bd_name.split() if len(part) > 1):
                            base_road_addr = juso.get('roadAddr')
                            api_bd_nm = bd_name
                            break
                    
                    if not base_road_addr:
                        base_road_addr = juso_list[0].get('roadAddr')
                        api_bd_nm = juso_list[0].get('bdNm', '')
                        
                    if base_road_addr:
                        break
        except Exception:
            continue

    if not base_road_addr:
        return remove_duplicate_words(kw_str)

    target_bd = api_bd_nm.strip() if api_bd_nm else building_name_candidate.strip()
    if target_bd and target_bd not in base_road_addr:
        if '(' in base_road_addr and ')' in base_road_addr:
            base_road_addr = re.sub(r'\(([^)]+)\)', r'(\1, ' + target_bd + ')', base_road_addr)
        else:
            base_road_addr = f"{base_road_addr} ({target_bd})"

    full_result = base_road_addr
    if extra_details:
        needed_details = []
        for p in extra_details:
            p_clean = p.strip('()')
            if p_clean not in base_road_addr and p not in base_road_addr:
                needed_details.append(p)
        if needed_details:
            full_result = f"{base_road_addr} {' '.join(needed_details)}"

    return remove_duplicate_words(full_result)


# --- [ Streamlit 웹 UI ] ---
st.set_page_config(page_title="자동 주소 변환기", page_icon="🚚", layout="centered")

st.title("🚚 만능 주소 변환 & 엑셀 정제 웹 앱")
st.write("엑셀 파일을 업로드하면 도로명 주소 변환, 우편번호 0 보존, 엑셀 서식을 자동으로 적용해 줍니다.")

uploaded_file = st.file_uploader("변환할 엑셀 파일(.xlsx)을 업로드하세요", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("### 📄 업로드 데이터 미리보기 (상위 5건)")
    st.dataframe(df.head())

    if st.button("🚀 주소 변환 및 서식 적용 시작"):
        with st.spinner("주소를 변환하는 중입니다... 데이터 양에 따라 시간이 걸릴 수 있습니다."):
            session = requests.Session()
            
            if '우편번호' in df.columns:
                df['우편번호'] = df['우편번호'].apply(fix_zipcode)

            if '배송지' in df.columns:
                progress_bar = st.progress(0)
                total = len(df)
                
                converted_addrs = []
                for i, row in df.iterrows():
                    res = master_juso_converter(row['배송지'], session)
                    converted_addrs.append(res)
                    progress_bar.progress((i + 1) / total)
                    
                df['배송지'] = converted_addrs
            else:
                st.warning("[주의] '배송지' 컬럼을 찾을 수 없습니다.")

            target_columns = ['수취인명', '전화', '우편번호', '배송지', '선택정보', '기타', '구분']
            if len(df.columns) == len(target_columns):
                df.columns = target_columns

            # 메모리 내 엑셀 세이빙 (서식 적용)
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