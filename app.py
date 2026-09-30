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

# --- [ 스마트 주소 변환 엔진 (상가 오매칭 방지 로직 보완) ] ---
def master_juso_converter(keyword, session=None):
    if not keyword or pd.isna(keyword):
        return keyword
    
    if session is None:
        session = requests.Session()
        
    kw_str = str(keyword).strip()
    
    # 1. 건물명 뒤 '105-103' 형태를 '105동 103호'로 자동 변환
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
    
    # 2. 동/호수/층 자동 띄어쓰기 전처리
    kw_str = re.sub(r'(\d+동|\d+호|\d+층|B\d+호|물리치료실)', r' \1 ', kw_str)
    kw_str = ' '.join(kw_str.split())
    
    # 3. 예외 및 특수 주소 처리
    if '월산동 986-3' in kw_str or '월산동 986' in kw_str:
        extra = kw_str.replace('광주광역시', '').replace('전남광주통합특별시', '').replace('남구', '').replace('월산동', '').replace('986-3', '').replace('986', '').strip()
        return remove_duplicate_words(f"광주광역시 남구 대남대로 363 (월산동) {extra}".strip())

    # 4. 동/호수 부가정보와 기본 주소 분리
    tokens = kw_str.split()
    base_tokens = []
    extra_details = []
    
    for t in tokens:
        if re.search(r'(\d+동|\d+호|\d+층|B\d+호|물리치료실)', t):
            extra_details.append(t)
        else:
            base_tokens.append(t)
            
    search_q1 = " ".join(base_tokens)
    
    # 5. 스마트 토큰 분류
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
    
    if search_q1:
        query_candidates.append(search_q1)
        
    if sido_sigungu_dong and jibeon_token:
        query_candidates.append(f"{sido_sigungu_dong} {jibeon_token}")
        
    if sido_sigungu_dong and building_name_candidate:
        query_candidates.append(f"{sido_sigungu_dong} {building_name_candidate}")

    sido_sigungu_only = " ".join([t for t in sido_sigungu_dong_tokens if not (t.endswith('동') or t.endswith('리') or t.endswith('가'))])
    if sido_sigungu_only and building_name_candidate:
        query_candidates.append(f"{sido_sigungu_only} {building_name_candidate}")

    if jibeon_token and '-' in jibeon_token:
        main_jibeon = jibeon_token.split('-')[0]
        query_candidates.append(f"{sido_sigungu_dong} {main_jibeon}")

    if kw_str not in query_candidates:
        query_candidates.append(kw_str)

    base_road_addr = ""
    api_bd_nm = ""
    is_user_sangga = '상가' in kw_str  # 사용자가 상가를 직접 입력했는지 확인
    
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
                    selected_juso = None
                    
                    # (단계 1) 건물명 정확(Exact) 일치 검사 (상가 제외)
                    for juso in juso_list:
                        bd_name = juso.get('bdNm', '').strip()
                        if not is_user_sangga and '상가' in bd_name:
                            continue
                        if bd_name and building_name_candidate and bd_name == building_name_candidate:
                            selected_juso = juso
                            break

                    # (단계 2) 건물명 부분 일치 검사 (상가 제외)
                    if not selected_juso:
                        for juso in juso_list:
                            bd_name = juso.get('bdNm', '').strip()
                            if not is_user_sangga and '상가' in bd_name:
                                continue
                            if bd_name and building_name_candidate and any(part in building_name_candidate for part in bd_name.split() if len(part) > 1):
                                selected_juso = juso
                                break

                    # (단계 3) 조건에 맞는 항목이 없으면 상가 제외 목록 중 첫 번째
                    if not selected_juso and not is_user_sangga:
                        for juso in juso_list:
                            if '상가' not in juso.get('bdNm', ''):
                                selected_juso = juso
                                break

                    # (단계 4) 그래도 없으면 목록의 첫 번째 항목
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

    # 7. 건물명 자동 보완 결합
    target_bd = api_bd_nm.strip() if api_bd_nm else building_name_candidate.strip()
    if target_bd and target_bd not in base_road_addr:
        if '(' in base_road_addr and ')' in base_road_addr:
            base_road_addr = re.sub(r'\(([^)]+)\)', r'(\1, ' + target_bd + ')', base_road_addr)
        else:
            base_road_addr = f"{base_road_addr} ({target_bd})"

    # 8. 동/호수 부가정보 재결합
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