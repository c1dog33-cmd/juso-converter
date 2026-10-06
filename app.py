import streamlit as st
import pandas as pd
import openpyxl
import io
import re

st.set_page_config(page_title="사무실 통합 재고 및 매출 정산 시스템", layout="wide")

st.title("🏢 사무실 통합 재고 및 매출 정산 대시보드")
st.write("매일 우편배송/택배 출력 폼을 업로드하면, 기존 장부에 자동 누적 합산되고 일별판매표와 상단 요약 지표까지 완벽하게 연동됩니다.")

# 1. 파일 업로드 섹션
st.subheader("1. 파일 업로드")
col1, col2 = st.columns(2)

with col1:
    uploaded_master = st.file_uploader("📂 기존 통합 관리 엑셀 파일 업로드 (.xlsx)", type=["xlsx"])
with col2:
    uploaded_daily = st.file_uploader("📄 오늘 출고/우편배송 폼 파일 업로드 (.xls / .xlsx)", type=["xls", "xlsx"])

target_date = st.date_input("📅 정산 적용 출고 일자 선택")

if st.button("🚀 정산 및 대시보드 자동 업데이트 실행", type="primary"):
    if uploaded_master is None or uploaded_daily is None:
        st.error("기존 통합 엑셀 파일과 오늘의 출고 폼 파일을 모두 업로드해 주세요!")
    else:
        try:
            # 엑셀 워크북 로드
            wb = openpyxl.load_workbook(uploaded_master)
            ws_log = wb['출고세부일지']
            ws_daily = wb['일별판매표']
            
            # 일일 출고 폼 읽기
            if uploaded_daily.name.endswith('.xls'):
                df_daily = pd.read_excel(uploaded_daily, sheet_name=0, engine='xlrd')
            else:
                df_daily = pd.read_excel(uploaded_daily, sheet_name=0)
                
            date_str = target_date.strftime("%Y-%m-%d")
            
            # 1. 해당 날짜 데이터가 이미 있으면 기존 행 삭제 (중복 방지 / 덮어쓰기 기능)
            rows_to_delete = []
            for r in range(4, ws_log.max_row + 1):
                row_date = ws_log.cell(row=r, column=1).value
                if row_date and str(row_date).startswith(date_str):
                    rows_to_delete.append(r)
            
            for r in reversed(rows_to_delete):
                ws_log.delete_rows(r)
                
            # 2. 새 데이터 출고세부일지에 추가 (Append)
            next_row = ws_log.max_row + 1
            if next_row < 4:
                next_row = 4
                
            max_p_row = 35 # 상품단가표 기준 행
            
            for idx, row in df_daily.iterrows():
                recipient = row.get('수취인명', '')
                opt = str(row.get('선택정보', '')).strip()
                
                # 수량 안전하게 파싱
                qty_raw = row.get('Unnamed: 6', 1) if 'Unnamed: 6' in row else 1
                qty = 1
                try:
                    if pd.notna(qty_raw):
                        qty = int(float(str(qty_raw).strip()))
                except:
                    qty = 1
                
                has_film = False
                if '+필름' in opt or '필름' in opt:
                    has_film = True
                    
                clean_model = opt.replace('+필름', '').replace('필름', '').strip()
                clean_model = re.sub(r'\[\d+\]', '', clean_model).strip()
                clean_model = clean_model.replace(' 택배', '').strip()
                
                if clean_model == 'A20/A30':
                    model = 'A20/30'
                elif clean_model in ['노트10플러스', '노트10']:
                    model = '노트10플러스'
                elif clean_model == 'S10 5G':
                    model = 'S10'
                else:
                    model = clean_model
                    
                if '[2]' in opt:
                    qty = 2
                elif '[3]' in opt:
                    qty = 3
                    
                ship_type = '우편(1개)'
                if qty == 2 or '우편(2개)' in opt:
                    ship_type = '우편(2개)'
                elif has_film:
                    ship_type = '등기(필름)'
                elif '택배' in opt:
                    ship_type = '택배'
                    
                curr_row = next_row + idx
                ws_log.cell(row=curr_row, column=1, value=date_str)
                ws_log.cell(row=curr_row, column=2, value=recipient)
                ws_log.cell(row=curr_row, column=3, value=opt)
                ws_log.cell(row=curr_row, column=4, value=model)
                ws_log.cell(row=curr_row, column=5, value=ship_type)
                ws_log.cell(row=curr_row, column=6, value=qty)
                
                # 수식 입력
                ws_log.cell(row=curr_row, column=7, value=f'=IF(E{curr_row}="등기(필름)", 4900, IF(E{curr_row}="우편(1개)", VLOOKUP(D{curr_row}, 상품단가표!$B$4:$G${max_p_row}, 3, FALSE), IF(E{curr_row}="우편(2개)", VLOOKUP(D{curr_row}, 상품단가표!$B$4:$G${max_p_row}, 4, FALSE)/2, IF(E{curr_row}="등기", VLOOKUP(D{curr_row}, 상품단가표!$B$4:$G${max_p_row}, 5, FALSE), VLOOKUP(D{curr_row}, 상품단가표!$B$4:$G${max_p_row}, 6, FALSE)))))')
                ws_log.cell(row=curr_row, column=8, value=f'=IF(E{curr_row}="등기(필름)", VLOOKUP(D{curr_row}, 상품단가표!$B$4:$G${max_p_row}, 2, FALSE) + 1200, VLOOKUP(D{curr_row}, 상품단가표!$B$4:$G${max_p_row}, 2, FALSE))')
                ws_log.cell(row=curr_row, column=9, value=f'=IF(OR(E{curr_row}="등기(필름)", E{curr_row}="등기"), 1800, IF(E{curr_row}="우편(1le)", 590, IF(E{curr_row}="우편(2개)", 710, 2600)))')
                ws_log.cell(row=curr_row, column=10, value=f'=F{curr_row}*G{curr_row}')
                ws_log.cell(row=curr_row, column=11, value=f'=F{curr_row}*H{curr_row}')
                ws_log.cell(row=curr_row, column=12, value=f'=I{curr_row}')
                ws_log.cell(row=curr_row, column=13, value=f'=J{curr_row}-K{curr_row}-L{curr_row}')
                
                # 서식 적용
                for c in range(1, 14):
                    cell = ws_log.cell(row=curr_row, column=c)
                    cell.font = openpyxl.styles.Font(name='맑은 고딕', size=10)
                    cell.border = openpyxl.styles.Border(left=openpyxl.styles.Side(style='thin', color='D9D9D9'),
                                                         right=openpyxl.styles.Side(style='thin', color='D9D9D9'),
                                                         top=openpyxl.styles.Side(style='thin', color='D9D9D9'),
                                                         bottom=openpyxl.styles.Side(style='thin', color='D9D9D9'))
                    if c in [6, 7, 8, 9, 10, 11, 12, 13]:
                        cell.alignment = openpyxl.styles.Alignment(horizontal='right', vertical='center')
                        cell.number_format = '#,##0'
                    else:
                        cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')

            # 3. 일별판매표 수식 연동 및 0값 보기 좋게 숨기기 설정 (셀 서식 '0;-0;""')
            max_log_row = ws_log.max_row
            for r in range(7, 39):  # 모델 목록 행 범위
                model_name = ws_daily.cell(row=r, column=2).value # B열: 모델명
                if model_name:
                    for day_idx in range(1, 32):
                        col_idx = 8 + day_idx # I열(9)부터 AM열(39)
                        d_str = f"{target_date.strftime('%Y-%m')}-{day_idx:02d}"
                        
                        formula = f'=SUMIFS(출고세부일지!$F$4:$F${max_log_row}, 출고세부일지!$D$4:$D${max_log_row}, B{r}, 출고세부일지!$A$4:$A${max_log_row}, "{d_str}")'
                        
                        cell = ws_daily.cell(row=r, column=col_idx)
                        cell.value = formula
                        cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
                        cell.font = openpyxl.styles.Font(name='맑은 고딕', size=10)
                        # 값이 0일 때 빈칸처럼 보이게 하는 엑셀 사용자 지정 서식 적용
                        cell.number_format = '#,##0;(#,##0);""'

            # 4. 상단 요약 지표(월 총 출고량 등) 자동 계산 수식 주입
            # 월수량 합계 (D열)
            ws_daily['B4'] = '=SUM(D7:D38)'
            # 월 총 매출액 (E열)
            ws_daily['H4'] = '=SUM(E7:E38)'
            # 월 총 순이익 (H열)
            ws_daily['J4'] = '=SUM(H7:H38)'

            # 가상 메모리에 저장
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            
            st.success(f"🎉 성공적으로 반영되었습니다! ({date_str} 기준 출고 내역, 일별판매표 및 상단 요약 연동 완료)")
            
            # 다운로드 버튼 제공
            st.download_button(
                label="📥 정산 완료된 엑셀 파일 다운로드하기",
                data=output,
                file_name=f"사무실_통합정산관리표_{date_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"처리 중 오류가 발생했습니다: {e}")
