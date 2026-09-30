import streamlit as st
import streamlit.components.v1 as components
import sqlite3
import pandas as pd
from datetime import date, datetime, timedelta
from dateutil.relativedelta import relativedelta
import io

# ==========================================
# 1. 페이지 설정
# ==========================================
st.set_page_config(
    page_title="(주)케이알메딕스 연차 및 장기 근속 포상 현황 대시보드",
    page_icon="🗓️",
    layout="wide"
)

# 메인 헤더 및 CSS 스타일
st.markdown("""
    <style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');

    html, body, [class*="css"] {
        font-family: 'Pretendard SemiBold', -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif;
    }

    .main-header {
        font-size: 2.0rem;
        font-weight: 700;
        color: #111827;
        margin-bottom: 0.3rem;
    }
    .sub-header {
        font-size: 0.95rem;
        color: #4B5563;
        margin-bottom: 1.5rem;
    }

    /* 오늘 버튼 스타일 */
    div[data-testid="stColumn"] > div > button[key="btn_today"] {
        margin-top: 28px;
        width: 100%;
        font-weight: 700;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. 데이터베이스 관리
# ==========================================
DB_FILE = "annual_leave.db"


def get_db_connection():
    return sqlite3.connect(DB_FILE)


def init_db():
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
                       CREATE TABLE IF NOT EXISTS employees
                       (
                           id
                           INTEGER
                           PRIMARY
                           KEY
                           AUTOINCREMENT,
                           department
                           TEXT
                           NOT
                           NULL
                           DEFAULT
                           '미지정',
                           name
                           TEXT
                           NOT
                           NULL,
                           hire_date
                           TEXT
                           NOT
                           NULL
                       )
                       """)
        cursor.execute("PRAGMA table_info(employees)")
        columns = [column[1] for column in cursor.fetchall()]
        if 'department' not in columns:
            cursor.execute("ALTER TABLE employees ADD COLUMN department TEXT NOT NULL DEFAULT '미지정'")
        conn.commit()


def insert_sample_data():
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM employees")
        if cursor.fetchone()[0] == 0:
            samples = [
                ("경영지원팀", "김민수", "2023-09-15"),
                ("개발팀", "이서연", "2025-09-01"),
                ("마케팅팀", "박준혁", "2019-09-20"),
                ("영업팀", "최유진", "2024-07-01"),
                ("개발팀", "정현우", "2021-05-12")
            ]
            cursor.executemany("INSERT INTO employees (department, name, hire_date) VALUES (?, ?, ?)", samples)
            conn.commit()


# ==========================================
# 3. 법정 연차 및 장기근속일 계산 로직
# ==========================================
def calculate_vacation_info(hire_date_str: str, target_date: date = None):
    if target_date is None:
        target_date = date.today()

    try:
        hire_date = datetime.strptime(str(hire_date_str)[:10], "%Y-%m-%d").date()
    except (ValueError, TypeError):
        hire_date = date.today()

    years = target_date.year - hire_date.year
    this_year_anniversary = date(target_date.year, hire_date.month, hire_date.day)
    if target_date < this_year_anniversary:
        years -= 1

    if years < 1:
        rd = relativedelta(target_date, hire_date)
        passed_months = rd.years * 12 + rd.months
        start_date = hire_date
        end_date = hire_date + relativedelta(years=1) - timedelta(days=1)

        promote_1st = end_date - relativedelta(months=3)
        promote_2nd = end_date - relativedelta(months=1)
        service_year_display = f"1년 미만 ({passed_months}개월 차)"

    else:
        start_date = hire_date + relativedelta(years=years)
        end_date = start_date + relativedelta(years=1) - timedelta(days=1)

        promote_1st = end_date - relativedelta(months=6)
        promote_2nd = end_date - relativedelta(months=2)
        service_year_display = f"{years + 1}년차"

    anniversary_1yr = hire_date + relativedelta(years=1)
    anniversary_3yr = hire_date + relativedelta(years=3)
    anniversary_5yr = hire_date + relativedelta(years=5)
    anniversary_7yr = hire_date + relativedelta(years=7)

    return {
        "service_year": service_year_display,
        "start_date": start_date,
        "end_date": end_date,
        "promote_1st": promote_1st,
        "promote_2nd": promote_2nd,
        "anniversary_1yr": anniversary_1yr,
        "anniversary_3yr": anniversary_3yr,
        "anniversary_5yr": anniversary_5yr,
        "anniversary_7yr": anniversary_7yr
    }


def get_sample_excel():
    df_sample = pd.DataFrame([
        {"부서명": "경영지원팀", "성명": "홍길동", "입사일자": "2024-01-15"},
        {"부서명": "개발팀", "성명": "김철수", "입사일자": "2023-09-10"},
        {"부서명": "영업부", "성명": "이영희", "입사일자": "2019-09-01"}
    ])
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_sample.to_excel(writer, index=False, sheet_name='신규등록')
    return output.getvalue()


# ==========================================
# 4. 세션 상태 초기화 및 콜백 정의
# ==========================================
if 'sim_date' not in st.session_state:
    st.session_state.sim_date = date.today()
if 'card_filter' not in st.session_state:
    st.session_state.card_filter = "ALL"


def reset_to_today():
    st.session_state.sim_date = date.today()


# ==========================================
# 5. 메인 화면 구성
# ==========================================
init_db()

st.markdown('<div class="main-header">🗓️ (주)케이알메딕스 연차 및 장기 근속 포상 현황 대시보드</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">주식회사 케이알메딕스 전사 임직원의 입사일 기준 연차 소멸·촉진일과 1·3·5·7년 장기근속 포상일을 자동 관리합니다.</div>',
            unsafe_allow_html=True)

tab1, tab2, tab3 = st.tabs(["📊 대시보드 & 현황 조회", "➕ 임직원 등록 (개별 / 엑셀 대량)", "⚙️ 시스템 관리"])

# ------------------------------------------
# TAB 1: 대시보드 및 현황 조회
# ------------------------------------------
with tab1:
    with get_db_connection() as conn:
        df_emp = pd.read_sql_query("SELECT * FROM employees ORDER BY id ASC", conn)

    if df_emp.empty:
        st.info("등록된 임직원 데이터가 없습니다. 상단 '임직원 등록' 탭에서 신규 등록해 보세요.")
    else:
        # 상단 검색 필터 영역
        col_f1, col_btn, col_f2, col_f3 = st.columns([1.5, 0.6, 1.5, 2])

        with col_f1:
            simulated_date = st.date_input("조회 기준일자", key="sim_date")

        with col_btn:
            st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
            st.button("📅 오늘", key="btn_today", help="오늘 날짜로 재설정합니다.", on_click=reset_to_today)

        with col_f2:
            dept_list = ["전체"] + sorted([str(d) for d in df_emp['department'].unique() if d])
            selected_dept = st.selectbox("부서 선택", dept_list)

        with col_f3:
            search_name = st.text_input("직원 이름 검색", placeholder="성명 입력...")

        # 익월 (다음 달) 계산
        next_month_date = simulated_date + relativedelta(months=1)
        next_year = next_month_date.year
        next_month = next_month_date.month

        # 기본 데이터 필터링 (부서/이름)
        filtered_emp = df_emp.copy()
        if selected_dept != "전체":
            filtered_emp = filtered_emp[filtered_emp['department'] == selected_dept]
        if search_name.strip():
            filtered_emp = filtered_emp[filtered_emp['name'].astype(str).str.contains(search_name.strip(), case=False)]

        processed_list = []
        under_1yr_count = 0
        expiring_this_month = 0
        expiring_3months = 0
        promote_this_month_count = 0
        next_month_bonus_count = 0

        for _, row in filtered_emp.iterrows():
            emp_id = row['id']
            dept = row['department']
            name = row['name']
            hire_date = row['hire_date']

            info = calculate_vacation_info(hire_date, simulated_date)
            end_date = info['end_date']
            p1 = info['promote_1st']
            p2 = info['promote_2nd']

            is_under_1yr = "1년 미만" in info['service_year']
            if is_under_1yr:
                under_1yr_count += 1

            # 연차 소멸일 조건
            is_expiring_this_month = (end_date.year == simulated_date.year and end_date.month == simulated_date.month)
            days_until_expiry = (end_date - simulated_date).days
            is_expiring_3months = (0 <= days_until_expiry <= 90) and not is_expiring_this_month

            if is_expiring_this_month:
                expiring_this_month += 1
                expiry_rank = 1
            elif is_expiring_3months:
                expiring_3months += 1
                expiry_rank = 2
            else:
                expiry_rank = 3

            # 당월 촉진일 조건
            is_p1_this_month = (p1.year == simulated_date.year and p1.month == simulated_date.month)
            is_p2_this_month = (p2.year == simulated_date.year and p2.month == simulated_date.month)
            is_promote_this_month = is_p1_this_month or is_p2_this_month

            if is_promote_this_month:
                promote_this_month_count += 1

            # 익월 근속 수당 조건 (1, 3, 5, 7년)
            a1 = info['anniversary_1yr']
            a3 = info['anniversary_3yr']
            a5 = info['anniversary_5yr']
            a7 = info['anniversary_7yr']

            is_a1_next_month = (a1.year == next_year and a1.month == next_month)
            is_a3_next_month = (a3.year == next_year and a3.month == next_month)
            is_a5_next_month = (a5.year == next_year and a5.month == next_month)
            is_a7_next_month = (a7.year == next_year and a7.month == next_month)
            is_bonus_next_month = is_a1_next_month or is_a3_next_month or is_a5_next_month or is_a7_next_month

            if is_bonus_next_month:
                next_month_bonus_count += 1

            processed_list.append({
                "emp_id": emp_id,
                "dept": dept,
                "name": name,
                "hire_date": hire_date,
                "info": info,
                "end_date": end_date,
                "p1": p1,
                "p2": p2,
                "a1": a1,
                "a3": a3,
                "a5": a5,
                "a7": a7,
                "is_under_1yr": is_under_1yr,
                "is_expiring_this_month": is_expiring_this_month,
                "is_expiring_3months": is_expiring_3months,
                "is_p1_this_month": is_p1_this_month,
                "is_p2_this_month": is_p2_this_month,
                "is_promote_this_month": is_promote_this_month,
                "is_a1_next_month": is_a1_next_month,
                "is_a3_next_month": is_a3_next_month,
                "is_a5_next_month": is_a5_next_month,
                "is_a7_next_month": is_a7_next_month,
                "is_bonus_next_month": is_bonus_next_month,
                "expiry_rank": expiry_rank
            })

        st.markdown("---")

        # 상단 요약 카드를 '클릭 가능한 필터 버튼'으로 구성
        m1, m2, m3, m4, m5, m6 = st.columns(6)

        c_filter = st.session_state.card_filter

        with m1:
            btn_type = "primary" if c_filter == "ALL" else "secondary"
            if st.button(f"👥 전체👥\n\n {len(processed_list)} 명", key="card_all", use_container_width=True,
                         type=btn_type):
                st.session_state.card_filter = "ALL"
                st.rerun()

        with m2:
            btn_type = "primary" if c_filter == "UNDER_1YR" else "secondary"
            if st.button(f"🐣 1년 미만🐣\n\n {under_1yr_count} 명", key="card_under_1yr", use_container_width=True,
                         type=btn_type):
                st.session_state.card_filter = "UNDER_1YR"
                st.rerun()

        with m3:
            btn_type = "primary" if c_filter == "EXP_THIS_MONTH" else "secondary"
            if st.button(f"🔴 당월 소멸 예정🔴\n\n {expiring_this_month} 명", key="card_exp_month", use_container_width=True,
                         type=btn_type):
                st.session_state.card_filter = "EXP_THIS_MONTH"
                st.rerun()

        with m4:
            btn_type = "primary" if c_filter == "EXP_3MONTHS" else "secondary"
            if st.button(f"🔵 3개월내 소멸🔵\n\n {expiring_3months} 명", key="card_exp_3m", use_container_width=True,
                         type=btn_type):
                st.session_state.card_filter = "EXP_3MONTHS"
                st.rerun()

        with m5:
            btn_type = "primary" if c_filter == "PROMOTE_THIS_MONTH" else "secondary"
            if st.button(f"🟣 당월 촉진 대상🟣\n\n {promote_this_month_count} 명", key="card_promote",
                         use_container_width=True, type=btn_type):
                st.session_state.card_filter = "PROMOTE_THIS_MONTH"
                st.rerun()

        with m6:
            btn_type = "primary" if c_filter == "BONUS_NEXT_MONTH" else "secondary"
            if st.button(f"🟢 익월 근속수당🟢\n\n {next_month_bonus_count} 명", key="card_bonus", use_container_width=True,
                         type=btn_type):
                st.session_state.card_filter = "BONUS_NEXT_MONTH"
                st.rerun()

        # 요약 카드 선택에 따른 데이터 2차 필터링
        final_list = []
        for item in processed_list:
            if c_filter == "ALL":
                final_list.append(item)
            elif c_filter == "UNDER_1YR" and item["is_under_1yr"]:
                final_list.append(item)
            elif c_filter == "EXP_THIS_MONTH" and item["is_expiring_this_month"]:
                final_list.append(item)
            elif c_filter == "EXP_3MONTHS" and item["is_expiring_3months"]:
                final_list.append(item)
            elif c_filter == "PROMOTE_THIS_MONTH" and item["is_promote_this_month"]:
                final_list.append(item)
            elif c_filter == "BONUS_NEXT_MONTH" and item["is_bonus_next_month"]:
                final_list.append(item)

        st.markdown("<br>", unsafe_allow_html=True)

        # 현재 필터 상태 안내 및 범례
        filter_labels = {
            "ALL": "전체 임직원 목록",
            "UNDER_1YR": "1년 미만 입사자 목록",
            "EXP_THIS_MONTH": "당월 소멸 예정자 목록",
            "EXP_3MONTHS": "3개월 내 소멸 예정자 목록",
            "PROMOTE_THIS_MONTH": "당월 촉진 대상자 목록",
            "BONUS_NEXT_MONTH": f"익월({next_month}월) 근속포상 대상자 목록"
        }

        st.markdown(f"""
        <div style="display: flex; gap: 12px; align-items: center; font-size: 0.85rem; margin-bottom: 10px; flex-wrap: wrap;">
            <span style="font-weight: bold; font-size: 1.05rem; color: #1E3A8A;">📋 {filter_labels[c_filter]} ({len(final_list)}명)</span>
            <span style="color: #4B5563; font-size: 0.85rem;">💡 상단 카드를 클릭하면 해당 조건의 대상자만 필터링됩니다.</span>
            <span style="color: #854D0E; font-weight: 700; background-color: #FEF08A; padding: 2px 8px; border-radius: 4px; border: 1.5px solid #EAB308;">🟡 관리 대상자 성명</span>
            <span style="color: #DC2626; font-weight: 700; background-color: #FEE2E2; padding: 2px 8px; border-radius: 4px; border: 1.5px solid #DC2626;">🔴 당월 소멸일</span>
            <span style="color: #2563EB; font-weight: 700; background-color: #DBEAFE; padding: 2px 8px; border-radius: 4px; border: 1.5px solid #2563EB;">🔵 3개월내 소멸일</span>
            <span style="color: #6B21A8; font-weight: 700; background-color: #F3E8FF; padding: 2px 8px; border-radius: 4px; border: 1.5px solid #9333EA;">🟣 당월 촉진일</span>
            <span style="color: #059669; font-weight: 700; background-color: #D1FAE5; padding: 2px 8px; border-radius: 4px; border: 1.5px solid #059669;">🟢 익월({next_month}월) 근속포상</span>
        </div>
        """, unsafe_allow_html=True)

        # HTML 행 구성 및 data-* 속성 부여
        rows_html = ""
        for item in final_list:
            emp_id = item["emp_id"]
            dept = item["dept"]
            name = item["name"]
            hire_date = item["hire_date"]
            info = item["info"]
            end_date = item["end_date"]
            p1 = item["p1"]
            p2 = item["p2"]
            a1 = item["a1"]
            a3 = item["a3"]
            a5 = item["a5"]
            a7 = item["a7"]

            # '1년 미만' 탭을 눌렀을 때만 이름 앞뒤에 🐣 이모티콘 추가
            if c_filter == "UNDER_1YR":
                display_name = f"🐣 {name} 🐣"
            else:
                display_name = name

            has_any_highlight = False

            # 연차 소멸일 셀
            td_end_date_class = "col-center col-bold"
            if item["is_expiring_this_month"]:
                td_end_date_class += " cell-red-highlight"
                has_any_highlight = True
            elif item["is_expiring_3months"]:
                td_end_date_class += " cell-blue-highlight"
                has_any_highlight = True

            # 촉진일 셀
            td_p1_class = "col-center col-bold"
            td_p2_class = "col-center col-bold"
            if item["is_p1_this_month"]:
                td_p1_class += " cell-purple-highlight"
                has_any_highlight = True
            if item["is_p2_this_month"]:
                td_p2_class += " cell-purple-highlight"
                has_any_highlight = True

            # 근속일 셀 (1, 3, 5, 7년)
            td_a1_class = "col-center"
            td_a3_class = "col-center"
            td_a5_class = "col-center"
            td_a7_class = "col-center"

            if item["is_a1_next_month"]:
                td_a1_class += " cell-green-highlight"
                has_any_highlight = True
            if item["is_a3_next_month"]:
                td_a3_class += " cell-green-highlight"
                has_any_highlight = True
            if item["is_a5_next_month"]:
                td_a5_class += " cell-green-highlight"
                has_any_highlight = True
            if item["is_a7_next_month"]:
                td_a7_class += " cell-green-highlight"
                has_any_highlight = True

            # 성명 셀 (관리 대상 발생 시 노란색)
            td_name_class = "col-center col-semibold"
            if has_any_highlight:
                td_name_class += " cell-yellow-highlight"

            # 정렬을 위한 색상 가중치
            color_name_score = 1 if has_any_highlight else 0
            color_expiry_score = item["expiry_rank"]
            color_p1_score = 1 if item["is_p1_this_month"] else 0
            color_p2_score = 1 if item["is_p2_this_month"] else 0
            color_a1_score = 1 if item["is_a1_next_month"] else 0
            color_a3_score = 1 if item["is_a3_next_month"] else 0
            color_a5_score = 1 if item["is_a5_next_month"] else 0
            color_a7_score = 1 if item["is_a7_next_month"] else 0

            rows_html += f"""
            <tr data-id="{emp_id}" data-dept="{dept}" data-name="{name}" data-hire="{hire_date}">
                <td class="col-center col-bold">{emp_id}</td>
                <td class="col-center col-semibold">{dept}</td>
                <td class="{td_name_class}" data-color-score="{color_name_score}">{display_name}</td>
                <td class="col-center">{hire_date}</td>
                <td class="col-center">{info['service_year']}</td>
                <td class="col-center">{info['start_date'].strftime("%Y-%m-%d")}</td>
                <td class="{td_end_date_class}" data-color-score="{color_expiry_score}">{end_date.strftime("%Y-%m-%d")}</td>
                <td class="{td_p1_class}" data-color-score="{color_p1_score}">{p1.strftime("%Y-%m-%d")}</td>
                <td class="{td_p2_class}" data-color-score="{color_p2_score}">{p2.strftime("%Y-%m-%d")}</td>
                <td class="{td_a1_class}" data-color-score="{color_a1_score}">{a1.strftime("%Y-%m-%d")}</td>
                <td class="{td_a3_class}" data-color-score="{color_a3_score}">{a3.strftime("%Y-%m-%d")}</td>
                <td class="{td_a5_class}" data-color-score="{color_a5_score}">{a5.strftime("%Y-%m-%d")}</td>
                <td class="{td_a7_class}" data-color-score="{color_a7_score}">{a7.strftime("%Y-%m-%d")}</td>
            </tr>
            """

        # 테이블 HTML & JS
        full_html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <link rel="stylesheet" as="style" crossorigin href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css" />
            <style>
                * {{
                    font-family: 'Pretendard SemiBold', -apple-system, BlinkMacSystemFont, sans-serif;
                    box-sizing: border-box;
                }}
                body {{
                    margin: 0;
                    padding: 0;
                    background-color: transparent;
                }}
                .custom-table {{
                    width: 100%;
                    border-collapse: separate;
                    border-spacing: 0 4px;
                    font-size: 0.85rem;
                }}
                .custom-table th {{
                    background-color: #F3F4F6;
                    color: #111827;
                    font-weight: 700 !important;
                    text-align: center !important;
                    padding: 10px 2px;
                    border-top: 1.5px solid #D1D5DB;
                    border-bottom: 2px solid #9CA3AF;
                    cursor: pointer;
                    user-select: none;
                    transition: background-color 0.15s ease;
                }}
                .custom-table th:hover {{
                    background-color: #E5E7EB;
                }}
                .sort-icon {{
                    font-size: 0.75rem;
                    color: #6B7280;
                    margin-left: 2px;
                }}
                .custom-table td {{
                    padding: 10px 4px;
                    background-color: #FFFFFF;
                    border-top: 1px solid #E5E7EB;
                    border-bottom: 1px solid #E5E7EB;
                    color: #111827;
                }}
                .col-center {{ text-align: center !important; }}
                .col-bold {{ font-weight: 700 !important; }}
                .col-semibold {{ font-weight: 600 !important; }}

                /* 셀 강조 스타일 */
                td.cell-yellow-highlight {{
                    border: 2px solid #EAB308 !important;
                    background-color: #FEF08A !important;
                    color: #713F12 !important;
                    font-weight: 700 !important;
                    border-radius: 6px;
                }}
                td.cell-red-highlight {{
                    border: 2px solid #EF4444 !important;
                    background-color: #FEE2E2 !important;
                    color: #B91C1C !important;
                    font-weight: 700 !important;
                    border-radius: 6px;
                }}
                td.cell-blue-highlight {{
                    border: 2px solid #3B82F6 !important;
                    background-color: #DBEAFE !important;
                    color: #1D4ED8 !important;
                    font-weight: 700 !important;
                    border-radius: 6px;
                }}
                td.cell-purple-highlight {{
                    border: 2px solid #9333EA !important;
                    background-color: #F3E8FF !important;
                    color: #6B21A8 !important;
                    font-weight: 700 !important;
                    border-radius: 6px;
                }}
                td.cell-green-highlight {{
                    border: 2px solid #10B981 !important;
                    background-color: #D1FAE5 !important;
                    color: #047857 !important;
                    font-weight: 700 !important;
                    border-radius: 6px;
                }}
            </style>
        </head>
        <body>
            <table class="custom-table" id="empTable">
                <thead>
                    <tr>
                        <th onclick="sortTable(0, 'number')" style="width: 4%;">ID <span class="sort-icon">↕️</span></th>
                        <th onclick="sortTable(1, 'string')" style="width: 8%;">부서명 <span class="sort-icon">↕️</span></th>
                        <th onclick="sortTable(2, 'color_str')" style="width: 8%;">성명 <span class="sort-icon">↕️</span></th>
                        <th onclick="sortTable(3, 'date')" style="width: 8%;">입사일자 <span class="sort-icon">↕️</span></th>
                        <th onclick="sortTable(4, 'string')" style="width: 9%;">구분/회차 <span class="sort-icon">↕️</span></th>
                        <th onclick="sortTable(5, 'date')" style="width: 8%;">연차 발생일 <span class="sort-icon">↕️</span></th>
                        <th onclick="sortTable(6, 'color_expiry')" style="width: 8%;">연차 소멸일 <span class="sort-icon">↕️</span></th>
                        <th onclick="sortTable(7, 'color_date')" style="width: 8%;">1차 촉진일 <span class="sort-icon">↕️</span></th>
                        <th onclick="sortTable(8, 'color_date')" style="width: 8%;">2차 촉진일 <span class="sort-icon">↕️</span></th>
                        <th onclick="sortTable(9, 'color_date')" style="width: 7.5%;">1년 근속 <span class="sort-icon">↕️</span></th>
                        <th onclick="sortTable(10, 'color_date')" style="width: 7.5%;">3년 근속 <span class="sort-icon">↕️</span></th>
                        <th onclick="sortTable(11, 'color_date')" style="width: 7.5%;">5년 근속 <span class="sort-icon">↕️</span></th>
                        <th onclick="sortTable(12, 'color_date')" style="width: 7.5%;">7년 근속 <span class="sort-icon">↕️</span></th>
                    </tr>
                </thead>
                <tbody id="empTableBody">
                    {rows_html if final_list else '<tr><td colspan="13" style="text-align:center; padding:20px; color:#6B7280;">선택하신 조건에 해당하는 대상자가 없습니다.</td></tr>'}
                </tbody>
            </table>

            <script>
                let sortDirections = {{}};

                function sortTable(colIndex, type) {{
                    const table = document.getElementById("empTable");
                    const tbody = document.getElementById("empTableBody");
                    const rows = Array.from(tbody.querySelectorAll("tr"));
                    const headers = table.querySelectorAll("th");

                    if (rows.length === 0 || rows[0].cells.length === 1) return;

                    let dir = sortDirections[colIndex] === "asc" ? "desc" : "asc";
                    sortDirections[colIndex] = dir;

                    headers.forEach((th, idx) => {{
                        const icon = th.querySelector(".sort-icon");
                        if (idx === colIndex) {{
                            icon.textContent = dir === "asc" ? " ▲" : " ▼";
                            th.style.backgroundColor = "#E5E7EB";
                        }} else {{
                            icon.textContent = " ↕️";
                            th.style.backgroundColor = "#F3F4F6";
                        }}
                    }});

                    rows.sort((rowA, rowB) => {{
                        const cellA = rowA.children[colIndex];
                        const cellB = rowB.children[colIndex];

                        let valA = cellA ? cellA.innerText.trim() : "";
                        let valB = cellB ? cellB.innerText.trim() : "";

                        if (type === 'color_expiry') {{
                            let scoreA = parseInt((cellA && cellA.getAttribute('data-color-score')) || '3');
                            let scoreB = parseInt((cellB && cellB.getAttribute('data-color-score')) || '3');

                            if (scoreA !== scoreB) {{
                                return dir === "asc" ? scoreA - scoreB : scoreB - scoreA;
                            }}
                            return dir === "asc" ? valA.localeCompare(valB) : valB.localeCompare(valA);
                        }}

                        if (type === 'color_str' || type === 'color_date') {{
                            let scoreA = parseInt((cellA && cellA.getAttribute('data-color-score')) || '0');
                            let scoreB = parseInt((cellB && cellB.getAttribute('data-color-score')) || '0');

                            if (scoreA !== scoreB) {{
                                return dir === "asc" ? scoreB - scoreA : scoreA - scoreB;
                            }}
                        }}

                        if (type === 'number') {{
                            let numA = parseFloat(valA) || 0;
                            let numB = parseFloat(valB) || 0;
                            return dir === "asc" ? numA - numB : numB - numA;
                        }}

                        return dir === "asc" ? valA.localeCompare(valB) : valB.localeCompare(valA);
                    }});

                    rows.forEach(row => tbody.appendChild(row));
                }}
            </script>
        </body>
        </html>
        """

        calculated_height = max(160, len(final_list) * 52 + 70)
        components.html(full_html, height=calculated_height, scrolling=False)

# ------------------------------------------
# TAB 2: 임직원 등록 (개별 / 엑셀 대량 업로드)
# ------------------------------------------
with tab2:
    sub_tab1, sub_tab2 = st.tabs(["📁 엑셀/CSV 파일로 대량 등록", "✍️ 1명씩 개별 등록"])

    with sub_tab1:
        st.subheader("📁 엑셀 파일 대량 업로드")
        st.markdown("엑셀 파일(.xlsx)이나 CSV 파일(.csv)을 업로드하여 여러 명을 한 번에 등록합니다.")

        col_desc, col_dl = st.columns([3, 1])
        with col_desc:
            st.caption("※ 엑셀 파일의 첫 번째 행에는 반드시 **`부서명`**, **`성명`**, **`입사일자`** 컬럼 헤더가 포함되어 있어야 합니다.")
            st.caption("※ 입사일자 형식 예시: `2024-05-20` 또는 `2024/05/20`")
        with col_dl:
            sample_excel = get_sample_excel()
            st.download_button(
                label="📥 표준 양식 엑셀 다운로드",
                data=sample_excel,
                file_name="임직원_등록_양식.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        st.markdown("---")

        uploaded_file = st.file_uploader("엑셀 또는 CSV 파일 선택", type=["xlsx", "xls", "csv"])

        if uploaded_file is not None:
            try:
                if uploaded_file.name.endswith('.csv'):
                    df_upload = pd.read_csv(uploaded_file)
                else:
                    df_upload = pd.read_excel(uploaded_file)

                required_cols = {"부서명", "성명", "입사일자"}
                if not required_cols.issubset(set(df_upload.columns)):
                    st.error("❌ 파일 형식이 올바르지 않습니다. 필수 컬럼(`부서명`, `성명`, `입사일자`)이 존재하는지 확인해 주세요.")
                else:
                    df_upload["입사일자"] = pd.to_datetime(df_upload["입사일자"]).dt.strftime('%Y-%m-%d')

                    st.markdown("##### 🔍 업로드 데이터 미리보기")
                    st.dataframe(df_upload[["부서명", "성명", "입사일자"]], use_container_width=True)

                    if st.button("🚀 위 데이터 전체 DB 저장하기", type="primary"):
                        with get_db_connection() as conn:
                            cursor = conn.cursor()
                            count = 0
                            for _, row in df_upload.iterrows():
                                cursor.execute(
                                    "INSERT INTO employees (department, name, hire_date) VALUES (?, ?, ?)",
                                    (str(row["부서명"]).strip(), str(row["성명"]).strip(), str(row["입사일자"]).strip())
                                )
                                count += 1
                            conn.commit()

                        st.success(f"🎉 총 {count}명의 신규 임직원이 성공적으로 일괄 등록되었습니다!")
                        st.rerun()

            except Exception as e:
                st.error(f"⚠️ 파일 처리 중 오류가 발생했습니다: {e}")

    with sub_tab2:
        st.subheader("👤 개별 임직원 정보 입력")
        col_form, _ = st.columns([2, 1])
        with col_form:
            with st.form("add_emp_form", clear_on_submit=True):
                dept_name = st.text_input("부서명", placeholder="예: 경영지원팀, 개발팀, 영업부")
                new_name = st.text_input("직원 성명", placeholder="예: 홍길동")
                new_hire_date = st.date_input("입사일자 선택", value=date.today())

                submit = st.form_submit_button("신규 임직원 저장")

                if submit:
                    if not dept_name.strip():
                        st.error("부서명을 입력해 주세요.")
                    elif not new_name.strip():
                        st.error("직원 성명을 입력해 주세요.")
                    else:
                        with get_db_connection() as conn:
                            cursor = conn.cursor()
                            cursor.execute(
                                "INSERT INTO employees (department, name, hire_date) VALUES (?, ?, ?)",
                                (dept_name.strip(), new_name.strip(), new_hire_date.strftime("%Y-%m-%d"))
                            )
                            conn.commit()
                        st.success(f"✅ '{dept_name}' 부서의 '{new_name}' 직원이 추가되었습니다!")
                        st.rerun()

# ------------------------------------------
# TAB 3: 시스템 관리
# ------------------------------------------
with tab3:
    st.subheader("🛠️ 시스템 테스트 및 데이터 관리")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### 1. 빠른 샘플 데이터 생성")
        st.caption("클릭 한 번으로 5명의 샘플 직원을 등록합니다.")
        if st.button("샘플 임직원 데이터 생성"):
            insert_sample_data()
            st.success("샘플 데이터 생성이 완료되었습니다!")
            st.rerun()

    with col2:
        st.markdown("#### 2. 데이터베이스 초기화")
        st.caption("모든 데이터를 초기 상태로 지웁니다.")
        if st.button("전체 데이터 초기화", type="primary"):
            with get_db_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("DROP TABLE IF EXISTS employees")
                conn.commit()
            init_db()
            st.warning("모든 데이터가 초기화되었습니다.")
            st.rerun()