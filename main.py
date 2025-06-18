import streamlit as st
import pandas as pd
from datetime import datetime, date
import os
import hashlib
from excel_generator import create_advanced_business_trip_report, create_business_trip_application
from data_manager import load_data, save_data, get_all_data
from employee_manager import employee_manager

# 비밀번호 인증 함수
def check_password():
    """회사 직원 인증을 위한 비밀번호 체크"""
    
    # 회사 비밀번호들 (해시값으로 저장 - 보안강화)
    correct_passwords = {
        # "company2024" 의 해시값
        "company2024": "8c6976e5b5410415bde908bd4dee15dfb167a9c873fc4bb8a81f6f2ab448a918",
        # "출장보고서2024" 의 해시값  
        "출장보고서2024": "a94a8fe5ccb19ba61c4c0873d391e987982fbbd3",
        # "admin123" 의 해시값
        "admin123": "240be518fabd2724ddb6f04eeb1da5967448d7e831c08c8fa822809f74c720a9"
    }
    
    def password_entered():
        """비밀번호 확인"""
        entered_password = st.session_state["password"]
        password_hash = hashlib.sha256(entered_password.encode()).hexdigest()
        
        # 간단한 비밀번호들도 허용 (해시 변환 없이)
        simple_passwords = ["company2024", "출장보고서2024", "admin123"]
        
        if (entered_password in simple_passwords or 
            any(password_hash == correct_hash for correct_hash in correct_passwords.values())):
            st.session_state["password_correct"] = True
            st.session_state["user_authenticated"] = True
            del st.session_state["password"]  # 보안을 위해 비밀번호 제거
            st.success("✅ 인증 성공! 출장문서 시스템에 접근합니다.")
            st.rerun()  # 페이지 새로고침
        else:
            st.session_state["password_correct"] = False

    # 인증 상태 확인
    if "password_correct" not in st.session_state:
        # 최초 접속 - 로그인 화면 표시
        st.markdown("# 🔐 출장문서 자동화 시스템")
        st.markdown("### 🏢 회사 직원 전용 시스템")
        
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.info("🔒 회사 직원만 접근 가능합니다.\n비밀번호를 입력해주세요.")
            
            st.text_input(
                "🔑 비밀번호", 
                type="password", 
                on_change=password_entered, 
                key="password",
                placeholder="회사 비밀번호를 입력하세요",
                help="비밀번호를 모르시면 관리자에게 문의하세요."
            )
            
            # 힌트 표시 (개발/테스트용)
            with st.expander("💡 비밀번호 힌트 (테스트용)"):
                st.code("company2024 또는 출장보고서2024 또는 admin123")
        
        if "password_correct" in st.session_state and not st.session_state["password_correct"]:
            st.error("❌ 잘못된 비밀번호입니다. 회사 담당자에게 문의하세요.")
        
        # 회사 정보 표시
        st.markdown("---")
        st.markdown("### 📋 시스템 소개")
        st.markdown("""
        - 🎯 **목적**: 출장신청서 및 출장복명서 자동 생성
        - 👥 **대상**: 회사 직원 전용
        - 🔒 **보안**: 인증된 사용자만 접근 가능
        - 📄 **기능**: Excel 문서 자동 생성 및 다운로드
        """)
        
        return False
        
    elif not st.session_state.get("password_correct", False):
        # 비밀번호 틀림
        st.error("❌ 접근 권한이 없습니다.")
        if st.button("🔄 다시 로그인"):
            for key in st.session_state.keys():
                del st.session_state[key]
            st.rerun()
        return False
    
    else:
        # 인증 성공 - 사이드바에 로그아웃 버튼 추가
        with st.sidebar:
            st.success("✅ 인증됨")
            if st.button("🚪 로그아웃"):
                for key in st.session_state.keys():
                    del st.session_state[key]
                st.rerun()
        return True

# 메인 앱 실행 전 인증 체크
if not check_password():
    st.stop()

# 페이지 설정
st.set_page_config(
    page_title="📋 출장문서 자동화 시스템",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 메인 제목
st.title("📋 출장문서 자동화 시스템")
st.markdown("---")

# 탭 구성
tab1, tab2 = st.tabs(["📝 출장신청서", "📋 출장복명서"])

# 사이드바 - 기본 정보 설정
st.sidebar.header("⚙️ 시스템 설정")

# 데이터 로드 (연구과제명 포함)
data = get_all_data()

# ============================================================================
# 탭 1: 출장신청서
# ============================================================================
with tab1:
    st.header("📝 출장신청서 작성")
    st.info("💡 출장 전에 미리 작성하는 신청서입니다.")
    
    # 기본 정보
    col_app1, col_app2 = st.columns(2)
    
    with col_app1:
        # 과제책임자
        app_project_manager_option = st.selectbox(
            "과제책임자 (선택 또는 직접 입력)",
            options=["직접 입력"] + data['project_managers'],
            key="app_project_manager_option"
        )
        
        if app_project_manager_option == "직접 입력":
            app_project_manager = st.text_input(
                "과제책임자 직접 입력",
                placeholder="과제책임자 이름을 입력하세요",
                key="app_project_manager_custom"
            )
        else:
            app_project_manager = app_project_manager_option
    
    with col_app2:
        # 출장지
        app_destination_option = st.selectbox(
            "🗺️ 출장지 (선택 또는 직접 입력)",
            options=["직접 입력"] + data['destinations'],
            key="app_destination_option"
        )
        
        if app_destination_option == "직접 입력":
            app_destination = st.text_input(
                "출장지 직접 입력",
                placeholder="출장지를 입력하세요",
                key="app_destination_custom"
            )
        else:
            app_destination = app_destination_option
    
    # 연구과제명
    app_project_name_option = st.selectbox(
        "📋 연구과제명 (선택 또는 직접 입력)",
        options=["직접 입력"] + data.get('project_names', []),
        key="app_project_name_option"
    )
    
    if app_project_name_option == "직접 입력":
        app_project_name = st.text_input(
            "연구과제명 직접 입력",
            placeholder="연구과제명을 입력하세요",
            key="app_project_name_custom"
        )
    else:
        app_project_name = app_project_name_option
    
    # 출장 기간
    col_app_date1, col_app_date2 = st.columns(2)
    
    with col_app_date1:
        app_start_date = st.date_input(
            "출장 시작일",
            value=date.today(),
            key="app_start_date"
        )
    
    with col_app_date2:
        app_end_date = st.date_input(
            "출장 종료일",
            value=date.today(),
            key="app_end_date"
        )
    
    # 출장기간 문자열 생성
    if app_start_date and app_end_date:
        trip_duration = (app_end_date - app_start_date).days + 1
        app_trip_period = f"{app_start_date.strftime('%Y.%m.%d')} ~ {app_end_date.strftime('%m.%d')}({trip_duration-1}박 {trip_duration}일)"
    else:
        app_trip_period = ""
    
    st.text_input("📅 출장기간 (자동생성)", value=app_trip_period, disabled=True)
    
    # 출장 목적
    app_trip_purpose = st.text_area(
        "출장 목적",
        placeholder="출장 목적을 자세히 입력하세요",
        height=100,
        key="app_trip_purpose"
    )
    
    # 교통 정보
    st.markdown("---")
    st.subheader("🚗 교통 정보")
    
    col_transport1, col_transport2 = st.columns(2)
    
    with col_transport1:
        # 법인차량 목록
        company_vehicles = [
            "",
            "봉고3(탑차) - 83나9834",
            "스타렉스 - 81서0873", 
            "렉스턴 - 86주8548",
            "XM3 - 317거1539",
            "팰리세이드 - 332버2092",
            "K3 - 100소8118",
            "포터2 - 808라5997"
        ]
        
        app_company_car_option = st.selectbox(
            "법인차량 선택",
            options=["직접 입력"] + company_vehicles,
            key="app_company_car_option"
        )
        
        if app_company_car_option == "직접 입력":
            app_company_car = st.text_input(
                "법인차량 직접 입력",
                placeholder="예: 스타렉스",
                key="app_company_car_custom"
            )
        else:
            app_company_car = app_company_car_option
    
    with col_transport2:
        app_public_transport = st.selectbox(
            "대중교통",
            options=["", "항공", "철도", "선박", "기타"],
            key="app_public_transport"
        )
    
    # 출장자 관리
    st.markdown("---")
    st.subheader("👥 출장자 관리")
    
    # 세션 상태 초기화 (출장신청서용)
    if 'app_travelers_list' not in st.session_state:
        st.session_state.app_travelers_list = []
    
    # 출장자 추가 폼
    with st.expander("➕ 출장자 추가", expanded=True):
        employee_names = employee_manager.get_employee_names()
        
        if employee_names:
            app_selected_employee = st.selectbox(
                "직원 선택",
                options=employee_names,
                key="app_selected_employee"
            )
            
            # 선택된 직원 정보 표시
            if app_selected_employee:
                emp_info = employee_manager.get_employee_info(app_selected_employee)
                if emp_info:
                    col_emp1, col_emp2 = st.columns(2)
                    
                    with col_emp1:
                        st.metric("직급", emp_info['position'])
                    with col_emp2:
                        app_account = st.text_input(
                            "계좌번호",
                            placeholder="급여 계좌 또는 계좌번호 입력",
                            key="app_account"
                        )
                    
                    app_note = st.text_input(
                        "비고",
                        placeholder="특이사항이 있으면 입력하세요",
                        key="app_note"
                    )
                    
                    if st.button("✅ 출장자 추가", key="add_app_traveler"):
                        new_traveler = {
                            'position': emp_info['position'],
                            'name': emp_info['name'],
                            'account': app_account if app_account else "급여 계좌",
                            'note': app_note
                        }
                        
                        # 중복 체크
                        existing_names = [t['name'] for t in st.session_state.app_travelers_list]
                        if new_traveler['name'] not in existing_names:
                            st.session_state.app_travelers_list.append(new_traveler)
                            st.success(f"✅ {new_traveler['name']}({new_traveler['position']}) 추가됨")
                            st.rerun()
                        else:
                            st.warning(f"⚠️ {new_traveler['name']}는 이미 추가되었습니다.")
    
    # 현재 출장자 목록 표시
    if st.session_state.app_travelers_list:
        st.subheader("📋 현재 출장자 목록")
        
        for idx, traveler in enumerate(st.session_state.app_travelers_list):
            col_trav1, col_trav2, col_trav3 = st.columns([3, 3, 1])
            
            with col_trav1:
                st.write(f"**{traveler['name']}** ({traveler['position']})")
            with col_trav2:
                st.write(f"계좌: {traveler['account']}")
            with col_trav3:
                if st.button("🗑️", key=f"remove_app_traveler_{idx}"):
                    st.session_state.app_travelers_list.pop(idx)
                    st.rerun()
    
    # 출장신청서 생성 버튼
    st.markdown("---")
    
    if st.button("📄 출장신청서 생성", type="primary", key="generate_application"):
        if not app_project_name:
            st.error("❌ 연구과제명을 입력해주세요.")
        elif not app_trip_purpose:
            st.error("❌ 출장 목적을 입력해주세요.")
        elif not st.session_state.app_travelers_list:
            st.error("❌ 최소 1명의 출장자를 추가해주세요.")
        else:
            try:
                # 출장신청서 데이터 준비
                application_data = {
                    'project_manager': app_project_manager,
                    'project_name': app_project_name,
                    'trip_period': app_trip_period,
                    'destination': app_destination,
                    'trip_purpose': app_trip_purpose,
                    'company_car': app_company_car,
                    'public_transport': app_public_transport,
                    'travelers': st.session_state.app_travelers_list
                }
                
                # 파일명 생성
                current_date = datetime.now().strftime('%Y%m%d')
                safe_destination = app_destination.replace(' ', '_').replace('/', '_')
                filename = f"출장신청서_{safe_destination}_{current_date}.xlsx"
                
                # 출장신청서 생성
                file_path = create_business_trip_application(application_data, filename)
                
                st.success(f"✅ 출장신청서가 성공적으로 생성되었습니다!")
                st.info(f"📁 파일 위치: {file_path}")
                
                # 파일 다운로드 링크
                if os.path.exists(file_path):
                    with open(file_path, "rb") as file:
                        st.download_button(
                            label="💾 출장신청서 다운로드",
                            data=file.read(),
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                
            except Exception as e:
                st.error(f"❌ 출장신청서 생성 중 오류가 발생했습니다: {str(e)}")

# ============================================================================
# 탭 2: 출장복명서  
# ============================================================================
with tab2:
    # 자동 계산 출장보고서
    st.header("📋 출장복명서 작성")
    st.info("💡 직급별 출장비.csv 파일의 데이터를 기반으로 자동 계산됩니다.")
    
    # 기본 정보
    col_basic1, col_basic2 = st.columns(2)
    
    with col_basic1:
        # 과제책임자
        project_manager_option = st.selectbox(
            "과제책임자 (선택 또는 직접 입력)",
            options=["직접 입력"] + data['project_managers'],
            key="report_project_manager_option"
        )
        
        if project_manager_option == "직접 입력":
            project_manager = st.text_input(
                "과제책임자 직접 입력",
                placeholder="과제책임자 이름을 입력하세요",
                key="report_project_manager_custom"
            )
        else:
            project_manager = project_manager_option
    
    with col_basic2:
        # 출장지
        destination_option = st.selectbox(
            "🗺️ 출장지 (선택 또는 직접 입력)",
            options=["직접 입력"] + data['destinations'],
            key="report_destination_option"
        )
        
        if destination_option == "직접 입력":
            destination = st.text_input(
                "출장지 직접 입력",
                placeholder="출장지를 입력하세요",
                key="report_destination_custom"
            )
        else:
            destination = destination_option
    
    # 연구과제명 - 드롭다운 + 직접입력 가능
    project_name_option = st.selectbox(
        "📋 연구과제명 (선택 또는 직접 입력)",
        options=["직접 입력"] + data.get('project_names', []),
        key="project_name_option"
    )
    
    if project_name_option == "직접 입력":
        project_name = st.text_input(
            "연구과제명 직접 입력",
            placeholder="연구과제명을 입력하세요",
            key="report_project_name_custom"
        )
    else:
        project_name = project_name_option
    
    # 출장 기간
    col_date1, col_date2 = st.columns(2)
    
    with col_date1:
        start_date = st.date_input(
            "출장 시작일",
            value=date.today(),
            key="report_start_date"
        )
        
        start_time = st.time_input(
            "출발시간",
            key="report_start_time"
        )
    
    with col_date2:
        end_date = st.date_input(
            "출장 종료일",
            value=date.today(),
            key="report_end_date"
        )
        
        end_time = st.time_input(
            "도착시간",
            key="report_end_time"
        )
    
    # 출장 목적
    trip_purpose = st.text_area(
        "출장 목적 / 결과",
        placeholder="출장 목적과 결과를 입력하세요",
        height=100,
        key="report_trip_purpose"
    )
    
    st.markdown("---")
    
    # 추가 수당 정보
    st.subheader("➕ 추가 수당 (해당시)")
    
    col_extra1, col_extra2, col_extra3 = st.columns(3)
    
    with col_extra1:
        holiday_work = st.checkbox("휴일출장 (토,일요일,공휴일)", key="report_holiday_work")
        holiday_days = st.selectbox(
            "휴일출장 일수",
            options=[0, 1, 2, 3, 4, 5],
            disabled=not holiday_work,
            key="report_holiday_days"
        )
    
    with col_extra2:
        special_work = st.checkbox("특수출장 (선박탑승,해양조사)", key="report_special_work")
        special_days = st.selectbox(
            "특수출장 일수",
            options=[0, 1, 2, 3, 4, 5],
            disabled=not special_work,
            key="report_special_days"
        )
    
    with col_extra3:
        dangerous_work = st.checkbox("위험출장 (다이빙,해저조사)", key="report_dangerous_work")
        dangerous_days = st.selectbox(
            "위험출장 일수",
            options=[0, 1, 2, 3, 4, 5],
            disabled=not dangerous_work,
            key="report_dangerous_days"
        )
    
    if special_work or dangerous_work:
        st.info("💡 특수/위험 출장시 사진파일 첨부가 필요합니다.")
    
    st.markdown("---")
    
    # 출장자 관리 섹션
    st.subheader("👥 출장자 관리")
    
    # 세션 상태 초기화
    if 'employees_list' not in st.session_state:
        st.session_state.employees_list = []
    
    if 'additional_costs' not in st.session_state:
        st.session_state.additional_costs = []
    
    # 출장자와 추가 비용을 나란히 배치
    col_employees, col_costs = st.columns(2)
    
    with col_employees:
        # 출장자 추가 폼
        with st.expander("➕ 출장자 추가", expanded=True):
            employee_names = employee_manager.get_employee_names()
            
            if employee_names:
                selected_employee = st.selectbox(
                    "직원 선택",
                    options=employee_names,
                    key="selected_employee"
                )
                
                # 선택된 직원 정보 표시
                if selected_employee:
                    emp_info = employee_manager.get_employee_info(selected_employee)
                    if emp_info:
                        col_info1, col_info2 = st.columns(2)
                        
                        with col_info1:
                            st.metric("직급", emp_info['position'])
                            st.metric("일비", f"{emp_info['daily_allowance']:,}원")
                        with col_info2:
                            st.metric("식비", f"{emp_info['meal_cost']:,}원")
                        
                        # 출장일수 자동 계산 및 표시
                        trip_days = employee_manager.calculate_trip_days(
                            start_date, start_time, end_date, end_time
                        )
                        
                        st.info(f"📅 계산된 출장일수: **{trip_days}일**")
                        
                        # 예상 비용 계산
                        daily_total = trip_days * emp_info['daily_allowance']
                        meal_total = trip_days * emp_info['meal_cost']
                        total_cost = daily_total + meal_total
                        
                        col_cost1, col_cost2 = st.columns(2)
                        with col_cost1:
                            st.metric("일비 총액", f"{daily_total:,}원")
                        with col_cost2:
                            st.metric("식비 총액", f"{meal_total:,}원")
                        
                        st.metric("총 비용", f"{total_cost:,}원")
                        
                        # 추가 버튼
                        if st.button("👤 출장자 추가", type="primary"):
                            # 중복 확인
                            existing_names = [emp['employee_name'] for emp in st.session_state.employees_list]
                            if selected_employee not in existing_names:
                                employee_data = {
                                    'employee_name': selected_employee,
                                    'position': emp_info['position'],
                                    'trip_days': trip_days,
                                    'daily_allowance_per_day': emp_info['daily_allowance'],
                                    'meal_cost_per_day': emp_info['meal_cost'],
                                    'daily_allowance_total': daily_total,
                                    'meal_cost_total': meal_total,
                                    'project_manager': project_manager,
                                    'project_name': project_name,
                                    'destination': destination,
                                    'start_date': start_date,
                                    'start_time': start_time,
                                    'end_date': end_date,
                                    'end_time': end_time,
                                    'trip_purpose': trip_purpose,
                                    'holiday_work': holiday_work,
                                    'holiday_days': holiday_days if holiday_work else 0,
                                    'special_work': special_work,
                                    'special_days': special_days if special_work else 0,
                                    'dangerous_work': dangerous_work,
                                    'dangerous_days': dangerous_days if dangerous_work else 0
                                }
                                st.session_state.employees_list.append(employee_data)
                                st.success(f"✅ {selected_employee} 추가되었습니다!")
                            else:
                                st.warning("이미 추가된 직원입니다.")
            else:
                st.error("직급별 출장비.csv 파일을 확인해주세요.")

        with col_costs:
            # 추가 비용 입력 폼
            with st.expander("💰 추가 비용 입력", expanded=True):
                # 일반적인 비용 항목 리스트
                common_cost_items = [
                    "숙박비", "주유비", "재료비", "교통비", "식비", "통신비", 
                    "소모품비", "회의비", "주차비", "기타"
                ]
                
                with st.form("cost_form", clear_on_submit=True):
                    # 비용 항목 선택 방식
                    cost_item_type = st.radio(
                        "비용 항목 입력 방식",
                        ["드롭다운 선택", "직접 입력"],
                        horizontal=True
                    )
                    
                    if cost_item_type == "드롭다운 선택":
                        cost_item = st.selectbox(
                            "비용 항목 선택",
                            options=common_cost_items
                        )
                    else:
                        cost_item = st.text_input(
                            "비용 항목 직접 입력",
                            placeholder="예: 숙박비, 주유비, 재료비"
                        )
                    
                    payment_method = st.text_area(
                        "결제 방식",
                        placeholder="예: 법인카드\n(3619)",
                        height=60
                    )
                    
                    # 금액 입력
                    cost_amount = st.number_input(
                        "금액 (원)",
                        min_value=0,
                        value=0,
                        step=1000,
                        format="%d"
                    )
                    
                    # 실시간 금액 표시
                    if cost_amount > 0:
                        st.success(f"💰 입력 금액: **{cost_amount:,}원**")
                    
                    # 선택된 항목의 현재 상태 표시
                    if cost_item and st.session_state.additional_costs:
                        current_item_costs = [cost for cost in st.session_state.additional_costs if cost['item'] == cost_item]
                        if current_item_costs:
                            current_total = sum(cost['amount'] for cost in current_item_costs)
                            st.info(f"📊 현재 **{cost_item}** 누적: {len(current_item_costs)}건, 총 {current_total:,}원")
                    
                    # 등록 버튼
                    submitted = st.form_submit_button("💰 추가 비용 등록", type="secondary")
                    
                    if submitted:
                        if cost_item and cost_amount > 0:
                            cost_data = {
                                'item': cost_item,
                                'payment_method': payment_method,
                                'amount': cost_amount
                            }
                            st.session_state.additional_costs.append(cost_data)
                            
                            # 해당 항목의 누적 정보 계산
                            same_item_costs = [cost for cost in st.session_state.additional_costs if cost['item'] == cost_item]
                            total_for_item = sum(cost['amount'] for cost in same_item_costs)
                            
                            st.success(f"✅ {cost_item} {cost_amount:,}원 추가되었습니다!")
                            st.info(f"📈 {cost_item} 현재 총액: {total_for_item:,}원 ({len(same_item_costs)}건)")
                        else:
                            st.warning("항목명과 금액을 입력해주세요.")

    # 현재 추가된 출장자 목록
    if st.session_state.employees_list:
        st.subheader("📋 추가된 출장자 목록")
        
        total_employee_cost = 0
        for i, emp in enumerate(st.session_state.employees_list):
            emp_total = emp['daily_allowance_total'] + emp['meal_cost_total']
            total_employee_cost += emp_total
            
            col1, col2, col3 = st.columns([3, 2, 1])
            
            with col1:
                st.write(f"**{emp['employee_name']}** ({emp['position']})")
                st.write(f"일비: {emp['daily_allowance_total']:,}원 | 식비: {emp['meal_cost_total']:,}원")
            
            with col2:
                st.metric("개인 총액", f"{emp_total:,}원")
            
            with col3:
                if st.button("🗑️", key=f"remove_emp_{i}", help="제거"):
                    st.session_state.employees_list.pop(i)
                    st.rerun()

    # 현재 추가된 비용 목록
    if st.session_state.additional_costs:
        st.subheader("💰 추가된 비용 목록")
        
        # 비용 항목별로 그룹화
        cost_groups = {}
        for i, cost in enumerate(st.session_state.additional_costs):
            item_name = cost['item']
            if item_name not in cost_groups:
                cost_groups[item_name] = []
            cost_groups[item_name].append({'cost': cost, 'index': i})
        
        total_additional_cost = 0
        
        # 각 비용 항목별로 표시
        for item_name, cost_list in cost_groups.items():
            with st.expander(f"📋 {item_name} ({len(cost_list)}건)", expanded=True):
                item_total = 0
                
                for cost_info in cost_list:
                    cost = cost_info['cost']
                    index = cost_info['index']
                    item_total += cost['amount']
                    
                    col1, col2, col3 = st.columns([3, 2, 1])
                    
                    with col1:
                        st.write(f"**{cost['item']}** #{cost_list.index(cost_info) + 1}")
                        st.write(f"결제: {cost['payment_method']}")
                    
                    with col2:
                        st.write(f"금액: {cost['amount']:,}원")
                    
                    with col3:
                        if st.button("🗑️", key=f"remove_cost_{index}", help="개별 항목 제거"):
                            st.session_state.additional_costs.pop(index)
                            st.rerun()
                
                # 항목별 합계 표시
                st.info(f"💰 **{item_name} 총액**: {item_total:,}원")
                total_additional_cost += item_total

    # 전체 총액 및 생성 버튼
    if st.session_state.employees_list or st.session_state.additional_costs:
        st.markdown("---")
        
        total_all_cost = 0
        if st.session_state.employees_list:
            total_all_cost += sum(emp['daily_allowance_total'] + emp['meal_cost_total'] 
                                 for emp in st.session_state.employees_list)
        if st.session_state.additional_costs:
            total_all_cost += sum(cost['amount'] for cost in st.session_state.additional_costs)
        
        st.metric("🏆 전체 총 비용", f"{total_all_cost:,}원")
        
        # 출장보고서 생성 버튼
        if st.button("📋 출장복명서 생성", type="primary", use_container_width=True):
            if not all([destination, project_name, project_manager]):
                st.error("필수 정보를 모두 입력해주세요!")
            elif not st.session_state.employees_list:
                st.error("최소 1명의 출장자를 추가해주세요!")
            else:
                try:
                    # 파일명 생성
                    filename = f"출장복명서_{destination}_{start_date.strftime('%Y%m%d')}.xlsx"
                    
                    # 엑셀 파일 생성
                    output_path = create_advanced_business_trip_report(
                        st.session_state.employees_list, 
                        st.session_state.additional_costs,
                        filename
                    )
                    
                    st.success(f"✅ 출장복명서가 생성되었습니다!")
                    st.info(f"📁 파일 위치: {output_path}")
                    
                    # 다운로드 버튼
                    with open(output_path, "rb") as file:
                        st.download_button(
                            label="📥 파일 다운로드",
                            data=file.read(),
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    
                except Exception as e:
                    st.error(f"❌ 오류 발생: {str(e)}")
        
        # 목록 초기화 버튼
        col_reset1, col_reset2 = st.columns(2)
        with col_reset1:
            if st.button("🔄 출장자 목록 초기화", type="secondary"):
                st.session_state.employees_list = []
                st.rerun()
        
        with col_reset2:
            if st.button("🔄 비용 목록 초기화", type="secondary"):
                st.session_state.additional_costs = []
                st.rerun()

    else:
        st.info("👆 위에서 출장자를 추가해주세요.")

# 사이드바 - 데이터 관리
st.sidebar.markdown("---")
st.sidebar.header("📚 데이터 관리")

# 데이터 추가 섹션
with st.sidebar.expander("➕ 새 데이터 추가"):
    data_type = st.selectbox(
        "데이터 유형",
        ["project_managers", "destinations", "project_names"]
    )
    
    new_value = st.text_input("새 값 입력")
    
    if st.button("추가", key="add_data"):
        if new_value and new_value not in data[data_type]:
            data[data_type].append(new_value)
            save_data(data)
            st.success(f"'{new_value}' 추가됨!")
            st.rerun()
        elif new_value in data[data_type]:
            st.warning("이미 존재하는 값입니다.")
        else:
            st.warning("값을 입력해주세요.")

# 현재 데이터 표시
with st.sidebar.expander("📋 현재 데이터 보기"):
    for key, values in data.items():
        st.write(f"**{key}**: {len(values)}개")
        st.write(", ".join(values))

# 데이터 초기화 버튼
st.sidebar.markdown("---")
if st.sidebar.button("🔄 데이터 초기화", type="secondary"):
    from data_manager import reset_to_default
    reset_to_default()
    st.sidebar.success("데이터가 초기화되었습니다!")
    st.rerun() 