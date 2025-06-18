import pandas as pd
import os
from datetime import datetime, date

class EmployeeManager:
    """직원 정보 및 출장비 관리 클래스"""
    
    def __init__(self, csv_file="직급별 출장비.csv"):
        self.csv_file = csv_file
        self.employee_data = self.load_employee_data()
    
    def load_employee_data(self):
        """CSV 파일에서 직원 데이터 로드"""
        try:
            if os.path.exists(self.csv_file):
                df = pd.read_csv(self.csv_file, encoding='cp949')
                # 공백 제거 및 컬럼명 정리
                df.columns = df.columns.str.strip()
                df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
                
                # 일비, 식비에서 쉼표 제거하고 숫자로 변환
                df['일비'] = df['일비'].str.replace(',', '').str.replace(' ', '').astype(int)
                df['식비'] = df['식비'].str.replace(',', '').str.replace(' ', '').astype(int)
                
                return df
            else:
                print(f"파일을 찾을 수 없습니다: {self.csv_file}")
                return pd.DataFrame()
        except Exception as e:
            print(f"데이터 로드 오류: {e}")
            return pd.DataFrame()
    
    def get_employee_names(self):
        """전체 직원 이름 리스트 반환"""
        if not self.employee_data.empty:
            return self.employee_data['이름'].tolist()
        return []
    
    def get_employee_info(self, name):
        """특정 직원의 정보 반환"""
        if not self.employee_data.empty:
            employee = self.employee_data[self.employee_data['이름'] == name]
            if not employee.empty:
                return {
                    'name': employee.iloc[0]['이름'],
                    'position': employee.iloc[0]['직급'],
                    'daily_allowance': employee.iloc[0]['일비'],
                    'meal_cost': employee.iloc[0]['식비']
                }
        return None
    
    def calculate_trip_days(self, start_date, start_time, end_date, end_time):
        """출장일수 계산"""
        try:
            # datetime 객체 생성
            start_datetime = datetime.combine(start_date, start_time)
            end_datetime = datetime.combine(end_date, end_time)
            
            # 날짜 차이 계산
            days_diff = (end_datetime.date() - start_datetime.date()).days
            
            # 최소 1일, 당일 출장도 1일로 계산
            trip_days = max(1, days_diff + 1)
            
            return trip_days
            
        except Exception as e:
            print(f"출장일수 계산 오류: {e}")
            return 1
    
    def calculate_trip_expenses(self, employee_name, start_date, start_time, end_date, end_time):
        """출장 비용 자동 계산"""
        employee_info = self.get_employee_info(employee_name)
        if not employee_info:
            return None
        
        trip_days = self.calculate_trip_days(start_date, start_time, end_date, end_time)
        
        # 비용 계산
        daily_allowance_total = trip_days * employee_info['daily_allowance']
        meal_cost_total = trip_days * employee_info['meal_cost']
        total_cost = daily_allowance_total + meal_cost_total
        
        return {
            'employee_info': employee_info,
            'trip_days': trip_days,
            'daily_allowance_per_day': employee_info['daily_allowance'],
            'meal_cost_per_day': employee_info['meal_cost'],
            'daily_allowance_total': daily_allowance_total,
            'meal_cost_total': meal_cost_total,
            'total_cost': total_cost
        }

# 전역 인스턴스 생성
employee_manager = EmployeeManager() 