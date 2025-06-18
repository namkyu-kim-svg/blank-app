import json
import os
import pandas as pd

# 기본 데이터 파일 경로
DATA_FILE = "business_trip_data.json"
PROJECT_NAMES_FILE = "연구과제명.csv"

# 기본 데이터
DEFAULT_DATA = {
    "project_managers": ["이정석", "최태섭", "한영석", "김병모", "문성대", "김남현"],
    "destinations": [
        "고창", "해평", "서울", "부산", "인천", "울산", 
        "여수", "목포", "포항", "통영", "제주", "완도",
        "군산", "보령", "태안", "안산", "화성"
    ]
}

def load_data():
    """데이터 파일에서 데이터를 로드하거나 기본 데이터를 반환"""
    try:
        if os.path.exists(DATA_FILE):
            with open(DATA_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # 기본 데이터에 없는 키가 있으면 추가
            for key in DEFAULT_DATA:
                if key not in data:
                    data[key] = DEFAULT_DATA[key].copy()
            
            return data
        else:
            # 파일이 없으면 기본 데이터로 파일 생성
            save_data(DEFAULT_DATA)
            return DEFAULT_DATA.copy()
            
    except Exception as e:
        print(f"데이터 로드 오류: {e}")
        return DEFAULT_DATA.copy()

def save_data(data):
    """데이터를 파일에 저장"""
    try:
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"데이터 저장 오류: {e}")
        return False

def add_new_data(data_type, value, data=None):
    """새로운 데이터 추가"""
    if data is None:
        data = load_data()
    
    if data_type in data and value not in data[data_type]:
        data[data_type].append(value)
        save_data(data)
        return True
    return False

def remove_data(data_type, value, data=None):
    """데이터 제거"""
    if data is None:
        data = load_data()
    
    if data_type in data and value in data[data_type]:
        data[data_type].remove(value)
        save_data(data)
        return True
    return False

def reset_to_default():
    """기본 데이터로 초기화"""
    save_data(DEFAULT_DATA)
    return DEFAULT_DATA.copy()

# 데이터 검증 함수들
def validate_data(data):
    """데이터 유효성 검사"""
    required_keys = ["project_managers", "destinations"]
    
    for key in required_keys:
        if key not in data:
            return False, f"필수 키 '{key}'가 없습니다."
        if not isinstance(data[key], list):
            return False, f"'{key}'는 리스트여야 합니다."
        if len(data[key]) == 0:
            return False, f"'{key}'에 최소 하나의 값이 있어야 합니다."
    
    return True, "데이터가 유효합니다."

def get_data_summary(data=None):
    """데이터 요약 정보 반환"""
    if data is None:
        data = load_data()
    
    summary = {}
    for key, values in data.items():
        summary[key] = {
            "count": len(values),
            "items": values
        }
    
    return summary

def load_project_names():
    """연구과제명.csv 파일에서 연구과제명 목록을 로드"""
    try:
        if os.path.exists(PROJECT_NAMES_FILE):
            df = pd.read_csv(PROJECT_NAMES_FILE, encoding='utf-8')
            # 첫 번째 열의 데이터를 리스트로 반환 (NaN 값 제외)
            project_names = df.iloc[:, 0].dropna().tolist()
            return project_names
        else:
            print(f"'{PROJECT_NAMES_FILE}' 파일이 존재하지 않습니다.")
            return []
    except Exception as e:
        print(f"연구과제명 로드 오류: {e}")
        return []

def get_all_data():
    """모든 데이터를 통합하여 반환 (기본 데이터 + 연구과제명)"""
    data = load_data()
    project_names = load_project_names()
    
    # 연구과제명 추가
    data["project_names"] = project_names
    
    return data
