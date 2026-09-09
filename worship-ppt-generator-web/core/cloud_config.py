import json
import os
import re
import shutil

# 텔레그램 공식 봇 인증 정보 및 채널 초대 링크 (Streamlit secrets 우선 참조)
DEFAULT_BOT_TOKEN = "8667891459:AAFGv5aNMfsj7DoIemJ-xzvHnyzey8yehq4"
DEFAULT_CHAT_ID = "-1004424616579"
DEFAULT_CHANNEL_LINK = "https://t.me/+V_sh3JAebg03ZjNl"

try:
    import streamlit as st
    BOT_TOKEN = st.secrets.get("TELEGRAM_BOT_TOKEN", DEFAULT_BOT_TOKEN) if hasattr(st, "secrets") else DEFAULT_BOT_TOKEN
    CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID", DEFAULT_CHAT_ID) if hasattr(st, "secrets") else DEFAULT_CHAT_ID
    TELEGRAM_CHANNEL_LINK = st.secrets.get("TELEGRAM_CHANNEL_LINK", DEFAULT_CHANNEL_LINK) if hasattr(st, "secrets") else DEFAULT_CHANNEL_LINK
except Exception:
    BOT_TOKEN = DEFAULT_BOT_TOKEN
    CHAT_ID = DEFAULT_CHAT_ID
    TELEGRAM_CHANNEL_LINK = DEFAULT_CHANNEL_LINK

# 기본 기준 부서 (기본값: 기타)
DEFAULT_DEPARTMENT = "기타"

# 관리자가 인가한 기본 공식 부서 목록 (텔레그램 클라우드에서 등록 전까지 초기 0개)
OFFICIAL_DEPARTMENTS = []

# 프로젝트 루트 및 데이터 보관함 4대 디렉토리 경로
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_ROOT = os.path.join(BASE_DIR, "데이터 보관함")

DIR_PRESETS = os.path.join(DATA_ROOT, "1_기준_PPT_프리셋")
DIR_PPTX = os.path.join(DATA_ROOT, "2_찬양_PPT")
DIR_SONGS = os.path.join(DATA_ROOT, "3_찬양곡_라이브러리")
DIR_CONTI = os.path.join(DATA_ROOT, "4_찬양_콘티")

# 레거시 assets 디렉토리
LEGACY_ASSETS_DIR = os.path.join(BASE_DIR, "assets")

def ensure_data_directories():
    """
    데이터 보관함 및 4대 하위 디렉토리를 생성합니다.
    """
    for d in [DATA_ROOT, DIR_PRESETS, DIR_PPTX, DIR_SONGS, DIR_CONTI]:
        os.makedirs(d, exist_ok=True)


def get_official_departments():
    """
    클라우드 및 시스템에서 공식으로 인가된 '공식 부서' 목록을 반환합니다.
    기본 공식 부서('청년부')와 텔레그램 클라우드 카탈로그(.cloud_catalog.json)의 presets에 관리자가 등록한 부서만 포함됩니다.
    """
    official = set(OFFICIAL_DEPARTMENTS)
    catalog_path = os.path.join(DATA_ROOT, ".cloud_catalog.json")
    if os.path.exists(catalog_path):
        try:
            with open(catalog_path, "r", encoding="utf-8") as f:
                cat = json.load(f)
                presets_dict = cat.get("presets", {})
                if isinstance(presets_dict, dict):
                    for k in presets_dict.keys():
                        cleaned_k = re.sub(r'[\[\]]', '', str(k)).replace("기준 PPT", "").replace("기준PPT", "").strip()
                        if cleaned_k:
                            official.add(cleaned_k)
        except Exception:
            pass
    return sorted(list(official))


def is_official_department(name):
    """지정된 이름이 인가된 공식 부서인지 여부를 판별합니다."""
    if not name:
        return False
    cleaned = re.sub(r'[\[\]]', '', str(name)).replace("기준 PPT", "").replace("기준PPT", "").strip()
    return cleaned in get_official_departments()


def normalize_department_name(name):
    """
    부서명/프리셋명을 순수 부서명으로 통일하되,
    관리자가 인가한 공식 프리셋이 아니거나 '내 PC에서 직접 업로드'인 경우
    모두 안전하게 '기타'로 엄격 분류합니다.
    
    예: '[청년부] 기준 PPT' -> '청년부' (공식 인가)
    예: '청년부' -> '청년부' (공식 인가)
    예: '[청년부  테스트] 기준 PPT' -> '기타' (비공인 로컬 임의 프리셋 차단)
    예: '내 PC에서 직접 업로드' -> '기타'
    예: '' 또는 None -> '기타'
    """
    if not name or name == "내 PC에서 직접 업로드":
        return "기타"
    cleaned = re.sub(r'[\[\]]', '', str(name)).replace("기준 PPT", "").replace("기준PPT", "").strip()
    if cleaned in get_official_departments():
        return cleaned
    return "기타"

def sanitize_filename_part(text):
    r"""
    윈도우 파일명에 부적합한 문자(\ / : * ? " < > |)를 안전하게 정제합니다.
    (대괄호 [ ] 는 윈도우에서 허용되므로 유지)
    """
    if not text:
        return "미지정"
    cleaned = re.sub(r'[\\/:*?"<>|]', '', str(text)).strip()
    return cleaned if cleaned else "미지정"

def format_standard_filename(preset_name, file_type, title, date_str, ext):
    """
    사용자 정의 표준 파일명 규칙 생성:
    [프리셋명][분류][제목][작성일자YYYYMMDD].확장자
    예: [청년부][PPT][2026.09.13 찬양 가사][20260913].pptx
    예: [기타][PPT][2026.09.13 찬양 가사][20260913].pptx (비공식 프리셋/직접 업로드 시)
    예: [청년부][찬양곡][주의 친절한 팔에 안기세][20260913].json
    """
    # 프리셋명을 공식 부서명으로 검증 및 정제 (비공인 프리셋은 자동으로 '기타' 분류)
    p_clean = normalize_department_name(preset_name)
        
    t_clean = sanitize_filename_part(title)
    f_type = sanitize_filename_part(file_type)
    
    # 날짜 정제 (YYYYMMDD 정확히 8자리)
    d_raw = re.sub(r'[^0-9]', '', str(date_str))
    d_clean = d_raw[:8] if len(d_raw) >= 8 else "20260101"
    
    ext_clean = ext.lstrip('.').lower()
    return f"[{p_clean}][{f_type}][{t_clean}][{d_clean}].{ext_clean}"

def parse_standard_filename(filename):
    """
    표준 파일명([프리셋][분류][제목][날짜].확장자)을 분석하여 메타데이터 딕셔너리를 반환합니다.
    일치하지 않는 경우 None을 반환합니다.
    """
    pattern = r'^\[(.*?)\]\[(.*?)\]\[(.*?)\]\[(\d{8,})\]\.([a-zA-Z0-9]+)$'
    match = re.match(pattern, filename.strip())
    if match:
        return {
            "preset": match.group(1),
            "type": match.group(2),
            "title": match.group(3),
            "date": match.group(4)[:8],
            "ext": match.group(5).lower(),
            "filename": filename
        }
    return None
