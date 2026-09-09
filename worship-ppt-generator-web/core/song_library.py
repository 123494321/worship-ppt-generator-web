import os
import json
import re

from core.cloud_config import (
    DIR_SONGS, 
    ensure_data_directories, 
    format_standard_filename, 
    parse_standard_filename,
    normalize_department_name,
)

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ensure_data_directories()
SONG_LIBRARY_DIR = DIR_SONGS

def ensure_library_dir():
    """찬양 라이브러리 디렉토리가 없으면 생성합니다."""
    ensure_data_directories()
    if not os.path.exists(SONG_LIBRARY_DIR):
        os.makedirs(SONG_LIBRARY_DIR, exist_ok=True)

def sanitize_filename(name):
    """폴더 및 파일명에 사용할 수 없는 특수문자를 안전하게 치환합니다."""
    return re.sub(r'[\\/*?:"<>|]', '', name).strip()

def get_all_library_songs():
    """
    데이터 보관함/3_찬양곡_라이브러리/ 내의 모든 찬양곡과 해당 곡의 날짜별 버전 목록을 반환합니다.
    """
    ensure_library_dir()
    songs = {}
    
    for song_folder in os.listdir(SONG_LIBRARY_DIR):
        folder_path = os.path.join(SONG_LIBRARY_DIR, song_folder)
        if os.path.isdir(folder_path):
            versions = []
            for file_name in os.listdir(folder_path):
                if file_name.endswith('.json'):
                    file_path = os.path.join(folder_path, file_name)
                    try:
                        with open(file_path, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                            # 표준 파일명 파싱 시도 ([프리셋][찬양곡][제목][날짜].json)
                            meta = parse_standard_filename(file_name)
                            dept = meta.get("preset", data.get("department", "청년부")) if meta else data.get("department", "청년부")
                            parsed_date = meta.get("date", data.get("date", file_name[:-5])) if meta else data.get("date", file_name[:-5])
                            
                            versions.append({
                                "date": parsed_date,
                                "routine": data.get("routine_raw", ""),
                                "title": data.get("title", song_folder),
                                "department": dept,
                                "file_path": file_path,
                                "data": data
                            })
                    except Exception:
                        pass
            if versions:
                versions.sort(key=lambda x: str(x["date"]), reverse=True)
                songs[song_folder] = versions
                
    return songs

def search_library(query):
    """
    검색어(곡 제목 일부)와 일치하는 찬양곡들을 반환합니다.
    """
    if not query or not query.strip():
        return get_all_library_songs()
        
    q = query.strip().lower()
    all_songs = get_all_library_songs()
    matched = {}
    
    for song_title, versions in all_songs.items():
        if q in song_title.lower():
            matched[song_title] = versions
            
    return matched

def get_song_version(song_title, date_str):
    """
    특정 곡의 특정 날짜 버전을 로드합니다.
    표준 파일명([부서][찬양곡][제목][날짜].json) 및 구형 파일명({날짜}.json)을 모두 지원합니다.
    """
    safe_title = sanitize_filename(song_title)
    safe_date = sanitize_filename(str(date_str)).replace(".", "").replace("-", "")
    song_folder = os.path.join(SONG_LIBRARY_DIR, safe_title)
    
    if not os.path.exists(song_folder):
        return None
        
    # 1. 구형 파일명 직접 검사 ({날짜}.json)
    target_file = os.path.join(song_folder, f"{safe_date}.json")
    if os.path.exists(target_file):
        try:
            with open(target_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
            
    # 2. 폴더 내 표준 파일명 및 JSON 데이터 내부 일자 전수 매칭
    for fname in os.listdir(song_folder):
        if fname.endswith('.json'):
            fpath = os.path.join(song_folder, fname)
            meta = parse_standard_filename(fname)
            if meta and meta.get("date") == safe_date:
                try:
                    with open(fpath, 'r', encoding='utf-8') as f:
                        return json.load(f)
                except Exception:
                    pass
            # 파일 내부 JSON date 필드 검사 fallback
            try:
                with open(fpath, 'r', encoding='utf-8') as f:
                    d = json.load(f)
                    d_clean = str(d.get("date", "")).replace(".", "").replace("-", "")
                    if d_clean == safe_date or str(d.get("date")) == str(date_str):
                        return d
            except Exception:
                pass
                
    return None

def search_library_with_version_limit(query, max_versions=3):
    """
    검색어와 일치하는 찬양곡들을 반환하되,
    곡별로 최신 max_versions개(기본 3개) 버전만 우선 필터링하여 반환합니다.
    """
    all_matched = search_library(query)
    limited = {}
    for title, versions in all_matched.items():
        limited[title] = versions[:max_versions]
    return limited

def normalize_slides_list(slides):
    """슬라이드 텍스트 목록의 공백 및 개행을 정규화합니다."""
    if isinstance(slides, str):
        slides = [slides]
    normalized = []
    for s in slides:
        lines = [line.strip() for line in str(s).split('\n') if line.strip()]
        cleaned = "\n".join(lines)
        if cleaned:
            normalized.append(cleaned)
    return normalized

def normalize_song_parts(parts):
    """찬양곡의 파트별 가사 슬라이드를 표준 비교 형식으로 정규화합니다."""
    if not isinstance(parts, dict):
        return {}
    norm = {}
    for part_name, slides in parts.items():
        norm_key = re.sub(r'[^A-Za-z0-9가-힣]', '', str(part_name)).upper()
        norm_slides = normalize_slides_list(slides)
        if norm_slides:
            norm[norm_key] = norm_slides
    return norm

def normalize_routine_text(routine_str):
    """루틴(송폼) 문자열을 표준 토큰 열로 정규화합니다."""
    if not routine_str:
        return ""
    tokens = [t.strip().upper() for t in re.split(r'[\s\-–—,>/]+', str(routine_str)) if t.strip()]
    return "-".join(tokens)

def is_song_content_identical(data1, data2):
    """
    두 찬양곡 데이터의 내용(파트별 가사 슬라이드 구성 및 루틴/송폼)이 완전히 동일한지 비교합니다.
    """
    # 1. 루틴(송폼) 비교
    r1 = normalize_routine_text(data1.get("routine_raw", ""))
    r2 = normalize_routine_text(data2.get("routine_raw", ""))
    if r1 != r2:
        return False
        
    # 2. 파트별 가사 슬라이드 구성 비교
    p1 = normalize_song_parts(data1.get("parts", {}))
    p2 = normalize_song_parts(data2.get("parts", {}))
    
    if set(p1.keys()) != set(p2.keys()):
        return False
        
    for k in p1:
        if p1[k] != p2[k]:
            return False
            
    return True

def prune_excess_song_versions(song_dir, max_versions=3):
    """
    한 곡의 폴더 내에서 버전 수가 max_versions(기본 3개)를 초과할 경우,
    가장 오래된 버전 파일부터 안전하게 정리합니다.
    """
    if not os.path.exists(song_dir):
        return
    files = [f for f in os.listdir(song_dir) if f.endswith('.json')]
    if len(files) <= max_versions:
        return
        
    parsed_files = []
    for f in files:
        fp = os.path.join(song_dir, f)
        try:
            with open(fp, 'r', encoding='utf-8') as jf:
                data = json.load(jf)
            date_val = str(data.get("date", f))
        except Exception:
            date_val = f
        parsed_files.append((date_val, fp))
        
    parsed_files.sort(key=lambda x: x[0], reverse=True)
    
    for _, old_fp in parsed_files[max_versions:]:
        try:
            os.remove(old_fp)
        except Exception:
            pass

def deduplicate_all_library_songs():
    """
    데이터 보관함/3_찬양곡_라이브러리/ 내의 모든 곡 폴더를 순회하며,
    내용(가사 슬라이드 및 루틴)이 100% 동일한 중복 버전이 있다면
    가장 최신 날짜의 파일만 남기고 구형 중복 파일들을 안전하게 정리합니다.
    """
    ensure_library_dir()
    total_pruned = 0
    if not os.path.exists(SONG_LIBRARY_DIR):
        return 0
        
    for song_name in os.listdir(SONG_LIBRARY_DIR):
        song_dir = os.path.join(SONG_LIBRARY_DIR, song_name)
        if not os.path.isdir(song_dir):
            continue
            
        files = sorted([f for f in os.listdir(song_dir) if f.endswith('.json')])
        if len(files) <= 1:
            continue
            
        to_delete = set()
        for i in range(len(files)):
            if files[i] in to_delete:
                continue
            fp_i = os.path.join(song_dir, files[i])
            try:
                with open(fp_i, 'r', encoding='utf-8') as f:
                    d_i = json.load(f)
            except Exception:
                continue
                
            for j in range(i + 1, len(files)):
                if files[j] in to_delete:
                    continue
                fp_j = os.path.join(song_dir, files[j])
                try:
                    with open(fp_j, 'r', encoding='utf-8') as f:
                        d_j = json.load(f)
                except Exception:
                    continue
                    
                if is_song_content_identical(d_i, d_j):
                    # 더 오래된 파일인 files[i]를 삭제 대상에 추가
                    to_delete.add(files[i])
                    break
                    
        for f in to_delete:
            try:
                os.remove(os.path.join(song_dir, f))
                total_pruned += 1
            except Exception:
                pass
                
        # 최대 3버전 유지 정리
        prune_excess_song_versions(song_dir, max_versions=3)
        
    return total_pruned

def save_song_to_library(song_data, date_str="미입력", department="청년부"):
    """
    단일 찬양곡 데이터를 표준 파일명([프리셋][찬양곡][제목][날짜].json)으로 라이브러리에 저장합니다.
    1. 동일 곡명의 기존 버전들로 탐색 범위를 한정
    2. 파트별 가사 슬라이드 및 루틴이 완전히 동일한지 비교
    3. 동일한 버전이 이미 존재하면 새 버전 파일로 추가하지 않고 날짜만 갱신(기존 파일 교체)
    4. 내용이 다르면 신규 버전으로 추가 (최신 3개 버전 유지 정책 준수)
    """
    ensure_library_dir()
    title = song_data.get("title", "").strip()
    if not title:
        return None
        
    safe_title = sanitize_filename(title)
    safe_date = sanitize_filename(date_str) if date_str != "미입력" else "기본버전"
    
    song_dir = os.path.join(SONG_LIBRARY_DIR, safe_title)
    os.makedirs(song_dir, exist_ok=True)
    
    clean_dept = normalize_department_name(department)
    new_filename = format_standard_filename(clean_dept, "찬양곡", safe_title, safe_date, "json")
    new_file_path = os.path.join(song_dir, new_filename)
    
    payload = {
        "title": title,
        "date": date_str,
        "department": clean_dept,
        "routine_raw": song_data.get("routine_raw", ""),
        "parts": song_data.get("parts", {})
    }
    
    # 1. 같은 이름의 찬양곡 폴더 내 기존 버전 파일들 탐색
    existing_files = [f for f in os.listdir(song_dir) if f.endswith('.json')]
    
    matched_existing_path = None
    old_file_to_remove = None
    
    for fname in existing_files:
        fpath = os.path.join(song_dir, fname)
        try:
            with open(fpath, 'r', encoding='utf-8') as f:
                existing_data = json.load(f)
            if is_song_content_identical(payload, existing_data):
                matched_existing_path = fpath
                if os.path.normpath(fpath) != os.path.normpath(new_file_path):
                    old_file_to_remove = fpath
                break
        except Exception:
            continue
            
    # 2. 동일한 내용의 기존 버전 발견: 날짜 갱신 (구 버전 파일 정리 후 최신 일자로 저장)
    if matched_existing_path:
        if old_file_to_remove and os.path.exists(old_file_to_remove):
            try:
                os.remove(old_file_to_remove)
            except Exception:
                pass
        with open(new_file_path, 'w', encoding='utf-8') as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    else:
        # 3. 다른 내용: 신규 버전으로 추가 저장
        with open(new_file_path, 'w', encoding='utf-8') as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        prune_excess_song_versions(song_dir, max_versions=3)
        
    return new_file_path

def decompose_and_save_songs_from_conti(conti_input, department="청년부"):
    """
    콘티 텍스트 또는 파싱된 콘티 딕셔너리에서 각 곡을 개별 분해하여
    3_찬양곡_라이브러리에 표준 파일명으로 각각 저장하고,
    저장된 파일 전체 경로 목록을 반환합니다.
    """
    if isinstance(conti_input, str):
        from core.text_engine import parse_user_conti
        parsed_conti = parse_user_conti(conti_input)
    elif isinstance(conti_input, dict):
        parsed_conti = conti_input
    else:
        return []

    date_str = parsed_conti.get("date_str", "미입력")
    songs = parsed_conti.get("songs", [])
    saved_files = []
    
    clean_dept = normalize_department_name(department)
    for s in songs:
        if s.get("title"):
            saved_path = save_song_to_library(s, date_str=date_str, department=clean_dept)
            if saved_path:
                saved_files.append(saved_path)
                
    return saved_files

def save_all_songs_from_conti(conti_input, department="청년부"):
    """
    하위 호환성 유지: 콘티 내 모든 곡을 저장하고 저장된 곡 수를 반환합니다.
    """
    files = decompose_and_save_songs_from_conti(conti_input, department=department)
    return len(files)

def format_song_to_conti_text(song_data, index=None):
    """
    라이브러리의 곡 데이터를 콘티 서식 텍스트로 변환합니다.
    """
    title = song_data.get("title", "")
    clean_title = re.sub(r'^\d+[\.\s]+', '', title).strip()
    idx_str = f"{index}. " if index is not None else ""
    routine = song_data.get("routine_raw", "")
    parts = song_data.get("parts", {})
    
    lines = [f"## {idx_str}{clean_title}"]
    if routine:
        lines.append(f"루틴: {routine}")
    lines.append("")
    
    for part_name, slides in parts.items():
        lines.append(f"[{part_name}]")
        for s_text in slides:
            lines.append(s_text.strip())
            lines.append("") # 슬라이드 간 빈 줄 (더블 엔터)
            
    return '\n'.join(lines).strip()

def replace_song_in_conti(conti_text, target_index, song_data):
    """
    콘티 본문 내에서 특정 순번(target_index, 1부터 시작)의 찬양곡 블록(## ...)을 
    라이브러리 곡 데이터로 1:1 교체합니다.
    """
    if not conti_text or not conti_text.strip():
        return format_song_to_conti_text(song_data, index=target_index)

    parts = re.split(r'(?m)^(?=##\s*)', conti_text)
    if not parts:
        return format_song_to_conti_text(song_data, index=target_index)

    preamble = parts[0]
    song_blocks = parts[1:]
    new_song_text = format_song_to_conti_text(song_data, index=target_index)

    if 1 <= target_index <= len(song_blocks):
        song_blocks[target_index - 1] = new_song_text + "\n\n"
    else:
        song_blocks.append(new_song_text + "\n\n")

    return (preamble.rstrip() + "\n\n" + "".join(song_blocks)).strip() + "\n"

def append_song_to_conti(conti_text, song_data):
    """
    콘티 본문의 마지막에 새로운 찬양곡을 다음 순번 번호와 함께 덧붙입니다.
    """
    if not conti_text or not conti_text.strip():
        return format_song_to_conti_text(song_data, index=1)

    parts = re.split(r'(?m)^(?=##\s*)', conti_text)
    song_count = len(parts) - 1 if len(parts) > 1 else 0
    next_index = song_count + 1
    new_song_text = format_song_to_conti_text(song_data, index=next_index)

    return conti_text.strip() + "\n\n" + new_song_text
