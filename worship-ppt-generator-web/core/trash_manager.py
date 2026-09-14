"""
core/trash_manager.py
관리자 모드 전용 휴지통(Trash Bin) 관리 및 일반 모드 격리 엔진
- 소프트 삭제(Soft Delete): 일반 모드(자료실 및 사이드바 검색)에서 즉시 숨기고 카탈로그 휴지통으로 이동
- 복구(Restore): 휴지통에서 다시 일반 자료실로 즉시 복원
- 하드 영구 삭제(Hard Delete): 텔레그램 Bot API deleteMessage로 클라우드 원본 메시지/파일을 실제 파기하고 로컬에서도 완전 제거
- 텔레그램 .cloud_catalog.json에 동기화되어 서버 재부팅 후에도 100% 영구 유지
"""

import os
import shutil
from datetime import datetime

from core.cloud_config import (
    DIR_PPTX,
    DIR_SONGS,
    CHAT_ID,
    parse_standard_filename,
    sanitize_filename_part,
    ensure_data_directories
)
from core.cloud_sync import (
    load_local_catalog,
    save_local_catalog,
    publish_catalog,
    telegram_api_call
)


def get_trash_catalog():
    """현재 카탈로그에 등록된 휴지통 항목 목록을 반환합니다."""
    cat = load_local_catalog()
    trash = cat.get("trash", {})
    if not isinstance(trash, dict):
        trash = {"pptx": [], "songs": []}
    trash.setdefault("pptx", [])
    trash.setdefault("songs", [])
    return trash


def is_pptx_trashed(filename):
    """지정된 PPT 파일이 휴지통에 들어있는지 확인합니다."""
    if not filename:
        return False
    trash = get_trash_catalog()
    trashed_files = {it.get("filename") for it in trash.get("pptx", []) if it.get("filename")}
    return filename in trashed_files


def is_song_trashed(song_title, date_str):
    """지정된 찬양곡 특정 버전이 휴지통에 들어있는지 확인합니다."""
    if not song_title or not date_str:
        return False
    clean_title = sanitize_filename_part(song_title)
    clean_date = str(date_str).replace(".", "").replace("-", "")
    trash = get_trash_catalog()
    for it in trash.get("songs", []):
        it_title = sanitize_filename_part(it.get("title", ""))
        it_date = str(it.get("date", "")).replace(".", "").replace("-", "")
        if it_title == clean_title and it_date == clean_date:
            return True
    return False


def move_pptx_to_trash(filename):
    """
    찬양 PPT 파일을 소프트 삭제(휴지통 이동)합니다.
    일반 모드 자료실에서 즉시 감춰집니다.
    """
    cat = load_local_catalog()
    trash = cat.setdefault("trash", {})
    trash_pptx = trash.setdefault("pptx", [])

    # 이미 휴지통에 있다면 무시
    if any(it.get("filename") == filename for it in trash_pptx):
        return True

    # 1. 카탈로그 items에서 원본 메타데이터 탐색
    item_meta = None
    for it in cat.get("items", []):
        if it.get("filename") == filename:
            item_meta = it
            break

    # 2. 로컬 디스크 파일 정보 탐색 (카탈로그에 없었던 경우)
    local_path = os.path.join(DIR_PPTX, filename)
    file_size = os.path.getsize(local_path) if os.path.exists(local_path) else 0

    meta = parse_standard_filename(filename) if filename else {}
    title = item_meta.get("title") if item_meta else meta.get("title", filename)
    date = item_meta.get("date") if item_meta else meta.get("date", "")
    preset = item_meta.get("preset") if item_meta else meta.get("preset", "기타")
    file_id = item_meta.get("file_id") if item_meta else None
    message_id = item_meta.get("message_id") if item_meta else None

    entry = {
        "filename": filename,
        "title": title,
        "date": date,
        "preset": preset,
        "file_size": file_size,
        "file_id": file_id,
        "message_id": message_id,
        "deleted_at": datetime.now().strftime("%Y-%m-%d %H:%M")
    }
    trash_pptx.append(entry)
    save_local_catalog(cat)
    publish_catalog(cat)
    return True


def restore_pptx_from_trash(filename):
    """휴지통에 있던 찬양 PPT를 복구하여 일반 모드 자료실에 다시 노출합니다."""
    cat = load_local_catalog()
    trash = cat.setdefault("trash", {})
    trash_pptx = trash.setdefault("pptx", [])

    cat["trash"]["pptx"] = [it for it in trash_pptx if it.get("filename") != filename]
    save_local_catalog(cat)
    publish_catalog(cat)
    return True


def hard_delete_pptx(filename):
    """
    찬양 PPT를 영구 삭제(하드 딜리트)합니다.
    - 텔레그램 채널에서 메시지를 실제 deleteMessage로 파기
    - 로컬 보관함 파일 영구 삭제
    - 카탈로그에서 완전 소멸
    """
    cat = load_local_catalog()
    trash = cat.setdefault("trash", {})
    trash_pptx = trash.setdefault("pptx", [])

    # 대상 항목 탐색
    target_entry = None
    for it in trash_pptx:
        if it.get("filename") == filename:
            target_entry = it
            break

    if not target_entry:
        for it in cat.get("items", []):
            if it.get("filename") == filename:
                target_entry = it
                break

    # 1. 텔레그램 채널 메시지 영구 파기
    if target_entry and target_entry.get("message_id"):
        mid = target_entry["message_id"]
        try:
            telegram_api_call("deleteMessage", {"chat_id": CHAT_ID, "message_id": mid})
        except Exception as e:
            print(f"[TrashManager] 텔레그램 메시지 삭제 실패 ({mid}): {e}")

    # 2. 로컬 디스크 파일 삭제
    local_path = os.path.join(DIR_PPTX, filename)
    if os.path.exists(local_path):
        try:
            os.remove(local_path)
        except Exception:
            pass

    # 3. 카탈로그에서 완전 제거
    cat["trash"]["pptx"] = [it for it in trash_pptx if it.get("filename") != filename]
    cat["items"] = [it for it in cat.get("items", []) if it.get("filename") != filename]
    save_local_catalog(cat)
    publish_catalog(cat)
    return True


def move_song_to_trash(song_title, date_str, routine="", department="청년부"):
    """
    찬양곡 특정 버전을 소프트 삭제(휴지통 이동)합니다.
    일반 모드 자료실 및 사이드바 검색에서 즉시 감춰집니다.
    """
    cat = load_local_catalog()
    trash = cat.setdefault("trash", {})
    trash_songs = trash.setdefault("songs", [])

    clean_title = sanitize_filename_part(song_title)
    clean_date = str(date_str).replace(".", "").replace("-", "")

    # 이미 휴지통에 있다면 무시
    for it in trash_songs:
        it_title = sanitize_filename_part(it.get("title", ""))
        it_date = str(it.get("date", "")).replace(".", "").replace("-", "")
        if it_title == clean_title and it_date == clean_date:
            return True

    # 카탈로그 items에서 원본 탐색
    item_meta = None
    for it in cat.get("items", []):
        if it.get("type") == "찬양곡" and sanitize_filename_part(it.get("title", "")) == clean_title:
            if str(it.get("date", "")).replace(".", "").replace("-", "") == clean_date:
                item_meta = it
                break

    filename = item_meta.get("filename") if item_meta else None
    file_id = item_meta.get("file_id") if item_meta else None
    message_id = item_meta.get("message_id") if item_meta else None

    # 로컬 파일에서 루틴/부서 정보 보강 시도
    song_folder = os.path.join(DIR_SONGS, clean_title)
    if os.path.exists(song_folder):
        for f in os.listdir(song_folder):
            if f.endswith(".json") and clean_date in f.replace(".", "").replace("-", ""):
                filename = f
                break

    entry = {
        "title": clean_title,
        "date": clean_date,
        "filename": filename or f"[{department}][찬양곡][{clean_title}][{clean_date}].json",
        "routine": routine,
        "department": department,
        "file_id": file_id,
        "message_id": message_id,
        "deleted_at": datetime.now().strftime("%Y-%m-%d %H:%M")
    }
    trash_songs.append(entry)
    save_local_catalog(cat)
    publish_catalog(cat)
    return True


def restore_song_from_trash(song_title, date_str):
    """휴지통에 있던 찬양곡 버전을 복구하여 일반 모드 자료실/검색에 다시 노출합니다."""
    cat = load_local_catalog()
    trash = cat.setdefault("trash", {})
    trash_songs = trash.setdefault("songs", [])

    clean_title = sanitize_filename_part(song_title)
    clean_date = str(date_str).replace(".", "").replace("-", "")

    new_list = []
    for it in trash_songs:
        it_title = sanitize_filename_part(it.get("title", ""))
        it_date = str(it.get("date", "")).replace(".", "").replace("-", "")
        if it_title == clean_title and it_date == clean_date:
            continue
        new_list.append(it)

    cat["trash"]["songs"] = new_list
    save_local_catalog(cat)
    publish_catalog(cat)
    return True


def hard_delete_song(song_title, date_str):
    """
    찬양곡 버전을 영구 삭제(하드 딜리트)합니다.
    - 텔레그램 채널에서 메시지를 실제 deleteMessage로 파기
    - 로컬 디스크 JSON 파일 영구 삭제 (폴더가 비면 폴더도 삭제)
    - 카탈로그에서 완전 소멸
    """
    cat = load_local_catalog()
    trash = cat.setdefault("trash", {})
    trash_songs = trash.setdefault("songs", [])

    clean_title = sanitize_filename_part(song_title)
    clean_date = str(date_str).replace(".", "").replace("-", "")

    target_entry = None
    for it in trash_songs:
        it_title = sanitize_filename_part(it.get("title", ""))
        it_date = str(it.get("date", "")).replace(".", "").replace("-", "")
        if it_title == clean_title and it_date == clean_date:
            target_entry = it
            break

    # 1. 텔레그램 메시지 파기
    if target_entry and target_entry.get("message_id"):
        mid = target_entry["message_id"]
        try:
            telegram_api_call("deleteMessage", {"chat_id": CHAT_ID, "message_id": mid})
        except Exception as e:
            print(f"[TrashManager] 텔레그램 곡 메시지 삭제 실패 ({mid}): {e}")

    # 2. 로컬 JSON 파일 삭제
    song_folder = os.path.join(DIR_SONGS, clean_title)
    if os.path.exists(song_folder):
        for f in os.listdir(song_folder):
            if f.endswith(".json") and clean_date in f.replace(".", "").replace("-", ""):
                try:
                    os.remove(os.path.join(song_folder, f))
                except Exception:
                    pass
        # 폴더 내에 남은 .json이 없으면 폴더 자체 삭제
        remaining = [f for f in os.listdir(song_folder) if f.endswith(".json")]
        if not remaining:
            try:
                shutil.rmtree(song_folder)
            except Exception:
                pass

    # 3. 카탈로그에서 완전 제거
    new_trash_songs = []
    for it in trash_songs:
        it_title = sanitize_filename_part(it.get("title", ""))
        it_date = str(it.get("date", "")).replace(".", "").replace("-", "")
        if it_title == clean_title and it_date == clean_date:
            continue
        new_trash_songs.append(it)
    cat["trash"]["songs"] = new_trash_songs

    new_items = []
    for it in cat.get("items", []):
        if it.get("type") == "찬양곡" and sanitize_filename_part(it.get("title", "")) == clean_title:
            if str(it.get("date", "")).replace(".", "").replace("-", "") == clean_date:
                continue
        new_items.append(it)
    cat["items"] = new_items

    save_local_catalog(cat)
    publish_catalog(cat)
    return True


def empty_all_trash():
    """휴지통 내의 모든 PPT 및 찬양곡을 일괄 영구 파기합니다."""
    trash = get_trash_catalog()
    pptx_list = list(trash.get("pptx", []))
    song_list = list(trash.get("songs", []))

    deleted_p_count = 0
    deleted_s_count = 0

    for it in pptx_list:
        fname = it.get("filename")
        if fname:
            hard_delete_pptx(fname)
            deleted_p_count += 1

    for it in song_list:
        stitle = it.get("title")
        sdate = it.get("date")
        if stitle and sdate:
            hard_delete_song(stitle, sdate)
            deleted_s_count += 1

    return deleted_p_count, deleted_s_count
