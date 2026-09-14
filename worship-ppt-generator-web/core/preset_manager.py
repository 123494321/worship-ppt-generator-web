"""
core/preset_manager.py
관리자 모드 전용 표준 양식(부서별 기준 PPT 및 표지 이미지) 관리 엔진
- 신규 표준 양식(부서) 등록 및 텔레그램 클라우드 자동 업로드
- 기존 표준 양식 수정/교체 (이름 변경, 기준 PPT 교체, 표지 이미지 교체/삭제)
- 표준 양식 안전 삭제 및 카탈로그 동기화
"""

import os
import re
import shutil
from datetime import datetime

from core.cloud_config import (
    DIR_PRESETS,
    CHAT_ID,
    sanitize_filename_part,
    ensure_data_directories
)
from core.cloud_sync import (
    load_local_catalog,
    save_local_catalog,
    publish_catalog,
    upload_multipart_document,
    telegram_api_call
)
from core.pptx_builder import get_preset_info_dict


def add_preset_department(dept_name, pptx_bytes, pptx_filename="기준 PPT.pptx", cover_bytes=None, cover_filename=None):
    """
    신규 표준 양식(부서)을 생성하고 로컬 디스크 및 텔레그램 클라우드에 등록합니다.
    """
    clean_dept = sanitize_filename_part(dept_name)
    if not clean_dept:
        return False, "부서(표준 양식) 이름을 올바르게 입력해 주세요."

    ensure_data_directories()
    dept_dir = os.path.join(DIR_PRESETS, clean_dept)
    if os.path.exists(dept_dir):
        return False, f"이미 '{clean_dept}' 표준 양식이 존재합니다. 다른 이름을 사용해 주세요."

    os.makedirs(dept_dir, exist_ok=True)

    # 1. 기준 PPTX 파일 로컬 저장
    target_pptx_name = f"{clean_dept} 기준 PPT.pptx"
    target_pptx_path = os.path.join(dept_dir, target_pptx_name)
    try:
        with open(target_pptx_path, "wb") as f:
            f.write(pptx_bytes)
    except Exception as e:
        return False, f"기준 PPT 파일 로컬 저장 실패: {e}"

    # 2. 표지 이미지 로컬 저장 (선택)
    target_cover_name = None
    if cover_bytes and cover_filename:
        _, ext = os.path.splitext(cover_filename)
        if not ext or ext.lower() not in [".png", ".jpg", ".jpeg"]:
            ext = ".jpg"
        target_cover_name = f"{clean_dept} 표지_{datetime.now().strftime('%y%m%d_%H%M%S')}{ext}"
        target_cover_path = os.path.join(dept_dir, target_cover_name)
        try:
            with open(target_cover_path, "wb") as f:
                f.write(cover_bytes)
        except Exception:
            target_cover_name = None

    # 3. 텔레그램 클라우드 업로드
    pptx_fid = None
    pptx_mid = None
    up_res = upload_multipart_document(
        chat_id=CHAT_ID,
        file_bytes=pptx_bytes,
        filename=target_pptx_name,
        caption=f"#기준 [{clean_dept}] 공식 기준 PPTX 자동 등록"
    )
    if up_res.get("ok"):
        pptx_mid = up_res["result"].get("message_id")
        pptx_fid = up_res["result"].get("document", {}).get("file_id")

    cover_fid = None
    cover_mid = None
    if cover_bytes and target_cover_name:
        cup_res = upload_multipart_document(
            chat_id=CHAT_ID,
            file_bytes=cover_bytes,
            filename=target_cover_name,
            caption=f"#표지 [{clean_dept}] 공식 표지 이미지 자동 등록"
        )
        if cup_res.get("ok"):
            cover_mid = cup_res["result"].get("message_id")
            cover_fid = cup_res["result"].get("document", {}).get("file_id") or cup_res["result"].get("photo", [{}])[-1].get("file_id")

    # 4. 카탈로그 등록 및 클라우드 게시
    cat = load_local_catalog()
    cat.setdefault("presets", {})[clean_dept] = {
        "template_filename": target_pptx_name,
        "template_file_id": pptx_fid,
        "template_msg_id": pptx_mid,
        "template_updated_at": datetime.now().isoformat(),
        "cover_filename": target_cover_name,
        "cover_file_id": cover_fid,
        "cover_msg_id": cover_mid,
        "cover_updated_at": datetime.now().isoformat() if target_cover_name else None
    }
    save_local_catalog(cat)
    publish_catalog(cat)

    return True, f"'{clean_dept}' 표준 양식이 성공적으로 생성되고 클라우드에 백업되었습니다!"


def update_preset_department(old_dept_name, new_dept_name=None, pptx_bytes=None, pptx_filename=None, cover_bytes=None, cover_filename=None, remove_cover=False):
    """
    기존 표준 양식을 수정/교체합니다.
    """
    clean_old = sanitize_filename_part(old_dept_name)
    target_dept = clean_old
    old_dir = os.path.join(DIR_PRESETS, clean_old)

    if not os.path.exists(old_dir):
        return False, f"'{clean_old}' 표준 양식을 찾을 수 없습니다."

    cat = load_local_catalog()
    presets_cat = cat.setdefault("presets", {})
    preset_meta = presets_cat.get(clean_old, {})

    # 1. 부서명 변경 처리
    if new_dept_name:
        clean_new = sanitize_filename_part(new_dept_name)
        if clean_new and clean_new != clean_old:
            new_dir = os.path.join(DIR_PRESETS, clean_new)
            if os.path.exists(new_dir):
                return False, f"이미 '{clean_new}' 이름의 표준 양식이 존재합니다."
            try:
                os.rename(old_dir, new_dir)
                target_dept = clean_new
                dept_dir = new_dir
                # 카탈로그 키 갱신
                presets_cat[clean_new] = preset_meta
                del presets_cat[clean_old]
                preset_meta = presets_cat[clean_new]
            except Exception as e:
                return False, f"부서 이름 변경 실패: {e}"
        else:
            dept_dir = old_dir
    else:
        dept_dir = old_dir

    # 2. 기준 PPTX 파일 교체
    if pptx_bytes:
        target_pptx_name = f"{target_dept} 기준 PPT.pptx"
        target_pptx_path = os.path.join(dept_dir, target_pptx_name)
        try:
            with open(target_pptx_path, "wb") as f:
                f.write(pptx_bytes)
        except Exception as e:
            return False, f"기준 PPT 교체 저장 실패: {e}"

        up_res = upload_multipart_document(
            chat_id=CHAT_ID,
            file_bytes=pptx_bytes,
            filename=target_pptx_name,
            caption=f"#기준 [{target_dept}] 기준 PPTX 교체"
        )
        if up_res.get("ok"):
            preset_meta["template_file_id"] = up_res["result"].get("document", {}).get("file_id")
            preset_meta["template_msg_id"] = up_res["result"].get("message_id")
            preset_meta["template_filename"] = target_pptx_name
            preset_meta["template_updated_at"] = datetime.now().isoformat()

    # 3. 표지 이미지 삭제 또는 교체
    if remove_cover:
        # 로컬 표지 이미지 파일 제거
        for f in os.listdir(dept_dir):
            if f.lower().endswith(('.png', '.jpg', '.jpeg')):
                try:
                    os.remove(os.path.join(dept_dir, f))
                except Exception:
                    pass
        preset_meta["cover_file_id"] = None
        preset_meta["cover_filename"] = None
        preset_meta["cover_msg_id"] = None
        preset_meta["cover_updated_at"] = None

    elif cover_bytes and cover_filename:
        # 기존 표지 정리 후 새 표지 저장
        for f in os.listdir(dept_dir):
            if f.lower().endswith(('.png', '.jpg', '.jpeg')):
                try:
                    os.remove(os.path.join(dept_dir, f))
                except Exception:
                    pass

        _, ext = os.path.splitext(cover_filename)
        if not ext or ext.lower() not in [".png", ".jpg", ".jpeg"]:
            ext = ".jpg"
        target_cover_name = f"{target_dept} 표지_{datetime.now().strftime('%y%m%d_%H%M%S')}{ext}"
        target_cover_path = os.path.join(dept_dir, target_cover_name)
        try:
            with open(target_cover_path, "wb") as f:
                f.write(cover_bytes)
        except Exception:
            target_cover_name = None

        if target_cover_name:
            cup_res = upload_multipart_document(
                chat_id=CHAT_ID,
                file_bytes=cover_bytes,
                filename=target_cover_name,
                caption=f"#표지 [{target_dept}] 표지 이미지 교체"
            )
            if cup_res.get("ok"):
                preset_meta["cover_file_id"] = cup_res["result"].get("document", {}).get("file_id") or cup_res["result"].get("photo", [{}])[-1].get("file_id")
                preset_meta["cover_msg_id"] = cup_res["result"].get("message_id")
                preset_meta["cover_filename"] = target_cover_name
                preset_meta["cover_updated_at"] = datetime.now().isoformat()

    save_local_catalog(cat)
    publish_catalog(cat)
    return True, f"'{target_dept}' 표준 양식 설정이 성공적으로 변경되었습니다!"


def delete_preset_department(dept_name):
    """
    표준 양식을 안전하게 삭제합니다 (기본 청년부 제외).
    """
    clean_dept = sanitize_filename_part(dept_name)
    if clean_dept == "청년부":
        return False, "기본 공식 양식인 '청년부'는 삭제할 수 없습니다."

    dept_dir = os.path.join(DIR_PRESETS, clean_dept)
    if os.path.exists(dept_dir):
        try:
            shutil.rmtree(dept_dir)
        except Exception as e:
            return False, f"로컬 표준 양식 폴더 삭제 실패: {e}"

    cat = load_local_catalog()
    presets = cat.get("presets", {})
    if clean_dept in presets:
        del presets[clean_dept]
        save_local_catalog(cat)
        publish_catalog(cat)

    return True, f"'{clean_dept}' 표준 양식이 성공적으로 삭제되었습니다."
