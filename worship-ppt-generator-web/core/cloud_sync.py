"""
core/cloud_sync.py
텔레그램 Bot API 기반 클라우드 동기화 및 OTA 프리셋 관리 엔진

핵심 원칙:
1. 오프라인 우선: 네트워크 오류가 발생해도 로컬 작업은 100% 정상 작동하며 에러를 던지지 않고 경고만 기록.
2. 핀된 카탈로그(Pinned Catalog): Bot API의 히스토리 조회 한계를 극복하기 위해,
   채널 상단 고정 메시지에 cloud_catalog.json을 첨부하여 모든 클라이언트가 언제든 전체 클라우드 자산을 즉시 복원.
3. OTA 프리셋/표지 원격 패치: 채널에 #프리셋 [부서] 또는 #표지 [부서] 태그가 달린 파일이 올라오면 로컬 기준 PPT 자동 최신화.
4. 로컬 캐시 롤링: 90일 지난 로컬 PPTX는 자동 정리하되 찬양곡 라이브러리와 콘티는 영구 보존.
"""

import os
import re
import io
import json
import time
import shutil
import urllib.request
import urllib.parse
from datetime import datetime, timedelta

from core.cloud_config import (
    BOT_TOKEN,
    CHAT_ID,
    DEFAULT_DEPARTMENT,
    DATA_ROOT,
    DIR_PRESETS,
    DIR_PPTX,
    DIR_SONGS,
    DIR_CONTI,
    ensure_data_directories,
    format_standard_filename,
    parse_standard_filename,
    sanitize_filename_part,
    get_official_departments,
    is_official_department,
    normalize_department_name,
)

# 로컬 캐시 및 상태 파일 경로
LOCAL_CATALOG_FILE = os.path.join(DATA_ROOT, ".cloud_catalog.json")
LAST_UPDATE_ID_FILE = os.path.join(DATA_ROOT, ".last_update_id.txt")
LAST_CATALOG_META_FILE = os.path.join(DATA_ROOT, ".last_catalog_meta.json")

# 채널 상단 고정 공식 안내문 텍스트 (관리자 갱신 시에만 단독 발송 및 고정)
GUIDE_CAPTION_TEXT = """📌 [로고스 찬양 PPT 시스템 공식 클라우드 저장소]
본 채널은 찬양 PPT 및 찬양곡 라이브러리가 자동 백업되는 공식 저장소입니다.

💡 [공식 프리셋 관리 및 서식 등록 방법]
1. 공식 프리셋 신규 생성:
- 일반 텍스트로 #추가 : [프리셋 이름] 전송
  예: #추가 : 청년부
  예: #추가 : 중고등부

2. 기준 PPT 서식 등록/교체:
- .pptx 파일 첨부 + 캡션 #기준 (또는 #기준 [프리셋 이름]) 전송
- 봇이 파일을 인식하면 적용할 프리셋을 선택합니다.

3. 표지 이미지 등록/교체:
- 이미지 파일 첨부 + 캡션 #표지 (또는 #표지 [프리셋 이름]) 전송
- 봇이 이미지를 인식하면 적용할 프리셋을 선택합니다."""


# ==============================================================================
# 1. 텔레그램 저수준 HTTP 통신 헬퍼
# ==============================================================================

def telegram_api_call(method, params=None, timeout=30):
    """
    텔레그램 Bot API POST JSON 호출 (복합 데이터 및 reply_markup 안전 지원)
    """
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/{method}"
    try:
        if params:
            payload = json.dumps(params).encode("utf-8")
            req = urllib.request.Request(
                url, 
                data=payload, 
                headers={"Content-Type": "application/json", "User-Agent": "LogosWorshipPPT/1.5.0"}
            )
        else:
            req = urllib.request.Request(url, headers={"User-Agent": "LogosWorshipPPT/1.5.0"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            return data
    except Exception as e:
        return {"ok": False, "error": str(e)}


def upload_multipart_document(chat_id, file_bytes, filename, caption=None, timeout=120):
    """
    sendDocument 엔드포인트로 파일 바이트 multipart/form-data 업로드
    """
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendDocument"
    boundary = "----LogosWorshipBoundary" + str(int(time.time() * 1000))
    body = io.BytesIO()

    def write_field(name, val):
        body.write(f"--{boundary}\r\n".encode("utf-8"))
        body.write(f'Content-Disposition: form-data; name="{name}"\r\n\r\n'.encode("utf-8"))
        body.write(str(val).encode("utf-8"))
        body.write(b"\r\n")

    write_field("chat_id", chat_id)
    if caption:
        write_field("caption", caption)

    # 파일 본문
    body.write(f"--{boundary}\r\n".encode("utf-8"))
    body.write(f'Content-Disposition: form-data; name="document"; filename="{filename}"\r\n'.encode("utf-8"))
    body.write(b"Content-Type: application/octet-stream\r\n\r\n")
    body.write(file_bytes)
    body.write(b"\r\n")

    body.write(f"--{boundary}--\r\n".encode("utf-8"))
    data_payload = body.getvalue()

    headers = {
        "Content-Type": f"multipart/form-data; boundary={boundary}",
        "Content-Length": str(len(data_payload)),
        "User-Agent": "LogosWorshipPPT/1.5.0"
    }

    req = urllib.request.Request(url, data=data_payload, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            res = json.loads(resp.read().decode("utf-8"))
            return res
    except Exception as e:
        return {"ok": False, "error": str(e)}


def upload_file_to_telegram(file_source, filename, caption=None, timeout=120):
    """
    파일 경로(str) 또는 바이트(bytes/BytesIO)를 받아 텔레그램 채널로 업로드합니다.
    성공 시: (True, message_dict, None)
    실패 시: (False, None, error_message)
    """
    if isinstance(file_source, str):
        if not os.path.exists(file_source):
            return False, None, f"파일을 찾을 수 없습니다: {file_source}"
        with open(file_source, "rb") as f:
            file_bytes = f.read()
    elif isinstance(file_source, bytes):
        file_bytes = file_source
    elif hasattr(file_source, "read"):
        file_bytes = file_source.read()
    else:
        return False, None, "유효하지 않은 파일 데이터 형식입니다."

    res = upload_multipart_document(CHAT_ID, file_bytes, filename, caption=caption, timeout=timeout)
    if res.get("ok"):
        return True, res.get("result", {}), None
    else:
        err = res.get("description") or res.get("error") or "알 수 없는 오류"
        return False, None, err


def download_file_by_id(file_id, dest_path, timeout=60):
    """
    Telegram file_id를 조회하여 지정된 로컬 경로로 안전하게 다운로드합니다.
    성공 시: (True, None)
    실패 시: (False, error_message)
    """
    file_info = telegram_api_call("getFile", {"file_id": file_id}, timeout=timeout)
    if not file_info.get("ok"):
        return False, file_info.get("description", "파일 정보를 가져올 수 없습니다.")

    file_path = file_info["result"].get("file_path")
    if not file_path:
        return False, "file_path가 응답에 없습니다."

    download_url = f"https://api.telegram.org/file/bot{BOT_TOKEN}/{file_path}"
    temp_path = dest_path + f".tmp_{int(time.time())}"
    try:
        os.makedirs(os.path.dirname(dest_path), exist_ok=True)
        req = urllib.request.Request(download_url, headers={"User-Agent": "LogosWorshipPPT/1.5.0"})
        with urllib.request.urlopen(req, timeout=timeout) as resp, open(temp_path, "wb") as out_f:
            shutil.copyfileobj(resp, out_f)
        if os.path.exists(dest_path):
            os.remove(dest_path)
        os.rename(temp_path, dest_path)
        return True, None
    except Exception as e:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass
        return False, str(e)


# ==============================================================================
# 2. 클라우드 카탈로그(Manifest) 관리
# ==============================================================================

def get_empty_catalog():
    return {
        "version": "1.0",
        "updated_at": datetime.now().isoformat(),
        "channel_id": CHAT_ID,
        "items": [],
        "presets": {}
    }


def load_local_catalog():
    """로컬에 저장된 클라우드 카탈로그 JSON을 읽습니다."""
    ensure_data_directories()
    if os.path.exists(LOCAL_CATALOG_FILE):
        try:
            with open(LOCAL_CATALOG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return get_empty_catalog()


def save_local_catalog(catalog_data):
    """로컬 클라우드 카탈로그 JSON을 저장합니다."""
    ensure_data_directories()
    catalog_data["updated_at"] = datetime.now().isoformat()
    try:
        with open(LOCAL_CATALOG_FILE, "w", encoding="utf-8") as f:
            json.dump(catalog_data, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"[CloudSync] 로컬 카탈로그 저장 실패: {e}")
        return False


def set_channel_catalog_ref(file_id):
    """
    채널 설명(description)에 최신 cloud_catalog.json의 file_id를 등록합니다.
    이를 통해 안내 메시지를 오염시키거나 고정 메시지를 중복 생성하지 않고도
    모든 클라이언트가 최신 카탈로그를 안전하게 조회할 수 있습니다.
    """
    try:
        chat_res = telegram_api_call("getChat", {"chat_id": CHAT_ID})
        curr_desc = chat_res.get("result", {}).get("description", "")
        # 기존 catalog: 태그 제거 후 최신 file_id로 갱신
        clean_desc = re.sub(r'\s*\|\s*catalog:\s*[^\s\]]+', '', curr_desc).strip()
        if not clean_desc:
            clean_desc = "찬양 PPT 제작 프로그램 온라인 저장소"
        new_desc = f"{clean_desc} | catalog: {file_id}"[:255]
        return telegram_api_call("setChatDescription", {"chat_id": CHAT_ID, "description": new_desc})
    except Exception as e:
        print(f"[CloudSync] 채널 설명 갱신 실패: {e}")
        return {"ok": False, "error": str(e)}


def fetch_remote_catalog():
    """
    1. 채널 설명(description)에서 catalog: {file_id}를 우선 추출하여 다운로드합니다.
    2. 없을 경우 채널 고정 메시지의 첨부파일을 확인하는 레거시 폴백을 수행합니다.
    성공 시: catalog_dict
    실패 시: None
    """
    chat_res = telegram_api_call("getChat", {"chat_id": CHAT_ID})
    if not chat_res.get("ok"):
        return None

    desc = chat_res.get("result", {}).get("description", "")
    m = re.search(r'catalog:\s*([^\s\]]+)', desc)
    file_id = None
    if m:
        file_id = m.group(1).strip()

    if not file_id:
        # 레거시 폴백: 고정 메시지 첨부파일 확인
        pinned = chat_res.get("result", {}).get("pinned_message", {})
        doc = pinned.get("document")
        if doc and doc.get("file_name") == "cloud_catalog.json":
            file_id = doc.get("file_id")

    if not file_id:
        return None

    file_info = telegram_api_call("getFile", {"file_id": file_id})
    if not file_info.get("ok"):
        return None

    file_path = file_info["result"].get("file_path")
    if not file_path:
        return None

    url = f"https://api.telegram.org/file/bot{BOT_TOKEN}/{file_path}"
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "LogosWorshipPPT/1.5.0"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            content = resp.read().decode("utf-8")
            return json.loads(content)
    except Exception as e:
        print(f"[CloudSync] 원격 카탈로그 다운로드 파싱 오류: {e}")
        return None


def merge_catalogs(local_cat, remote_cat):
    """
    로컬 카탈로그와 원격 카탈로그를 안전하게 병합합니다.
    filename 및 file_id 기준으로 중복을 방지하며 최신 정보를 유지합니다.
    """
    if not remote_cat or "items" not in remote_cat:
        return local_cat
    if not local_cat or "items" not in local_cat:
        return remote_cat

    item_map = {}
    for it in local_cat.get("items", []):
        fn = it.get("filename")
        if fn:
            item_map[fn] = it

    for it in remote_cat.get("items", []):
        fn = it.get("filename")
        if fn:
            # 원격에 더 상세한 정보(file_id 등)가 있으면 우선 병합
            if fn in item_map:
                item_map[fn].update(it)
            else:
                item_map[fn] = it

    merged_items = sorted(item_map.values(), key=lambda x: x.get("date", ""), reverse=True)
    return {
        "version": "1.0",
        "updated_at": datetime.now().isoformat(),
        "channel_id": CHAT_ID,
        "items": merged_items
    }


def publish_catalog(catalog_data):
    """
    현재 카탈로그 데이터를 cloud_catalog.json 파일로 채널에 조용히 업로드합니다.
    - 안내 메시지를 매번 보내거나 고정하지 않습니다 (안내문 중복 발송 원천 차단).
    - 이전 cloud_catalog.json 메시지가 있다면 자동 삭제하여 채널 피드를 항상 깨끗하게 유지합니다.
    - 채널 설명(description)에 최신 카탈로그 file_id를 등록하여 조회하도록 합니다.
    """
    save_local_catalog(catalog_data)
    catalog_bytes = json.dumps(catalog_data, indent=2, ensure_ascii=False).encode("utf-8")

    # 1. 이전 카탈로그 메시지 ID 확인
    prev_msg_id = None
    if os.path.exists(LAST_CATALOG_META_FILE):
        try:
            with open(LAST_CATALOG_META_FILE, "r", encoding="utf-8") as f:
                meta = json.load(f)
                prev_msg_id = meta.get("message_id")
        except Exception:
            pass

    # 2. 파일 업로드 (가이드 안내문 없이 시스템 로그용 캡션만 사용)
    upload_res = upload_multipart_document(
        chat_id=CHAT_ID,
        file_bytes=catalog_bytes,
        filename="cloud_catalog.json",
        caption=f"[시스템] 클라우드 카탈로그 갱신 ({datetime.now().strftime('%Y-%m-%d %H:%M')})"
    )
    if not upload_res.get("ok"):
        return False, upload_res.get("description", "카탈로그 업로드 실패")

    new_msg_id = upload_res["result"].get("message_id")
    doc = upload_res["result"].get("document", {})
    new_file_id = doc.get("file_id")

    if not new_msg_id or not new_file_id:
        return False, "카탈로그 업로드 결과 파싱 실패"

    # 3. 채널 설명(description)에 최신 카탈로그 file_id 안전 등록
    set_channel_catalog_ref(new_file_id)

    # 4. 이전 구형 카탈로그 메시지가 있다면 자동 삭제하여 채널 히스토리 정리
    if prev_msg_id and prev_msg_id != new_msg_id:
        try:
            telegram_api_call("deleteMessage", {"chat_id": CHAT_ID, "message_id": prev_msg_id})
        except Exception:
            pass

    # 5. 상태 파일 갱신
    try:
        with open(LAST_CATALOG_META_FILE, "w", encoding="utf-8") as f:
            json.dump({
                "message_id": new_msg_id,
                "file_id": new_file_id,
                "updated_at": datetime.now().isoformat()
            }, f, indent=2)
    except Exception:
        pass

    return True, new_msg_id


# 하위 호환성 유지용 별칭
publish_and_pin_catalog = publish_catalog


def ensure_channel_guide_message(force=False):
    """
    채널 상단 공식 안내 공지문 관리 함수:
    - 기본적으로 채널에 이미 고정된 안내문이 있으면 아무 작업도 하지 않습니다 (중복 발송/핀 방지).
    - force=True (관리자가 명시적으로 갱신을 요청한 경우)이거나 채널에 고정 안내문이 아예 없을 때만
      단독 텍스트 공지문으로 발송하고 상단에 핀으로 고정합니다.
    """
    chat_res = telegram_api_call("getChat", {"chat_id": CHAT_ID})
    pinned = chat_res.get("result", {}).get("pinned_message")

    # 이미 고정된 안내문이 있고 강제 갱신이 아닌 경우 그대로 유지
    if pinned and not force:
        text = pinned.get("text") or pinned.get("caption") or ""
        if "로고스 찬양 PPT 시스템" in text or "원격 프리셋" in text:
            return True, pinned.get("message_id")

    # 강제 갱신 시 기존 고정 메시지 언핀
    if force and pinned:
        try:
            telegram_api_call("unpinChatMessage", {"chat_id": CHAT_ID, "message_id": pinned.get("message_id")})
        except Exception:
            pass

    # 독립된 공식 공지 텍스트 메시지 발송
    send_res = telegram_api_call("sendMessage", {
        "chat_id": CHAT_ID,
        "text": GUIDE_CAPTION_TEXT
    })
    if not send_res.get("ok"):
        return False, send_res.get("description", "공식 안내문 발송 실패")

    guide_mid = send_res["result"].get("message_id")
    if guide_mid:
        telegram_api_call("pinChatMessage", {
            "chat_id": CHAT_ID,
            "message_id": guide_mid,
            "disable_notification": True
        })
        return True, guide_mid

    return False, "message_id 없음"


def clean_channel_obsolete_messages():
    """
    채널 내의 중복 고정 메시지를 안전하게 일괄 정리합니다.
    - 실제 찬양 PPT 및 찬양곡 자산은 절대로 손상시키지 않습니다.
    - 최신 공식 안내문 1개만 고정 상태로 남겨두고 과거 중복 고정들을 모두 해제합니다.
    - 반환값: 언핀 처리된 개수 (int)
    """
    catalog = load_local_catalog()
    protected_mids = set()
    for it in catalog.get("items", []):
        mid = it.get("message_id")
        if mid:
            try:
                protected_mids.add(int(mid))
            except Exception:
                pass

    if os.path.exists(LAST_CATALOG_META_FILE):
        try:
            with open(LAST_CATALOG_META_FILE, "r", encoding="utf-8") as f:
                meta = json.load(f)
                if meta.get("message_id"):
                    protected_mids.add(int(meta.get("message_id")))
        except Exception:
            pass

    chat_res = telegram_api_call("getChat", {"chat_id": CHAT_ID})
    pinned_mid = chat_res.get("result", {}).get("pinned_message", {}).get("message_id")
    if pinned_mid:
        try:
            protected_mids.add(int(pinned_mid))
        except Exception:
            pass

    unpinned_count = 0
    max_id = max(protected_mids) if protected_mids else 100
    for mid in range(1, max_id + 10):
        if mid != pinned_mid:
            u_res = telegram_api_call("unpinChatMessage", {"chat_id": CHAT_ID, "message_id": mid})
            if u_res.get("ok"):
                unpinned_count += 1

    return unpinned_count



# ==============================================================================
# 3. 실시간 채널 업데이트 감지 및 OTA 프리셋 원격 교체
# ==============================================================================

def get_last_update_id():
    if os.path.exists(LAST_UPDATE_ID_FILE):
        try:
            with open(LAST_UPDATE_ID_FILE, "r", encoding="utf-8") as f:
                return int(f.read().strip())
        except Exception:
            pass
    return 0


def set_last_update_id(update_id):
    try:
        with open(LAST_UPDATE_ID_FILE, "w", encoding="utf-8") as f:
            f.write(str(update_id))
    except Exception:
        pass


PENDING_OTA_FILE = os.path.join(DATA_ROOT, ".pending_ota.json")


def load_pending_ota():
    """진행 중인 텔레그램 인라인 버튼 인터랙션 세션을 로드합니다."""
    if os.path.exists(PENDING_OTA_FILE):
        try:
            with open(PENDING_OTA_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_pending_ota(data):
    """진행 중인 텔레그램 세션을 저장하고 1시간 지난 오래된 항목을 자동 정리합니다."""
    try:
        now = time.time()
        cleaned = {k: v for k, v in data.items() if now - v.get("timestamp", 0) < 3600}
        with open(PENDING_OTA_FILE, "w", encoding="utf-8") as f:
            json.dump(cleaned, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def get_available_preset_list():
    """
    현재 시스템 및 클라우드 카탈로그에 등록된 '공식 프리셋' 순수 이름 목록을 반환합니다.
    예: ['청년부', '중고등부']
    """
    return get_official_departments()


def extract_dept_name_from_preset(preset_name):
    """프리셋 폴더/이름에서 순수 프리셋명을 추출합니다. 예: '[청년부] 기준 PPT' 또는 '청년부' -> '청년부'"""
    if not preset_name:
        return "기타"
    cleaned = re.sub(r'[\[\]]', '', str(preset_name)).replace("기준 PPT", "").replace("기준PPT", "").strip()
    return cleaned if cleaned else "기타"


def check_and_process_channel_updates():
    """
    getUpdates를 호출하여 채널에 새로 게시된 메시지(OTA 태그, 파일 업로드, 인라인 버튼 클릭 등)를 스캔하고 처리합니다.
    - #기준 PPT / #프리셋 / #템플릿: 기준 PPTX 자동 교체 (불일치 시 인라인 버튼 선택 제공)
    - #표지: 부서별 표지 이미지 자동 교체 (불일치 시 인라인 버튼 선택 제공)
    - callback_query: 스마트폰에서 누른 인라인 버튼 이벤트 수신 및 자동 적용
    - 번호 답장(1, 2 등): 텍스트 숫자로도 프리셋 선택 가능
    - 표준 파일명 첨부파일: 카탈로그 등록 및 찬양곡 자동 다운로드
    반환값: {"processed_count": int, "ota_presets": list, "new_files": list}
    """
    last_id = get_last_update_id()
    params = {"allowed_updates": ["channel_post", "edited_channel_post", "message", "callback_query"]}
    if last_id > 0:
        params["offset"] = last_id + 1

    updates_res = telegram_api_call("getUpdates", params, timeout=30)
    if not updates_res.get("ok"):
        return {"processed_count": 0, "ota_presets": [], "new_files": [], "error": updates_res.get("error")}

    updates = updates_res.get("result", [])
    if not updates:
        return {"processed_count": 0, "ota_presets": [], "new_files": []}

    ota_presets = []
    new_files = []
    max_id = last_id

    catalog = load_local_catalog()
    catalog_dirty = False
    available_presets = get_available_preset_list()

    for upd in updates:
        upd_id = upd.get("update_id", 0)
        if upd_id > max_id:
            max_id = upd_id

        # -------------------------------------------------------------
        # 1. 텔레그램 인라인 버튼(터치) 클릭 이벤트 (callback_query) 처리
        # -------------------------------------------------------------
        cb_query = upd.get("callback_query")
        if cb_query:
            cq_id = cb_query.get("id")
            cq_data = cb_query.get("data", "")
            cq_msg = cb_query.get("message", {})
            cq_mid = cq_msg.get("message_id")
            cq_chat_id = cq_msg.get("chat", {}).get("id") or CHAT_ID

            if cq_data.startswith("ota:"):
                parts = cq_data.split(":")
                if len(parts) == 3:
                    _, req_id, idx_str = parts
                    idx = int(idx_str)
                    pending = load_pending_ota()
                    req = pending.get(req_id)
                    if req and 0 <= idx < len(available_presets):
                        target_preset = available_presets[idx]
                        dept = extract_dept_name_from_preset(target_preset)
                        req_type = req.get("type", "cover")
                        file_id = req.get("file_id")
                        file_name = req.get("file_name", "")

                        dept_dir = os.path.join(DIR_PRESETS, dept)
                        os.makedirs(dept_dir, exist_ok=True)

                        if req_type == "cover":
                            _, ext = os.path.splitext(file_name)
                            if not ext or ext.lower() not in [".png", ".jpg", ".jpeg"]:
                                ext = ".png"
                            target_file = os.path.join(dept_dir, f"{dept} 표지_{datetime.now().strftime('%y%m%d_%H%M%S')}{ext}")
                            ok, err = download_file_by_id(file_id, target_file)
                            if ok:
                                ota_presets.append({"type": "cover", "dept": dept, "path": target_file})
                                if "presets" not in catalog:
                                    catalog["presets"] = {}
                                if dept not in catalog["presets"]:
                                    catalog["presets"][dept] = {}
                                catalog["presets"][dept]["cover_file_id"] = file_id
                                catalog["presets"][dept]["cover_filename"] = os.path.basename(target_file)
                                catalog["presets"][dept]["cover_updated_at"] = datetime.now().isoformat()
                                catalog_dirty = True

                                # 버튼 메시지를 깔끔한 완료 확인문으로 교체
                                telegram_api_call("editMessageText", {
                                    "chat_id": cq_chat_id,
                                    "message_id": cq_mid,
                                    "text": f"✅ {idx + 1}. [{dept}] 기준 표지를 성공적으로 적용하였습니다!"
                                })
                                telegram_api_call("answerCallbackQuery", {
                                    "callback_query_id": cq_id,
                                    "text": f"[{dept}] 기준 표지 적용 완료!"
                                })
                                del pending[req_id]
                                save_pending_ota(pending)

                        elif req_type == "preset":
                            target_file = os.path.join(dept_dir, f"{dept} 기준 PPT.pptx")
                            # 기존 템플릿 백업 보존
                            if os.path.exists(target_file):
                                bk_name = f"{dept} 기준 PPT_이전_{datetime.now().strftime('%y%m%d_%H%M%S')}.pptx"
                                try:
                                    shutil.copy2(target_file, os.path.join(dept_dir, bk_name))
                                except Exception:
                                    pass
                            ok, err = download_file_by_id(file_id, target_file)
                            if ok:
                                ota_presets.append({"type": "preset", "dept": dept, "path": target_file})
                                if "presets" not in catalog:
                                    catalog["presets"] = {}
                                if dept not in catalog["presets"]:
                                    catalog["presets"][dept] = {}
                                catalog["presets"][dept]["template_file_id"] = file_id
                                catalog["presets"][dept]["template_filename"] = f"{dept} 기준 PPT.pptx"
                                catalog["presets"][dept]["template_updated_at"] = datetime.now().isoformat()
                                catalog_dirty = True

                                telegram_api_call("editMessageText", {
                                    "chat_id": cq_chat_id,
                                    "message_id": cq_mid,
                                    "text": f"✅ {idx + 1}. [{dept}] 기준 PPT 서식을 성공적으로 적용하였습니다!"
                                })
                                telegram_api_call("answerCallbackQuery", {
                                    "callback_query_id": cq_id,
                                    "text": f"[{dept}] 기준 PPT 갱신 완료!"
                                })
                                del pending[req_id]
                                save_pending_ota(pending)
                    else:
                        telegram_api_call("answerCallbackQuery", {
                            "callback_query_id": cq_id,
                            "text": "만료되었거나 이미 처리된 요청입니다."
                        })
            continue

        # -------------------------------------------------------------
        # 2. 일반 채널 포스트 및 메시지 수신 처리
        # -------------------------------------------------------------
        post = upd.get("channel_post") or upd.get("edited_channel_post") or upd.get("message")
        if not post:
            continue

        caption = (post.get("caption", "") or post.get("text", "")).strip()
        doc = post.get("document")
        photo_list = post.get("photo")
        source_chat_id = post.get("chat", {}).get("id") or CHAT_ID
        post_mid = post.get("message_id")

        # (A) 숫자로 번호를 답장한 경우 (예: '1', '2')
        if caption.isdigit():
            pending = load_pending_ota()
            if pending:
                latest_req_id = max(pending.keys(), key=lambda k: pending[k].get("timestamp", 0))
                req = pending[latest_req_id]
                num_idx = int(caption) - 1
                if 0 <= num_idx < len(available_presets):
                    target_preset = available_presets[num_idx]
                    dept = extract_dept_name_from_preset(target_preset)
                    req_type = req.get("type", "cover")
                    file_id = req.get("file_id")
                    file_name = req.get("file_name", "")
                    dept_dir = os.path.join(DIR_PRESETS, dept)
                    os.makedirs(dept_dir, exist_ok=True)

                    if req_type == "cover":
                        _, ext = os.path.splitext(file_name)
                        if not ext or ext.lower() not in [".png", ".jpg", ".jpeg"]:
                            ext = ".png"
                        target_file = os.path.join(dept_dir, f"{dept} 표지_{datetime.now().strftime('%y%m%d_%H%M%S')}{ext}")
                        ok, err = download_file_by_id(file_id, target_file)
                        if ok:
                            ota_presets.append({"type": "cover", "dept": dept, "path": target_file})
                            if "presets" not in catalog:
                                catalog["presets"] = {}
                            if dept not in catalog["presets"]:
                                catalog["presets"][dept] = {}
                            catalog["presets"][dept]["cover_file_id"] = file_id
                            catalog["presets"][dept]["cover_filename"] = os.path.basename(target_file)
                            catalog["presets"][dept]["cover_updated_at"] = datetime.now().isoformat()
                            catalog_dirty = True
                            telegram_api_call("sendMessage", {
                                "chat_id": source_chat_id,
                                "text": f"✅ {num_idx + 1}. [{dept}] 기준 표지를 성공적으로 적용하였습니다!",
                                "reply_to_message_id": post_mid
                            })
                            del pending[latest_req_id]
                            save_pending_ota(pending)
                            continue
                    elif req_type == "preset":
                        target_file = os.path.join(dept_dir, f"{dept} 기준 PPT.pptx")
                        if os.path.exists(target_file):
                            bk_name = f"{dept} 기준 PPT_이전_{datetime.now().strftime('%y%m%d_%H%M%S')}.pptx"
                            try:
                                shutil.copy2(target_file, os.path.join(dept_dir, bk_name))
                            except Exception:
                                pass
                        ok, err = download_file_by_id(file_id, target_file)
                        if ok:
                            ota_presets.append({"type": "preset", "dept": dept, "path": target_file})
                            if "presets" not in catalog:
                                catalog["presets"] = {}
                            if dept not in catalog["presets"]:
                                catalog["presets"][dept] = {}
                            catalog["presets"][dept]["template_file_id"] = file_id
                            catalog["presets"][dept]["template_filename"] = f"{dept} 기준 PPT.pptx"
                            catalog["presets"][dept]["template_updated_at"] = datetime.now().isoformat()
                            catalog_dirty = True
                            telegram_api_call("sendMessage", {
                                "chat_id": source_chat_id,
                                "text": f"✅ {num_idx + 1}. [{dept}] 기준 PPT 서식을 성공적으로 적용하였습니다!",
                                "reply_to_message_id": post_mid
                            })
                            del pending[latest_req_id]
                            save_pending_ota(pending)
                            continue

        # (B) 신규 공식 프리셋 생성 명령 (#추가 : [프리셋 이름]) - 첨부파일 없는 일반 텍스트 명령어
        add_match = re.search(r'#추가(?:\s*:\s*|\s*:|\s+)(?:\[)?([^\]\n]+?)(?:\])?(?:\s|$)', caption)
        if add_match:
            raw_name = add_match.group(1).strip()
            preset_name = re.sub(r'[\[\]]', '', raw_name).replace("기준 PPT", "").replace("기준PPT", "").strip()
            if preset_name:
                # 1. 로컬에 순수 프리셋 이름으로 폴더 생성
                dept_dir = os.path.join(DIR_PRESETS, preset_name)
                os.makedirs(dept_dir, exist_ok=True)

                # 2. 카탈로그에 공식 프리셋 등록
                if "presets" not in catalog:
                    catalog["presets"] = {}
                if preset_name not in catalog["presets"]:
                    catalog["presets"][preset_name] = {
                        "created_at": datetime.now().isoformat()
                    }
                    catalog_dirty = True

                available_presets = get_available_preset_list()

                telegram_api_call("sendMessage", {
                    "chat_id": source_chat_id,
                    "text": f"✨ '{preset_name}' 공식 프리셋이 생성되었습니다!\n\n"
                            f"📁 생성 폴더: {preset_name}\n\n"
                            f"💡 이제 이 프리셋에 서식을 등록하실 수 있습니다:\n"
                            f"• 기준 PPT: .pptx 파일 첨부 + 캡션 #기준\n"
                            f"• 표지 이미지: 이미지 첨부 + 캡션 #표지",
                    "reply_to_message_id": post_mid
                })
                continue
            else:
                telegram_api_call("sendMessage", {
                    "chat_id": source_chat_id,
                    "text": "⚠️ 추가할 프리셋 이름을 입력해 주세요.\n예: #추가 : 청년부 또는 #추가 : 중고등부",
                    "reply_to_message_id": post_mid
                })
                continue

        # 첨부 파일/사진 정보 추출
        file_id = None
        file_name = None
        file_size = 0
        if doc:
            file_id = doc.get("file_id")
            file_name = doc.get("file_name")
            file_size = doc.get("file_size", 0)
        elif photo_list and isinstance(photo_list, list) and len(photo_list) > 0:
            best_photo = photo_list[-1]
            file_id = best_photo.get("file_id")
            file_name = f"cover_{post.get('date', int(time.time()))}.jpg"
            file_size = best_photo.get("file_size", 0)

        if not file_id:
            continue

        # (B) OTA 표지 이미지 감지 (#표지)
        cover_match = re.search(r'#표지(?:\s*\[?([^\]\s]+)\]?)?', caption)
        is_image_asset = photo_list or (file_name and any(file_name.lower().endswith(ext) for ext in [".png", ".jpg", ".jpeg"]))
        if cover_match and is_image_asset:
            input_dept = cover_match.group(1)
            matched_preset = None
            if input_dept:
                input_dept = input_dept.strip()
                for p in available_presets:
                    d_name = extract_dept_name_from_preset(p)
                    if input_dept.lower() == d_name.lower() or input_dept.lower() == p.lower():
                        matched_preset = p
                        break

            if matched_preset:
                dept = extract_dept_name_from_preset(matched_preset)
                dept_dir = os.path.join(DIR_PRESETS, dept)
                os.makedirs(dept_dir, exist_ok=True)
                _, ext = os.path.splitext(file_name or "")
                if not ext or ext.lower() not in [".png", ".jpg", ".jpeg"]:
                    ext = ".png"
                target_file = os.path.join(dept_dir, f"{dept} 표지_{datetime.now().strftime('%y%m%d_%H%M%S')}{ext}")
                ok, err = download_file_by_id(file_id, target_file)
                if ok:
                    ota_presets.append({"type": "cover", "dept": dept, "path": target_file})
                    print(f"[CloudSync] OTA 표지 직접 갱신 성공: {target_file}")
                    if "presets" not in catalog:
                        catalog["presets"] = {}
                    if dept not in catalog["presets"]:
                        catalog["presets"][dept] = {}
                    catalog["presets"][dept]["cover_file_id"] = file_id
                    catalog["presets"][dept]["cover_filename"] = os.path.basename(target_file)
                    catalog["presets"][dept]["cover_updated_at"] = datetime.now().isoformat()
                    catalog_dirty = True
                    try:
                        telegram_api_call("sendMessage", {
                            "chat_id": source_chat_id,
                            "text": f"✅ [{dept}] 기준 표지 이미지가 정상 감지되어 최신 표지로 적용되었습니다!",
                            "reply_to_message_id": post_mid
                        })
                    except Exception:
                        pass
                continue
            else:
                if not available_presets:
                    telegram_api_call("sendMessage", {
                        "chat_id": source_chat_id,
                        "text": "⚠️ 현재 등록된 공식 프리셋이 없습니다.\n먼저 일반 텍스트로 #추가 : [프리셋 이름]을 전송하여 프리셋을 생성해 주세요.\n(예: #추가 : 청년부)",
                        "reply_to_message_id": post_mid
                    })
                    continue

                req_id = str(post_mid)
                pending = load_pending_ota()
                pending[req_id] = {
                    "type": "cover",
                    "file_id": file_id,
                    "file_name": file_name,
                    "source_chat_id": source_chat_id,
                    "message_id": post_mid,
                    "timestamp": time.time()
                }
                save_pending_ota(pending)

                guide_lines = [
                    "🖼️ 표지 이미지를 인식했습니다. 업로드할 프리셋을 선택해 주세요:\n"
                ]
                keyboard = []
                for idx, p in enumerate(available_presets):
                    guide_lines.append(f"{idx + 1}. {p}")
                    keyboard.append([{"text": f"{idx + 1}. {p}", "callback_data": f"ota:{req_id}:{idx}"}])

                telegram_api_call("sendMessage", {
                    "chat_id": source_chat_id,
                    "text": "\n".join(guide_lines),
                    "reply_to_message_id": post_mid,
                    "reply_markup": {"inline_keyboard": keyboard}
                })
                continue

        # (C) OTA 기준 PPT 템플릿 감지 (#기준 PPT, #기준, #프리셋, #템플릿)
        preset_match = re.search(r'#(?:기준\s*PPT|기준|템플릿|프리셋)(?:\s*\[?([^\]\s]+)\]?)?', caption)
        if preset_match and file_name and file_name.lower().endswith(".pptx"):
            input_dept = preset_match.group(1)
            matched_preset = None
            if input_dept:
                input_dept = input_dept.strip()
                for p in available_presets:
                    d_name = extract_dept_name_from_preset(p)
                    if input_dept.lower() == d_name.lower() or input_dept.lower() == p.lower():
                        matched_preset = p
                        break

            if matched_preset:
                dept = extract_dept_name_from_preset(matched_preset)
                dept_dir = os.path.join(DIR_PRESETS, dept)
                os.makedirs(dept_dir, exist_ok=True)
                target_file = os.path.join(dept_dir, f"{dept} 기준 PPT.pptx")
                if os.path.exists(target_file):
                    bk_name = f"{dept} 기준 PPT_이전_{datetime.now().strftime('%y%m%d_%H%M%S')}.pptx"
                    try:
                        shutil.copy2(target_file, os.path.join(dept_dir, bk_name))
                    except Exception:
                        pass
                ok, err = download_file_by_id(file_id, target_file)
                if ok:
                    ota_presets.append({"type": "preset", "dept": dept, "path": target_file})
                    print(f"[CloudSync] OTA 기준 PPT 직접 갱신 성공: {target_file}")
                    if "presets" not in catalog:
                        catalog["presets"] = {}
                    if dept not in catalog["presets"]:
                        catalog["presets"][dept] = {}
                    catalog["presets"][dept]["template_file_id"] = file_id
                    catalog["presets"][dept]["template_filename"] = f"{dept} 기준 PPT.pptx"
                    catalog["presets"][dept]["template_updated_at"] = datetime.now().isoformat()
                    catalog_dirty = True
                    try:
                        telegram_api_call("sendMessage", {
                            "chat_id": source_chat_id,
                            "text": f"✅ [{dept}] 기준 PPT 서식이 정상 감지되어 최신 서식으로 적용되었습니다!",
                            "reply_to_message_id": post_mid
                        })
                    except Exception:
                        pass
                continue
            else:
                if not available_presets:
                    telegram_api_call("sendMessage", {
                        "chat_id": source_chat_id,
                        "text": "⚠️ 현재 등록된 공식 프리셋이 없습니다.\n먼저 일반 텍스트로 #추가 : [프리셋 이름]을 전송하여 프리셋을 생성해 주세요.\n(예: #추가 : 청년부)",
                        "reply_to_message_id": post_mid
                    })
                    continue

                req_id = str(post_mid)
                pending = load_pending_ota()
                pending[req_id] = {
                    "type": "preset",
                    "file_id": file_id,
                    "file_name": file_name,
                    "source_chat_id": source_chat_id,
                    "message_id": post_mid,
                    "timestamp": time.time()
                }
                save_pending_ota(pending)

                guide_lines = [
                    "📄 기준 PPT 파일을 인식했습니다. 업로드할 프리셋을 선택해 주세요:\n"
                ]
                keyboard = []
                for idx, p in enumerate(available_presets):
                    guide_lines.append(f"{idx + 1}. {p}")
                    keyboard.append([{"text": f"{idx + 1}. {p}", "callback_data": f"ota:{req_id}:{idx}"}])

                telegram_api_call("sendMessage", {
                    "chat_id": source_chat_id,
                    "text": "\n".join(guide_lines),
                    "reply_to_message_id": post_mid,
                    "reply_markup": {"inline_keyboard": keyboard}
                })
                continue

        # (D) 표준 파일명 자산 감지
        meta = parse_standard_filename(file_name) if file_name else None
        if meta:
            item_entry = {
                "filename": file_name,
                "file_id": file_id,
                "file_size": file_size,
                "date": meta["date"],
                "type": meta["type"],
                "preset": meta["preset"],
                "title": meta["title"],
                "ext": meta["ext"],
                "message_id": post.get("message_id"),
                "caption": caption
            }
            existing = [it for it in catalog["items"] if it.get("filename") == file_name]
            if not existing:
                catalog["items"].append(item_entry)
                catalog_dirty = True
                new_files.append(file_name)

            # 찬양곡(.json)인 경우 로컬 라이브러리에 즉시 자동 다운로드
            if meta["type"] == "찬양곡" and meta["ext"] == "json":
                song_folder = os.path.join(DIR_SONGS, sanitize_filename_part(meta["title"]))
                os.makedirs(song_folder, exist_ok=True)
                local_song_path = os.path.join(song_folder, file_name)
                if not os.path.exists(local_song_path):
                    download_file_by_id(file_id, local_song_path)

    if catalog_dirty:
        catalog["items"] = sorted(catalog["items"], key=lambda x: x.get("date", ""), reverse=True)
        save_local_catalog(catalog)
        publish_catalog(catalog)

    if max_id > last_id:
        set_last_update_id(max_id)

    return {
        "processed_count": len(updates),
        "ota_presets": ota_presets,
        "new_files": new_files
    }


# ==============================================================================
# 4. 고수준 동기화 라이프사이클 API
# ==============================================================================

def sync_on_startup(department=DEFAULT_DEPARTMENT):
    """
    프로그램 시작 시 호출되는 원클릭 동기화 워크플로우:
    1. 데이터 보관함 4대 디렉토리 보장
    2. 원격 고정 카탈로그 확인 및 로컬 카탈로그 병합
    3. 채널 최신 업데이트 및 OTA 프리셋 확인
    4. 미보유 찬양곡 라이브러리 자동 다운로드
    5. 원격 최신 프리셋/표지 자동 동기화 (새 PC 설치 시 최신 표지만 자동 수신)
    6. 90일 경과한 로컬 구형 PPTX 캐시 롤링 정리 (기본 비활성화)
    반환값: dict (동기화 결과 요약)
    """
    ensure_data_directories()
    result = {
        "connected": False,
        "merged_items_count": 0,
        "ota_presets_count": 0,
        "downloaded_songs_count": 0,
        "pruned_pptx_count": 0,
        "error": None
    }

    # 1. 봇 연결 및 원격 카탈로그 조회
    remote_cat = fetch_remote_catalog()
    local_cat = load_local_catalog()

    if remote_cat is not None:
        result["connected"] = True
        merged = merge_catalogs(local_cat, remote_cat)
        # 원격 presets 정보 보존
        if "presets" in remote_cat:
            if "presets" not in merged:
                merged["presets"] = {}
            merged["presets"].update(remote_cat["presets"])
        save_local_catalog(merged)
        result["merged_items_count"] = len(merged.get("items", []))
    else:
        # 핀된 카탈로그가 없는 경우, 로컬 카탈로그를 원격에 게시 및 핀
        test_call = telegram_api_call("getMe")
        if test_call.get("ok"):
            result["connected"] = True
            publish_and_pin_catalog(local_cat)
        else:
            result["error"] = test_call.get("error", "텔레그램 봇에 연결할 수 없습니다.")
            # 오프라인 모드로 안전 폴백
            return result

    # 2. 채널 신규 업데이트 및 OTA 패치 스캔
    upd_res = check_and_process_channel_updates()
    result["ota_presets_count"] = len(upd_res.get("ota_presets", []))

    # 3. 로컬에 없는 찬양곡 라이브러리 자동 다운로드
    cat = load_local_catalog()
    dl_count = 0
    for item in cat.get("items", []):
        if item.get("type") == "찬양곡" and item.get("file_id"):
            song_title = sanitize_filename_part(item.get("title", "미지정"))
            filename = item.get("filename")
            song_dir = os.path.join(DIR_SONGS, song_title)
            dest_file = os.path.join(song_dir, filename)
            if not os.path.exists(dest_file):
                ok, _ = download_file_by_id(item["file_id"], dest_file)
                if ok:
                    dl_count += 1
    result["downloaded_songs_count"] = dl_count

    # 3-2. 원격 최신 프리셋/표지 자동 동기화 (새 PC 설치 시 최신 표지만 자동 수신)
    remote_presets = cat.get("presets", {})
    for p_dept, p_info in remote_presets.items():
        dept_name = extract_dept_name_from_preset(p_dept)
        dept_dir = os.path.join(DIR_PRESETS, dept_name)
        os.makedirs(dept_dir, exist_ok=True)
        # 최신 표지 이미지 확인
        c_fid = p_info.get("cover_file_id")
        c_fname = p_info.get("cover_filename")
        if c_fid and c_fname:
            c_dest = os.path.join(dept_dir, c_fname)
            if not os.path.exists(c_dest):
                ok, _ = download_file_by_id(c_fid, c_dest)
                if ok:
                    result["ota_presets_count"] += 1
        # 최신 기준 PPT 확인
        t_fid = p_info.get("template_file_id")
        t_fname = p_info.get("template_filename")
        if t_fid and t_fname:
            t_dest = os.path.join(dept_dir, t_fname)
            if not os.path.exists(t_dest):
                ok, _ = download_file_by_id(t_fid, t_dest)
                if ok:
                    result["ota_presets_count"] += 1

    # 4. 다운로드 후 로컬 중복 찬양곡 검사 및 최신 일자 단일 버전 정리
    try:
        from core.song_library import deduplicate_all_library_songs
        deduplicate_all_library_songs()
    except Exception:
        pass

    # 5. 90일 경과 구형 로컬 PPTX 캐시 롤링 (사용자 피드백에 따라 기본 비활성화)
    pruned = prune_local_old_pptx(days=90, enabled=False)
    result["pruned_pptx_count"] = pruned

    return result


def sync_upload_created_assets(pptx_path, song_paths, department=DEFAULT_DEPARTMENT, date_str=None):
    """
    [PPT 파일 생성] 시 생성된 PPTX 파일과 분해된 개별 찬양곡 JSON 파일들을
    텔레그램 채널로 일괄 전송하고 클라우드 카탈로그를 최신화하여 핀합니다.
    """
    ensure_data_directories()
    catalog = load_local_catalog()
    uploaded_records = []
    errors = []

    if not date_str:
        date_str = datetime.now().strftime("%Y%m%d")
    date_clean = re.sub(r"[^0-9]", "", str(date_str))
    if len(date_clean) < 8:
        date_clean = datetime.now().strftime("%Y%m%d")

    # 1. PPTX 파일 업로드
    if pptx_path and os.path.exists(pptx_path):
        fn = os.path.basename(pptx_path)
        meta = parse_standard_filename(fn)
        cap = f"📊 [찬양 PPT] {meta['title'] if meta else fn}\n부서: {department} | 날짜: {date_clean}"
        ok, res, err = upload_file_to_telegram(pptx_path, fn, caption=cap)
        if ok:
            doc = res.get("document", {})
            record = {
                "filename": fn,
                "file_id": doc.get("file_id"),
                "file_size": doc.get("file_size", os.path.getsize(pptx_path)),
                "date": meta["date"] if meta else date_clean,
                "type": "PPT",
                "preset": department,
                "title": meta["title"] if meta else fn,
                "ext": "pptx",
                "message_id": res.get("message_id"),
                "caption": cap
            }
            uploaded_records.append(record)
        else:
            errors.append(f"PPTX 업로드 실패: {err}")

    # 2. 분해된 찬양곡 JSON 업로드
    for sp in song_paths:
        if os.path.exists(sp):
            s_fn = os.path.basename(sp)
            s_meta = parse_standard_filename(s_fn)
            s_title = s_meta["title"] if s_meta else os.path.splitext(s_fn)[0]
            cap = f"🎵 [찬양곡 자산] {s_title}\n부서: {department} | 일자: {date_clean}"
            ok, res, err = upload_file_to_telegram(sp, s_fn, caption=cap)
            if ok:
                doc = res.get("document", {})
                record = {
                    "filename": s_fn,
                    "file_id": doc.get("file_id"),
                    "file_size": doc.get("file_size", os.path.getsize(sp)),
                    "date": s_meta["date"] if s_meta else date_clean,
                    "type": "찬양곡",
                    "preset": department,
                    "title": s_title,
                    "ext": "json",
                    "message_id": res.get("message_id"),
                    "caption": cap
                }
                uploaded_records.append(record)
            else:
                errors.append(f"찬양곡({s_fn}) 업로드 실패: {err}")

    # 3. 카탈로그 병합 및 채널 상단 핀 갱신
    if uploaded_records:
        existing_filenames = {it.get("filename") for it in catalog["items"]}
        for r in uploaded_records:
            if r["filename"] in existing_filenames:
                # 갱신
                for idx, old_it in enumerate(catalog["items"]):
                    if old_it.get("filename") == r["filename"]:
                        catalog["items"][idx] = r
                        break
            else:
                # 찬양곡인 경우: 로컬에서 날짜 갱신 또는 버전 정리로 이미 삭제된 구 버전 파일이 카탈로그에 남아있다면 정리
                if r.get("type") == "찬양곡":
                    song_title = r.get("title", "")
                    clean_title = re.sub(r'[\\/*?:"<>|]', '', song_title).strip()
                    song_dir = os.path.join(DIR_SONGS, clean_title)
                    local_files = set(os.listdir(song_dir)) if os.path.exists(song_dir) else set()
                    
                    obsolete_indices = []
                    for idx, old_it in enumerate(catalog["items"]):
                        if old_it.get("type") == "찬양곡" and old_it.get("title") == song_title:
                            old_fn = old_it.get("filename", "")
                            # 로컬 폴더에 더 이상 존재하지 않는 구 파일이면 클라우드에서도 정리
                            if old_fn not in local_files:
                                obsolete_indices.append(idx)
                                old_mid = old_it.get("message_id")
                                if old_mid:
                                    try:
                                        telegram_api_call("deleteMessage", {"chat_id": CHAT_ID, "message_id": old_mid})
                                    except Exception:
                                        pass
                                        
                    for idx in sorted(obsolete_indices, reverse=True):
                        catalog["items"].pop(idx)

                catalog["items"].append(r)

        catalog["items"] = sorted(catalog["items"], key=lambda x: x.get("date", ""), reverse=True)
        # 핀된 카탈로그 문서 업데이트
        publish_and_pin_catalog(catalog)

    return {
        "success": len(errors) == 0,
        "uploaded_count": len(uploaded_records),
        "records": uploaded_records,
        "errors": errors
    }


def prune_local_old_pptx(days=90, enabled=False):
    """
    데이터 보관함/2_찬양_PPT/ 폴더에서 90일(기본값)이 경과한 구형 PPTX를 안전하게 정리합니다.
    - 사용자 피드백에 따라 기본값은 자동 삭제 비활성화(enabled=False, 로컬 영구 보존).
    - 활성화 시에도 클라우드 카탈로그에 정상 업로드(file_id 보유)된 파일만 안전하게 정리합니다.
    - 3_찬양곡_라이브러리와 4_찬양_콘티는 절대로 삭제하지 않습니다.
    반환값: 삭제된 파일 개수
    """
    if not enabled:
        return 0

    ensure_data_directories()
    if not os.path.exists(DIR_PPTX):
        return 0

    cutoff_date = datetime.now() - timedelta(days=days)
    pruned_count = 0

    catalog = load_local_catalog()
    cloud_backed_files = {it.get("filename") for it in catalog.get("items", []) if it.get("type") == "PPT" and it.get("file_id")}

    for fname in os.listdir(DIR_PPTX):
        if not fname.lower().endswith(".pptx"):
            continue

        # 클라우드에 백업되지 않은 로컬 단독 파일은 절대로 자동 삭제하지 않음
        if fname not in cloud_backed_files:
            continue

        fpath = os.path.join(DIR_PPTX, fname)
        if not os.path.isfile(fpath):
            continue

        meta = parse_standard_filename(fname)
        file_dt = None
        if meta and meta.get("date"):
            try:
                file_dt = datetime.strptime(meta["date"], "%Y%m%d")
            except Exception:
                pass

        if file_dt is None:
            # 파일 수정 시각 기준
            mtime = os.path.getmtime(fpath)
            file_dt = datetime.fromtimestamp(mtime)

        if file_dt < cutoff_date:
            try:
                os.remove(fpath)
                pruned_count += 1
                print(f"[CloudSync] 90일 경과 로컬 캐시 정리 완료: {fname}")
            except Exception as e:
                print(f"[CloudSync] 파일 정리 실패 ({fname}): {e}")

    return pruned_count
