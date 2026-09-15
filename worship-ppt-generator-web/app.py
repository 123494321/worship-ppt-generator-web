import streamlit as st
import io
import sys
import os
import re
import json
import time
import shutil
import subprocess
from datetime import datetime


def get_user_desktop_path():
    """사용자의 Windows 바탕화면 경로를 감지합니다."""
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders')
        desktop, _ = winreg.QueryValueEx(key, 'Desktop')
        winreg.CloseKey(key)
        desktop = os.path.expandvars(desktop)
        if os.path.exists(desktop):
            return desktop
    except Exception:
        pass
    for p in [r'D:\Desktop', os.path.join(os.path.expanduser('~'), 'Desktop')]:
        if os.path.exists(p):
            return p
    return os.path.join(os.path.expanduser('~'), 'Desktop')

def open_folder_in_explorer(file_or_dir_path):
    """지정된 파일 또는 폴더의 위치를 윈도우 탐색기로 엽니다."""
    try:
        path = os.path.normpath(os.path.abspath(file_or_dir_path))
        folder = path if os.path.isdir(path) else os.path.dirname(path)
        if not os.path.exists(folder):
            os.makedirs(folder, exist_ok=True)

        # 1. 윈도우 표준 탐색기 직접 실행 (명시적 경로 호출로 항상 새 탐색기 창 오픈)
        try:
            subprocess.Popen(f'explorer "{folder}"')
            return True
        except Exception:
            pass

        # 2. Windows 기본 os.startfile fallback
        try:
            os.startfile(folder)
            return True
        except Exception:
            pass

        return False
    except Exception:
        return False


from core.slide_editor_component import render_slide_ide_editor
from core.conti_editor_component import render_conti_editor
from core.text_engine import (
    parse_user_conti, 
    generate_plan_text_from_conti, 
    parse_plan_text_to_slides,
    get_upcoming_sunday,
    get_default_conti_template,
    validate_conti_text
)
from core.cloud_config import (
    DIR_CONTI,
    DIR_PPTX,
    DIR_SONGS,
    DIR_PRESETS,
    TELEGRAM_CHANNEL_LINK,
    format_standard_filename,
    parse_standard_filename,
    sanitize_filename_part,
    normalize_department_name,
    ensure_data_directories
)
from core.cloud_sync import (
    sync_on_startup,
    sync_upload_created_assets,
    check_and_process_channel_updates,
    load_local_catalog,
    fetch_remote_catalog,
    download_file_by_id,
    prune_local_old_pptx,
    ensure_channel_guide_message,
    clean_channel_obsolete_messages
)
from core.pptx_builder import (
    build_praise_pptx, 
    get_pptx_bytes, 
    get_preset_ppts, 
    get_preset_info_dict,
    find_preset_cover_image, 
    extract_cover_image_from_template,
    save_pptx_to_archive,
    PRESET_DIR
)
from core.song_library import (
    get_all_library_songs,
    search_library,
    search_library_with_version_limit,
    get_song_version,
    save_all_songs_from_conti,
    decompose_and_save_songs_from_conti,
    format_song_to_conti_text,
    replace_song_in_conti,
    append_song_to_conti
)

st.set_page_config(
    page_title="LOGOS 찬양 PPT 제작 스튜디오",
    page_icon="🎵",
    layout="wide",
    initial_sidebar_state="expanded"
)

from core.pin_gate_component import render_pin_gate
from core.admin_component import render_admin_console
from core.trash_manager import is_pptx_trashed, is_song_trashed

# ==========================================
# 0. 찬양팀 접속 인증 비밀번호 게이트 (일반 0000 / 관리자 7777)
# ==========================================
try:
    ADMIN_PIN = str(st.secrets.get("ADMIN_PIN", "7777"))
except Exception:
    ADMIN_PIN = "7777"

if not st.session_state.get("authenticated", False):
    st.markdown("""
    <style>
        @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
        * { font-family: 'Pretendard', sans-serif; }
        /* Hide sidebar on authentication gate */
        [data-testid="stSidebar"], [data-testid="collapsedControl"] {
            display: none !important;
        }
        .main .block-container {
            padding-top: 5rem;
            max-width: 600px;
        }
    </style>
    """, unsafe_allow_html=True)
    
    try:
        entered_pin = render_pin_gate(key="logos_pin_gate", allowed_pins=["0000", ADMIN_PIN])
    except TypeError:
        # 하위 호환성 방어: core/pin_gate_component.py가 구버전인 경우 예외 없이 안전 호출
        entered_pin = render_pin_gate(key="logos_pin_gate")

    if entered_pin in ("0000", True):
        st.session_state.authenticated = True
        st.session_state.is_admin = False
        st.session_state.app_mode = "general"
        st.rerun()
    elif entered_pin == ADMIN_PIN:
        st.session_state.authenticated = True
        st.session_state.is_admin = True
        st.session_state.app_mode = "admin"
        st.rerun()
    st.stop()


# Custom Styling (In-app styling)
st.markdown("""
<style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    * {
        font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif;
    }
    .main-title-container {
        display: flex;
        align-items: center;
        justify-content: space-between;
        background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%);
        padding: 18px 24px;
        border-radius: 12px;
        color: white;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .main-title {
        font-size: 1.65rem;
        font-weight: 800;
        letter-spacing: -0.5px;
    }
    .main-subtitle {
        font-size: 0.95rem;
        opacity: 0.9;
    }
    
    /* 0. Comfortable Compact Top Layout (Generous Tab Hit-Target) */
    .main .block-container,
    [data-testid="stMainBlockContainer"],
    section.main > div:has(.block-container) {
        padding-top: 1.5rem !important;
        padding-bottom: 2rem !important;
    }
    header[data-testid="stHeader"],
    [data-testid="stHeader"],
    .stAppToolbar {
        pointer-events: none !important;
        background: transparent !important;
    }
    header button,
    header [role="button"],
    [data-testid="stToolbarActions"],
    [data-testid="stMainMenu"] {
        pointer-events: auto !important;
    }
    
    /* 1. Hide Deploy Button */
    [data-testid="stAppDeployButton"],
    .stAppDeployButton,
    [data-testid="stDeployButton"] {
        display: none !important;
        visibility: hidden !important;
    }
    
    /* 2. Hide Version Copy Button */
    .stMenuVersionCopyButton {
        display: none !important;
        visibility: hidden !important;
    }
    
    /* 3. Hide Page Bottom Footer */
    footer {
        display: none !important;
        visibility: hidden !important;
    }
    
    /* 4. Tab Bar Styling (Scoped to main container to prevent leaking into sidebar) */
    section.main div[data-testid="stRadio"] > div[role="radiogroup"],
    [data-testid="stMain"] div[data-testid="stRadio"] > div[role="radiogroup"] {
        display: flex;
        flex-direction: row;
        gap: 8px;
        border-bottom: 2px solid #e2e8f0;
        padding-bottom: 0px;
        margin-bottom: 20px;
        position: relative !important;
        z-index: 999999 !important;
    }
    section.main div[data-testid="stRadio"] label,
    [data-testid="stMain"] div[data-testid="stRadio"] label {
        padding: 8px 18px;
        border-radius: 6px 6px 0 0;
        font-size: 1.05rem !important;
        font-weight: 600 !important;
        cursor: pointer;
        background: transparent;
        border-bottom: 3px solid transparent;
        margin-bottom: -2px;
        transition: all 0.15s ease-in-out;
    }
    section.main div[data-testid="stRadio"] label > div:first-child,
    [data-testid="stMain"] div[data-testid="stRadio"] label > div:first-child {
        display: none !important;
    }
    section.main div[data-testid="stRadio"] label p,
    [data-testid="stMain"] div[data-testid="stRadio"] label p {
        color: #64748b !important;
        font-size: 1.05rem !important;
        font-weight: 600 !important;
        margin: 0 !important;
    }
    section.main div[data-testid="stRadio"] label:hover,
    [data-testid="stMain"] div[data-testid="stRadio"] label:hover {
        background: rgba(37, 99, 235, 0.04);
    }
    section.main div[data-testid="stRadio"] label:has(input:checked),
    [data-testid="stMain"] div[data-testid="stRadio"] label:has(input:checked) {
        border-bottom: 3px solid #2563eb !important;
        background: rgba(37, 99, 235, 0.05);
    }
    section.main div[data-testid="stRadio"] label:has(input:checked) p,
    [data-testid="stMain"] div[data-testid="stRadio"] label:has(input:checked) p {
        color: #2563eb !important;
        font-weight: 800 !important;
    }

    /* 5. Cloud Header Action Buttons Alignment */
    div[data-testid="stHorizontalBlock"]:has(.cloud-channel-info-box) {
        align-items: stretch !important;
    }
    div[data-testid="stHorizontalBlock"]:has(.cloud-channel-info-box) div[data-testid="stColumn"]:last-child div[data-testid="stVerticalBlock"] {
        gap: 6px !important;
    }
    div[data-testid="stHorizontalBlock"]:has(.cloud-channel-info-box) div[data-testid="stColumn"]:last-child div[data-testid="stElementContainer"],
    div[data-testid="stHorizontalBlock"]:has(.cloud-channel-info-box) div[data-testid="stColumn"]:last-child .stButton,
    div[data-testid="stHorizontalBlock"]:has(.cloud-channel-info-box) div[data-testid="stColumn"]:last-child .stLinkButton {
        margin: 0 !important;
        padding: 0 !important;
        height: 34px !important;
    }
    div[data-testid="stHorizontalBlock"]:has(.cloud-channel-info-box) div[data-testid="stColumn"]:last-child button,
    div[data-testid="stHorizontalBlock"]:has(.cloud-channel-info-box) div[data-testid="stColumn"]:last-child a {
        min-height: 34px !important;
        height: 34px !important;
        padding: 0 12px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        font-size: 0.86rem !important;
        margin: 0 !important;
        box-sizing: border-box !important;
        border-radius: 6px !important;
    }
    div[data-testid="stHorizontalBlock"]:has(.cloud-channel-info-box) div[data-testid="stColumn"]:last-child a {
        background-color: #f0f9ff !important;
        border: 1px solid #bae6fd !important;
        color: #0369a1 !important;
        font-weight: 600 !important;
        text-decoration: none !important;
    }
    div[data-testid="stHorizontalBlock"]:has(.cloud-channel-info-box) div[data-testid="stColumn"]:last-child a:hover {
        background-color: #e0f2fe !important;
        border-color: #0284c7 !important;
        color: #075985 !important;
    }
</style>
""", unsafe_allow_html=True)

# State initialization (콘티 본문 영구 보존 및 동적 버전 키 체계)
if "conti_version" not in st.session_state:
    st.session_state.conti_version = 0

if "pending_conti_update" in st.session_state:
    st.session_state.conti_text = st.session_state.pending_conti_update
    st.session_state.conti_version = st.session_state.get("conti_version", 0) + 1
    del st.session_state.pending_conti_update

if "active_preset_name" not in st.session_state:
    initial_presets = get_preset_ppts()
    st.session_state.active_preset_name = initial_presets[0] if initial_presets else None

if "conti_text" not in st.session_state or not st.session_state.conti_text:
    st.session_state.conti_text = get_default_conti_template(st.session_state.active_preset_name)
elif "가사 1행을 입력하세요" in st.session_state.conti_text:
    new_skel = get_default_conti_template(st.session_state.active_preset_name)
    st.session_state.conti_text = new_skel
    st.session_state.conti_version = st.session_state.get("conti_version", 0) + 1


if "plan_text" not in st.session_state:
    st.session_state.plan_text = ""
if "plan_editor" not in st.session_state:
    st.session_state.plan_editor = ""
if "selected_tab" not in st.session_state:
    st.session_state.selected_tab = "1. 콘티 작성"

# Cloud Sync Lifecycle: 최초 실행 시 자동 동기화 및 5초 주기 실시간 백그라운드 감시 데몬
if not hasattr(st, "_telegram_watchdog_started"):
    st._telegram_watchdog_started = True
    import threading
    def _telegram_watchdog_loop():
        while True:
            try:
                check_and_process_channel_updates()
            except Exception:
                pass
            time.sleep(5)
    _w_thread = threading.Thread(target=_telegram_watchdog_loop, daemon=True, name="TelegramOTAWatchdog")
    _w_thread.start()

if "last_cloud_sync_time" not in st.session_state:
    try:
        sync_res = sync_on_startup(st.session_state.active_preset_name)
        st.session_state.cloud_connected = sync_res.get("connected", False)
        st.session_state.last_sync_result = sync_res
    except Exception as e:
        st.session_state.cloud_connected = False
        st.session_state.last_sync_result = {"error": str(e)}
    st.session_state.last_cloud_sync_time = time.time()
else:
    if time.time() - st.session_state.last_cloud_sync_time > 1800:
        st.session_state.last_cloud_sync_time = time.time()

cloud_badge = "🟢 클라우드 연결됨" if st.session_state.get("cloud_connected", True) else "⚪ 오프라인 모드"

# ==========================================
# 관리자 모드(Admin Console) 전용 뷰 렌더링
# ==========================================
if st.session_state.get("app_mode") == "admin":
    st.markdown("""
    <style>
        [data-testid="stSidebar"], [data-testid="collapsedControl"] {
            display: none !important;
        }
        .main .block-container {
            max-width: 1200px !important;
            padding-top: 1.5rem !important;
        }
    </style>
    """, unsafe_allow_html=True)
    render_admin_console()
    st.stop()

# Sidebar: Global Settings & File Handlers
with st.sidebar:
    st.header("📁 표준 양식 설정")
    st.caption("텍스트 위치, 배율, 폰트 임베딩의 기준이 되는 원본 PPT 서식입니다.")
    
    template_mode = st.radio(
        "표준 양식 지정 방식",
        ["교회 표준 양식 선택", "내 PC에서 직접 업로드"],
        index=0
    )
    
    active_template_source = None
    active_template_name = None
    preset_info_dict = get_preset_info_dict()
    
    if template_mode == "교회 표준 양식 선택":
        preset_names = list(preset_info_dict.keys())
        if preset_names:
            default_preset_idx = 0
            for p_idx, p_name in enumerate(preset_names):
                if "청년부" in p_name:
                    default_preset_idx = p_idx
                    break
            selected_preset = st.radio(
                "표준 양식 선택",
                preset_names,
                index=default_preset_idx,
                help="프로젝트에 등록된 교회 표준 찬양 PPT 양식 목록입니다."
            )
            active_template_source = preset_info_dict[selected_preset]["template_path"]
            active_template_name = selected_preset
            st.caption(f"📌 적용 중: **{selected_preset}**")
            
            # 표준 양식 변경 시 처리 (사용자 작성 초안 보호)
            if st.session_state.get("active_preset_name") != selected_preset:
                old_preset = st.session_state.get("active_preset_name")
                st.session_state.active_preset_name = selected_preset
                cur_text = st.session_state.get("conti_text", "").strip()
                old_template = get_default_conti_template(old_preset).strip() if old_preset else ""
                # 기존 내용이 비어있거나 이전 표준 양식의 기본 서식 그대로인 경우에만 새 표준 양식 기본 서식으로 교체
                if not cur_text or cur_text == old_template:
                    st.session_state.conti_text = get_default_conti_template(selected_preset)
                    st.session_state.conti_version = st.session_state.get("conti_version", 0) + 1
                st.rerun()

        else:
            st.info("💡 아직 등록된 공식 표준 양식이 없습니다.\n\n텔레그램 채널에 일반 텍스트로 `#추가 : [양식 이름]`을 전송하여 표준 양식을 생성해 보세요.\n\n(또는 '내 PC에서 직접 업로드'를 이용하세요)")

    else:
        uploaded_template = st.file_uploader(
            "기준 PPT 파일 (.pptx)",
            type=["pptx"],
            help="사용할 기준 PPTX 파일을 업로드하세요. 해당 파일의 텍스트 박스 위치, 여백, 폰트 임베딩을 그대로 복제합니다."
        )
        if uploaded_template is not None:
            active_template_source = uploaded_template
            active_template_name = uploaded_template.name
            st.caption(f"📤 적용 중: **{uploaded_template.name}**")
        else:
            st.warning("⚠️ 기준이 될 .pptx 파일을 업로드해야 PPT 생성이 가능합니다.")

    st.divider()
    st.header("🖼️ 표지 이미지 설정 (선택)")
    st.caption("1번 슬라이드에 16:9 전체 화면으로 채워질 표지 이미지입니다.")
    
    active_cover_source = None
    active_cover_name = "미적용 (기본 텍스트 표지)"
    
    # 분기 1: 교회 표준 양식 모드
    if template_mode == "교회 표준 양식 선택" and active_template_name:
        preset_info = preset_info_dict.get(active_template_name, {})
        detected_cover_path = preset_info.get("cover_path")
        detected_cover_name = preset_info.get("cover_name")
        
        if detected_cover_path:
            use_preset_cover = st.checkbox(
                f"표준 양식 기본 표지 사용 ({detected_cover_name})",
                value=True,
                help="체크 해제 시 표지 이미지 없이 텍스트 표지 슬라이드로 생성됩니다."
            )
            
            if use_preset_cover:
                active_cover_source = detected_cover_path
                active_cover_name = detected_cover_name
                
                c_thumb, c_pop = st.columns([6, 4])
                with c_thumb:
                    st.image(detected_cover_path, use_container_width=True)
                with c_pop:
                    with st.popover("➕ 크게 보기", use_container_width=True):
                        st.caption(f"🖼️ {detected_cover_name} (고해상도 원본)")
                        st.image(detected_cover_path, use_container_width=True)
                        
                st.caption(f"📌 기본 표지 적용 중: `{detected_cover_name}`")
                
            custom_cover = st.file_uploader(
                "다른 이미지로 이번 주만 임시 교체 (.png, .jpg)",
                type=["png", "jpg", "jpeg"],
                help="업로드 시 표준 양식의 기본 표지 대신 이 이미지가 우선 적용됩니다."
            )

            if custom_cover is not None:
                active_cover_source = custom_cover
                active_cover_name = custom_cover.name
                c_uthumb, c_upop = st.columns([6, 4])
                with c_uthumb:
                    st.image(custom_cover, use_container_width=True)
                with c_upop:
                    with st.popover("➕ 크게 보기", use_container_width=True):
                        st.caption(f"🖼️ {custom_cover.name} (고해상도 원본)")
                        st.image(custom_cover, use_container_width=True)
                st.caption(f"🌟 임시 교체 적용 중: `{custom_cover.name}`")
        else:
            uploaded_cover = st.file_uploader(
                "표지 이미지 파일 업로드 (.png, .jpg)",
                type=["png", "jpg", "jpeg"],
                help="업로드 시 1번 슬라이드에 전체 화면으로 삽입됩니다."
            )
            if uploaded_cover is not None:
                active_cover_source = uploaded_cover
                active_cover_name = uploaded_cover.name
                c_thumb, c_pop = st.columns([6, 4])
                with c_thumb:
                    st.image(uploaded_cover, use_container_width=True)
                with c_pop:
                    with st.popover("➕ 크게 보기", use_container_width=True):
                        st.caption(f"🖼️ {uploaded_cover.name} (고해상도 원본)")
                        st.image(uploaded_cover, use_container_width=True)
                st.caption(f"📤 적용 중: `{uploaded_cover.name}`")
            else:
                st.caption("ℹ️ 표지 이미지가 없으면 기본 텍스트 표지 슬라이드로 생성됩니다.")

    # 분기 2: 내 PC에서 직접 업로드 모드
    else:
        extracted_cover = None
        if active_template_source is not None:
            extracted_cover = extract_cover_image_from_template(active_template_source)
            
        if extracted_cover is not None:
            ext_blob_io, ext_filename = extracted_cover
            use_ext_cover = st.checkbox(
                "기준 PPT 1번 슬라이드의 표지 이미지 유지",
                value=True,
                help="업로드한 기준 PPT 첫 페이지에 포함된 그림을 결과물 1번 슬라이드 표지로 그대로 사용합니다."
            )
            if use_ext_cover:
                active_cover_source = ext_blob_io
                active_cover_name = f"[기준 PPT 1번 슬라이드 표지]"
                c_thumb, c_pop = st.columns([6, 4])
                with c_thumb:
                    st.image(ext_blob_io, use_container_width=True)
                with c_pop:
                    with st.popover("➕ 크게 보기", use_container_width=True):
                        st.caption("🖼️ 기준 PPT 1번 슬라이드에서 자동 추출된 표지 (고해상도 원본)")
                        st.image(ext_blob_io, use_container_width=True)
                st.caption("📌 기준 PPT 1번 슬라이드 표지 적용 중")
                
            custom_override = st.file_uploader(
                "다른 이미지로 교체 업로드 (.png, .jpg)",
                type=["png", "jpg", "jpeg"],
                help="업로드 시 기준 PPT의 표지 이미지 대신 이 이미지가 우선 적용됩니다."
            )
            if custom_override is not None:
                active_cover_source = custom_override
                active_cover_name = custom_override.name
                c_uthumb, c_upop = st.columns([6, 4])
                with c_uthumb:
                    st.image(custom_override, use_container_width=True)
                with c_upop:
                    with st.popover("➕ 크게 보기", use_container_width=True):
                        st.caption(f"🖼️ {custom_override.name} (고해상도 원본)")
                        st.image(custom_override, use_container_width=True)
                st.caption(f"🌟 교체 적용 중: `{custom_override.name}`")
        else:
            uploaded_cover = st.file_uploader(
                "표지 이미지 파일 업로드 (.png, .jpg)",
                type=["png", "jpg", "jpeg"],
                help="업로드 시 1번 슬라이드에 전체 화면으로 삽입됩니다."
            )
            if uploaded_cover is not None:
                active_cover_source = uploaded_cover
                active_cover_name = uploaded_cover.name
                c_thumb, c_pop = st.columns([6, 4])
                with c_thumb:
                    st.image(uploaded_cover, use_container_width=True)
                with c_pop:
                    with st.popover("➕ 크게 보기", use_container_width=True):
                        st.caption(f"🖼️ {uploaded_cover.name} (고해상도 원본)")
                        st.image(uploaded_cover, use_container_width=True)
                st.caption(f"📤 적용 중: `{uploaded_cover.name}`")
            else:
                st.caption("ℹ️ 표지 이미지가 없으면 기본 텍스트 표지 슬라이드로 생성됩니다.")

    st.divider()
    st.header("📁 콘티 파일 업로드")
    uploaded_file = st.file_uploader(
        "작성한 콘티 파일 (.txt)",
        type=["txt"],
        help="콘티 텍스트 파일을 업로드하면 에디터에 내용이 자동으로 입력됩니다."
    )
    if uploaded_file is not None:
        file_key = f"{uploaded_file.name}_{uploaded_file.size}"
        if st.session_state.get("last_uploaded_file_key") != file_key:
            file_content = uploaded_file.read().decode("utf-8")
            st.session_state.last_uploaded_file_key = file_key
            st.session_state.pending_conti_update = file_content
            st.rerun()
    else:
        st.session_state.last_uploaded_file_key = None

    st.divider()
    st.header("🎵 등록된 찬양곡 검색")
    st.caption("자주 부르는 찬양을 검색하여 기존 곡 자리에 바꾸거나 끝에 추가할 수 있습니다.")
    lib_search = st.text_input("찬양 검색", placeholder="예: 꽃들도, 새 힘...", label_visibility="collapsed")
    matched_songs = search_library_with_version_limit(lib_search, max_versions=3)
    if matched_songs:
        st.caption(f"검색 결과 {len(matched_songs)}곡 (곡별 최신 3개 버전 우선 표시)")
        current_c = st.session_state.get("conti_text", "").strip()
        parsed_c = parse_user_conti(current_c)

        cur_songs = parsed_c.get("songs", [])
        
        for s_title, versions in matched_songs.items():
            with st.expander(f"🎵 {s_title} ({len(versions)}개 버전)"):
                for v in versions:
                    st.markdown(f"**날짜**: `{v['date']}`")
                    if v['routine']:
                        st.caption(f"루틴: {v['routine']}")
                    
                    # 기존 곡 자리에 바꾸기 버튼들
                    if cur_songs:
                        st.caption("🔄 기존 곡 바꾸기:")
                        cols_rep = st.columns(min(len(cur_songs), 3))
                        for idx, s in enumerate(cur_songs):
                            target_num = idx + 1
                            col_t = cols_rep[idx % len(cols_rep)]
                            if col_t.button(f"🔄 {target_num}번 곡 바꾸기", key=f"btn_rep_side_{s_title}_{v['date']}_{target_num}", help=f"{target_num}번 곡({s.get('title')}) 자리에 바꾸기", use_container_width=True):
                                new_c = replace_song_in_conti(current_c, target_num, v['data'])
                                st.session_state.pending_conti_update = new_c
                                st.rerun()
                                
                    next_num = len(cur_songs) + 1
                    if st.button(f"➕ 순서 끝에 추가 ({next_num}번)", key=f"btn_add_lib_{s_title}_{v['date']}", use_container_width=True):
                        new_c = append_song_to_conti(current_c, v['data'])
                        st.session_state.pending_conti_update = new_c
                        st.rerun()
    else:
        st.caption("일치하는 찬양곡이 없습니다.")

    # 사이드바 최하단 시스템 정보 (은은한 푸터)
    st.divider()
    st.markdown(f"""
    <div style="text-align: center; color: #94a3b8; font-size: 0.8rem; padding: 2px 0 10px 0; line-height: 1.6;">
        <div style="font-weight: 600; color: #64748b;">🎵 LOGOS 찬양 스튜디오</div>
        <div style="margin-top: 2px;">{cloud_badge} · v1.5.0</div>
    </div>
    """, unsafe_allow_html=True)

    if st.session_state.get("is_admin", False):
        st.markdown("<div style='height: 4px;'></div>", unsafe_allow_html=True)
        if st.button("🛡️ 관리자 콘솔로 전환", key="btn_switch_to_admin", use_container_width=True):
            st.session_state.app_mode = "admin"
            st.rerun()

# Stateful Tabs (100% Cross-origin Cloud Safe)

tab_options = ["1. 콘티 작성", "2. 슬라이드 편집", "3. 자료실", "4. 사용 설명서"]

if "switch_to_tab" in st.session_state:
    st.session_state.selected_tab = st.session_state.switch_to_tab
    del st.session_state.switch_to_tab

if "selected_tab" not in st.session_state or st.session_state.selected_tab not in tab_options:
    st.session_state.selected_tab = "1. 콘티 작성"

active_tab = st.radio(
    "메뉴 탭 선택",
    tab_options,
    key="selected_tab",
    horizontal=True,
    label_visibility="collapsed"
)

# ==========================================
# TAB 1: CONTI WRITING
# ==========================================
if active_tab == "1. 콘티 작성":
    parsed_preview = parse_user_conti(st.session_state.get("conti_text", ""))
    
    col_edit, col_guide = st.columns([7.2, 4.8])
    
    with col_edit:
        conti_user_input = render_conti_editor(
            value=st.session_state.get("conti_text", ""),
            height=540,
            key=f"conti_editor_v{st.session_state.get('conti_version', 0)}"
        )
        if conti_user_input is not None and conti_user_input != st.session_state.get("conti_text", ""):
            st.session_state.conti_text = conti_user_input

        # 슬라이드 생성 버튼 동적 레이블 생성
        if template_mode == "교회 표준 양식 선택":
            if active_template_name:
                slide_btn_label = f"슬라이드 생성 ➡️ [적용 양식 : {active_template_name}]"
            else:
                slide_btn_label = "슬라이드 생성 ➡️ [표준 양식 미설정]"
        else:
            if active_template_source is not None:
                slide_btn_label = "슬라이드 생성 ➡️ [적용 양식 : 사용자 지정]"
            else:
                slide_btn_label = "슬라이드 생성 ➡️ [기준 PPT 미업로드]"
            
        if st.button(slide_btn_label, type="primary", use_container_width=True):
            if not st.session_state.conti_text.strip():
                st.warning("⚠️ 콘티 내용을 입력하거나 .txt 파일을 업로드한 후 버튼을 눌러주세요!")
            else:
                parsed = parse_user_conti(st.session_state.conti_text)
                new_plan = generate_plan_text_from_conti(parsed)
                st.session_state.plan_text = new_plan
                st.session_state.plan_editor = new_plan
                st.session_state.switch_to_tab = "2. 슬라이드 편집"
                st.rerun()
            
    with col_guide:
        with st.expander("📌 콘티 작성 규칙 안내", expanded=True):
            st.markdown("""
            고정 정보(#), 곡 제목(##), 루틴 문자열, 파트([V], [C] 등)를 표준 서식에 맞춰 작성하세요. 슬라이드는 **빈 줄(더블 엔터)**로 자동 분할됩니다.
            
            - **표지 정보**: `# 2026.09.13 LOGOS 청년 모임`
            - **대표기도자**: `기도: ㅇㅇㅇ 청년`
            - **찬양곡 시작**: `## 1. 새 힘 얻으리`
            - **루틴(송폼)**: `루틴: Intro(8) - V - C - V - C...` (악보 기준)
            - **파트 선언**: `[V]`, `[V1]`, `[V2]`, `[P]`, `[C]`, `[B]` 등
            - **슬라이드 분할**: **빈 줄 하나(더블 엔터)**를 넣으면 다음 슬라이드로 넘어갑니다. (1슬라이드 당 1~2줄 권장)
            """)

        st.divider()

        st.markdown("##### 💡 작성 도움말")
        lint_res = validate_conti_text(parsed_preview)
        if lint_res["warnings"]:
            for w in lint_res["warnings"]:
                st.warning(f"⚠️ {w}")
        else:
            st.success("✅ 콘티 서식이 올바릅니다.")

# ==========================================
# TAB 2: SLIDE EDITING
# ==========================================
elif active_tab == "2. 슬라이드 편집":
    # 상단 슬라이드 장수 요약 자리 예약 (실시간 동기화)
    metric_placeholder = st.empty()
    
    # 에디터로부터 최신 텍스트 수신
    plan_user_input = render_slide_ide_editor(
        value=st.session_state.get("plan_text", ""),
        height=580,
        key="plan_ide_editor"
    )
    if plan_user_input is not None and plan_user_input != st.session_state.get("plan_text", ""):
        st.session_state.plan_text = plan_user_input
        st.session_state.plan_editor = plan_user_input
        
    # 최신 텍스트 기준으로 슬라이드 및 장수 즉시 계산
    current_plan = st.session_state.get("plan_text", "")
    current_slides = parse_plan_text_to_slides(current_plan)
    total_slide_count = len(current_slides)
    
    # 상단에 슬라이드 장수 미니멀 표시
    with metric_placeholder.container():
        st.markdown(
            f"<div style='font-size: 0.95rem; font-weight: 600; color: #475569; padding: 2px 0 4px 0;'>"
            f"슬라이드 {total_slide_count}장</div>",
            unsafe_allow_html=True
        )
    
    is_template_ready = (active_template_source is not None)
    
    if not is_template_ready:
        st.warning("⚠️ **[필수] 표준 양식이 설정되지 않았습니다.**\n\n좌측 사이드바의 **「📁 표준 양식 설정」**에서 [교회 표준 양식 선택]을 지정하거나 [내 PC에서 직접 업로드]를 완료해야 PPT를 생성할 수 있습니다.")
        
    if st.button("PPT 파일 생성", type="primary", use_container_width=True, disabled=not is_template_ready):
        if not current_slides:
            st.warning("⚠️ 슬라이드 내용이 비어있습니다. 1. 콘티 작성 탭에서 콘티를 작성 후 생성해 주세요.")
        else:
            with st.spinner("표준 양식 서식 및 폰트를 복제하여 PPT 파일 생성 중..."):
                prs = build_praise_pptx(current_slides, template_source=active_template_source, cover_image=active_cover_source)
                
                parsed = parse_user_conti(st.session_state.get("conti_text", ""))
                raw_date = parsed.get("date_str", "2026.09.06")
                clean_date = raw_date.replace(".", "").replace(" ", "").replace("-", "") if raw_date != "미입력" else "20260101"
                if template_mode == "내 PC에서 직접 업로드":
                    dept_label = "기타"
                else:
                    dept_label = normalize_department_name(active_template_name or "청년부")
                
                # 1. 완성본 PPTX를 데이터 보관함/2_찬양_PPT/ 에 표준 파일명으로 자동 보관
                saved_pptx_path, standard_filename, pptx_bytes = save_pptx_to_archive(
                    prs, 
                    preset_name=dept_label, 
                    title=f"{raw_date} 찬양 가사", 
                    date_str=clean_date
                )
                
                # 2. 콘티 내 찬양곡들을 낱개로 분해하여 데이터 보관함/3_찬양곡_라이브러리/ 에 표준 파일명으로 적재
                saved_song_files = decompose_and_save_songs_from_conti(parsed, department=dept_label)
                
                # 3. 작성 콘티 초안을 데이터 보관함/4_찬양_콘티/ 에 로컬 자동 보관
                try:
                    os.makedirs(DIR_CONTI, exist_ok=True)
                    conti_content = st.session_state.get("conti_text", "")
                    with open(os.path.join(DIR_CONTI, f"{clean_date}_최근작업초안.txt"), "w", encoding="utf-8") as f:
                        f.write(conti_content)
                except Exception:
                    pass
                
                # 4. 텔레그램 클라우드 공식 채널로 완성 PPTX 및 분해 찬양곡 일괄 자동 동기화
                cloud_status_msg = ""
                try:
                    with st.spinner("☁️ 텔레그램 공식 채널로 클라우드 백업 동기화 중..."):
                        cloud_res = sync_upload_created_assets(
                            saved_pptx_path, 
                            saved_song_files, 
                            department=dept_label, 
                            date_str=clean_date
                        )
                        if cloud_res.get("success"):
                            cloud_status_msg = f"\n\n☁️ **텔레그램 클라우드 공식 채널 백업 완료!** ({cloud_res.get('uploaded_count')}개 자산 동기화됨)"
                        else:
                            cloud_status_msg = f"\n\n⚠️ 로컬 저장은 완료되었으나 클라우드 전송 중 일부 지연: {', '.join(cloud_res.get('errors', []))}"
                except Exception as e:
                    cloud_status_msg = f"\n\n⚪ (오프라인 모드: 로컬 '데이터 보관함'에 안전하게 보관되었습니다)"
                
                st.download_button(
                    label=f"💾 PPT 파일 다운로드 ({standard_filename})",
                    data=pptx_bytes,
                    file_name=standard_filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    type="primary",
                    use_container_width=True
                )
                st.success(f"🎉 **{len(current_slides)}장의 찬양 PPT 생성 완료!**\n\n📁 `데이터 보관함/2_찬양_PPT/{standard_filename}` 에 자동 보관되었으며, {len(saved_song_files)}개 찬양곡이 라이브러리에 개별 분해 적재되었습니다.{cloud_status_msg}")

# ==========================================
# TAB 3: MATERIALS ARCHIVE
# ==========================================
elif active_tab == "3. 자료실":
    # 현재 적용된 부서 / 프리셋 확인
    if template_mode == "내 PC에서 직접 업로드":
        current_dept = "기타"
    else:
        current_dept = normalize_department_name(active_template_name)

    # 탭 3 내부 서브 메뉴 (필터)
    col_f1, col_f2 = st.columns([7, 5])
    with col_f1:
        archive_view = st.radio(
            "자료 구분",
            [f"📊 [{current_dept}] 찬양 PPT", "🌐 전체 찬양 PPT", "🎵 찬양곡 자료실"],
            horizontal=True,
            label_visibility="collapsed"
        )
    with col_f2:
        search_kw = st.text_input("자료 검색", placeholder="제목 또는 날짜 검색 (예: 2026.09, 꽃들도)", label_visibility="collapsed")

    catalog = load_local_catalog()
    all_items = catalog.get("items", [])

    # 1. 찬양 PPT 뷰
    if "찬양 PPT" in archive_view:
        is_dept_filtered = f"[{current_dept}]" in archive_view
        ppt_items = [it for it in all_items if it.get("type") == "PPT" and not is_pptx_trashed(it.get("filename"))]
        
        # 로컬 폴더(DIR_PPTX)에만 존재하는 파일도 카탈로그에 없는 경우 목록에 포함
        catalog_filenames = {it.get("filename") for it in ppt_items}
        if os.path.exists(DIR_PPTX):
            for f in os.listdir(DIR_PPTX):
                if f.endswith(".pptx") and f not in catalog_filenames and not is_pptx_trashed(f):
                    meta = parse_standard_filename(f)
                    ppt_items.append({
                        "filename": f,
                        "file_id": None,
                        "file_size": os.path.getsize(os.path.join(DIR_PPTX, f)),
                        "date": meta["date"] if meta else "",
                        "type": "PPT",
                        "preset": meta["preset"] if meta else "기타",
                        "title": meta["title"] if meta else f,
                        "ext": "pptx"
                    })

        if is_dept_filtered:
            ppt_items = [
                it for it in ppt_items 
                if normalize_department_name(it.get("preset")) == current_dept
            ]

        if search_kw.strip():
            kw = search_kw.strip().lower()
            ppt_items = [it for it in ppt_items if kw in it.get("title", "").lower() or kw in it.get("date", "").lower() or kw in it.get("filename", "").lower()]

        # 보유 현황 통계 산출
        local_count = sum(1 for it in ppt_items if os.path.exists(os.path.join(DIR_PPTX, it.get("filename", ""))))
        cloud_only_count = len(ppt_items) - local_count
        st.caption(f"총 {len(ppt_items)}개의 PPT 등록됨 (🟢 내려받기 가능: {local_count}개 | ☁️ 클라우드 보관 중: {cloud_only_count}개)")

        if not ppt_items:
            st.info(f"ℹ️ 등록된 [{current_dept if is_dept_filtered else '전체'}] 찬양 PPT가 없습니다.")
        else:
            for it in ppt_items:
                fname = it.get("filename")
                ftitle = it.get("title", fname)
                fdate = it.get("date", "")
                fsize = it.get("file_size", 0)
                fsize_str = f"{fsize / (1024*1024):.1f} MB" if fsize >= 1024*1024 else f"{fsize / 1024:.0f} KB"
                dept_badge = normalize_department_name(it.get("preset", "기타"))
                local_fpath = os.path.join(DIR_PPTX, fname)
                is_local = os.path.exists(local_fpath)

                status_badge = (
                    "<span style='background:#dcfce7; color:#15803d; border:1px solid #86efac; padding:2px 8px; border-radius:12px; font-size:0.8rem; font-weight:700;'>🟢 내려받기 가능</span>"
                    if is_local else
                    "<span style='background:#f1f5f9; color:#475569; border:1px solid #cbd5e1; padding:2px 8px; border-radius:12px; font-size:0.8rem; font-weight:600;'>☁️ 클라우드 보관 중</span>"
                )

                with st.container():
                    c_card_info, c_card_act = st.columns([7, 5])
                    with c_card_info:
                        st.markdown(f"**📊 {ftitle}** &nbsp; <span style='background:#e0e7ff; color:#3730a3; padding:2px 8px; border-radius:12px; font-size:0.8rem;'>{dept_badge}</span> &nbsp; {status_badge} &nbsp; <span style='color:#64748b; font-size:0.85rem;'>일자: {fdate} | {fsize_str}</span>", unsafe_allow_html=True)
                        st.caption(f"파일명: `{fname}`")
                    with c_card_act:
                        if is_local:
                            if os.path.exists(local_fpath):
                                with open(local_fpath, "rb") as f_dl:
                                    st.download_button(
                                        "📥 다운로드",
                                        data=f_dl.read(),
                                        file_name=fname,
                                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                        key=f"dl_local_{fname}",
                                        use_container_width=True
                                    )
                        else:
                            file_id = it.get("file_id")
                            if file_id:
                                if st.button("📥 다운로드", key=f"dl_cloud_{fname}", use_container_width=True, type="primary"):
                                    with st.spinner("클라우드에서 다운로드 중..."):
                                        dl_ok, dl_err = download_file_by_id(file_id, local_fpath)
                                        if dl_ok:
                                            st.toast(f"✅ 다운로드 완료: {fname}")
                                            st.rerun()
                                        else:
                                            st.error(f"다운로드 실패: {dl_err}")
                            else:
                                st.caption("클라우드 파일 ID 없음")
                    st.divider()

    # 2. 찬양곡 자료실 뷰
    else:
        song_items = [it for it in all_items if it.get("type") == "찬양곡" and not is_song_trashed(it.get("title"), it.get("date"))]
        if search_kw.strip():
            kw = search_kw.strip().lower()
            song_items = [it for it in song_items if kw in it.get("title", "").lower() or kw in it.get("filename", "").lower()]

        local_song_count = sum(1 for it in song_items if os.path.exists(os.path.join(DIR_SONGS, sanitize_filename_part(it.get("title", "")), it.get("filename", ""))) and not is_song_trashed(it.get("title"), it.get("date")))
        cloud_only_song_count = len(song_items) - local_song_count
        st.caption(f"총 {len(song_items)}개의 찬양곡 보관됨 (사이드바 '🎵 등록된 찬양곡 검색'을 통해 콘티에 바로 불러올 수 있습니다)")

        if not song_items:
            st.info("ℹ️ 등록된 찬양곡이 없습니다.")
        else:
            for it in song_items:
                fname = it.get("filename")
                stitle = it.get("title", fname)
                sdate = it.get("date", "")
                dept_badge = normalize_department_name(it.get("preset", "청년부"))
                local_dir = os.path.join(DIR_SONGS, sanitize_filename_part(stitle))
                local_spath = os.path.join(local_dir, fname)
                is_local = os.path.exists(local_spath)

                status_badge = (
                    "<span style='background:#dcfce7; color:#15803d; border:1px solid #86efac; padding:2px 8px; border-radius:12px; font-size:0.8rem; font-weight:700;'>🟢 보유 중</span>"
                    if is_local else
                    "<span style='background:#f1f5f9; color:#475569; border:1px solid #cbd5e1; padding:2px 8px; border-radius:12px; font-size:0.8rem; font-weight:600;'>☁️ 클라우드 보관</span>"
                )

                with st.container():
                    st.markdown(f"**🎵 {stitle}** &nbsp; <span style='background:#fef3c7; color:#92400e; padding:2px 8px; border-radius:12px; font-size:0.8rem;'>{dept_badge}</span> &nbsp; {status_badge} &nbsp; <span style='color:#64748b; font-size:0.85rem;'>버전일자: {sdate}</span>", unsafe_allow_html=True)
                    st.caption(f"파일명: `{fname}`")
                    st.divider()

# ==========================================
# TAB 4: USER MANUAL
# ==========================================
elif active_tab == "4. 사용 설명서":
    st.markdown("### 📖 LOGOS 찬양 PPT 제작 스튜디오 사용 설명서")
    
    st.info("💡 **접속 비밀번호**: 초기 설정 비밀번호는 **`0000`** 입니다.")
    
    st.markdown("""
---

#### 0. 사이드바 사용
프로그램의 기본 설정과 편의 기능을 제공합니다.

- **0-1. 표준 양식 설정**
  - 본 시스템은 기준이 되는 `.pptx` 파일의 서식과 폰트를 복제하여 새로운 찬양 PPT를 생성합니다.
  - **교회 표준 양식**: 사전 등록된 기준 PPT 파일과 각 부서별 표지 이미지로 구성되어 있으며, 양식 선택 시 해당 세팅으로 자동 적용됩니다. (기본값: 청년부)
  - **내 PC에서 직접 업로드**: PC에 보관 중인 기준 PPT 파일과 표지 이미지를 직접 업로드하여 일회성으로 사용할 수 있습니다.

- **0-2. 등록된 찬양곡 검색**
  - 사이드바 하단에서 시스템에 보관된 찬양곡을 검색할 수 있습니다.
  - 찬양곡은 콘티에서 추출되며 작성 일자에 따라서 같은 곡이라도 버전별로 보관됩니다.
  - 검색 결과에서 `[➕ 순서 끝에 추가]` 또는 `[🔄 n번 곡 바꾸기]`를 누르면 가사와 루틴이 에디터에 자동으로 채워집니다.

---

#### 1. 콘티 작성
찬양팀 악보를 바탕으로 사용자가 직접 작성하는 단계입니다.

- **1-1. 작성 기본**
  - 콘티는 PPT의 슬라이드에 들어가게 될 내용을 작성하는 단계로 찬양 제목, 루틴(송폼), 파트별 가사 등으로 작성됩니다.
  - 콘티의 양식은 우측 **[💡 작성 도움말]**의 안내를 참고하세요. 기본 양식이 자동으로 제공됩니다.
  - 찬양팀 악보에서 루틴과 각 파트의 가사를 적으신 후, 하나의 슬라이드에 적절한 양의 가사가 들어가도록 편집합니다.
  - **슬라이드 분할**: **더블 엔터(빈 줄 하나)**를 이용하여 슬라이드를 구분합니다. 줄바꿈(엔터 1번)으로 한 슬라이드 내 가사 모양을 정돈할 수 있습니다. (1슬라이드 당 1~2줄 권장)
  - 작성 방법을 잘 모르실 경우, 사이드바 검색에서 찬양곡을 불러와 보시면 쉽게 이해하실 수 있습니다.
  - 작성이 완료되면 하단의 **[슬라이드 생성 ➡️]** 버튼을 누릅니다.
""")

    st.warning("⚠️ **콘티 작성 시 주의 사항**\n\n본 시스템은 루틴을 바탕으로 각 파트를 자동으로 배치하는 방식으로 작동합니다. 각 파트의 이름(`[V]`, `[C]` 등)이 루틴에 존재하지 않거나 일치하지 않으면 PPT 생성이 불가하므로 주의해 주시기 바랍니다.")

    st.markdown("""
---

#### 2. 슬라이드 편집
콘티로부터 자동 생성된 슬라이드를 검토하고 편집합니다.

- 콘티와 마찬가지로 **더블 엔터(빈 줄 하나)**로 각 슬라이드를 구분합니다.
- 좌측 넘버링의 숫자는 **슬라이드 번호**를 의미하며, `[BLANK]`는 가사 없는 암전 화면을 의미합니다.
- 내용 검토와 수정을 마친 후 하단의 **[PPT 파일 생성]** 버튼을 누릅니다.

---

#### 3. 자료실
완성된 찬양 PPT 및 찬양곡 가사를 확인하고 내려받습니다.

- 완성된 찬양 PPT와 찬양곡 가사는 안전하게 보관되어 이전 작업물을 보존한 상태로 데이터를 쌓아가며 작동합니다.
- 부서별 찬양 PPT, 전체 찬양 PPT, 찬양곡 자료실 필터를 통해 원하는 자료를 검색할 수 있습니다.
- **[📥 다운로드]** 버튼을 누르면 내 기기(PC, 태블릿, 모바일)로 즉시 저장됩니다.
""")

# Global Style & DOM Injector (Localization + Custom Footer)
st.html("""
<style>
    [data-testid="stAppDeployButton"],
    .stAppDeployButton,
    [data-testid="stDeployButton"] {
        display: none !important;
        visibility: hidden !important;
    }
    .stMenuVersionCopyButton {
        display: none !important;
        visibility: hidden !important;
    }
    footer {
        display: none !important;
        visibility: hidden !important;
    }
</style>
<script>
(function() {
    const pDoc = document;
    
    const updateMenuLocalization = () => {
        try {
            const copyBtns = pDoc.querySelectorAll('.stMenuVersionCopyButton');
            copyBtns.forEach(btn => {
                btn.style.display = 'none';
                const parent = btn.parentElement;
                if (parent) {
                    const textEl = parent.firstElementChild;
                    if (textEl && textEl !== btn) {
                        if (textEl.textContent !== 'Made by @loose_lab v1.5.0') {
                            textEl.textContent = 'Made by @loose_lab v1.5.0';
                            textEl.style.fontSize = '0.82rem';
                            textEl.style.color = '#808495';
                            textEl.style.userSelect = 'text';
                            textEl.style.webkitUserSelect = 'text';
                            textEl.style.cursor = 'text';
                        }
                    }
                }
            });

            const deployBtn = pDoc.querySelector('[data-testid="stAppDeployButton"], .stAppDeployButton, [data-testid="stDeployButton"]');
            if (deployBtn) {
                deployBtn.style.display = 'none';
            }

            const popover = pDoc.querySelector('div[data-baseweb="popover"], div[data-testid="stMainMenuPopover"]');
            if (popover) {
                const walker = pDoc.createTreeWalker(popover, NodeFilter.SHOW_TEXT, null, false);
                let node;
                while (node = walker.nextNode()) {
                    const val = node.nodeValue.trim();
                    if (val === 'System') node.nodeValue = '시스템';
                    else if (val === 'Light') node.nodeValue = '라이트';
                    else if (val === 'Dark') node.nodeValue = '다크';
                    else if (val.startsWith('Rerun')) node.nodeValue = node.nodeValue.replace(/Rerun/g, '다시 실행');
                    else if (val === 'Auto rerun') node.nodeValue = '자동 다시 실행';
                    else if (val.startsWith('Clear cache')) node.nodeValue = node.nodeValue.replace(/Clear cache/g, '캐시 삭제');
                    else if (val === 'Print') node.nodeValue = '인쇄';
                    else if (val === 'Record screen') node.nodeValue = '화면 녹화';
                }
            }
        } catch(e) {
            console.error('[Menu Localization Error]', e);
        }
    };

    if (!window.__worship_ppt_initialized) {
        window.__worship_ppt_initialized = true;
        const observer = new MutationObserver(() => {
            updateMenuLocalization();
        });
        observer.observe(pDoc.body, { childList: true, subtree: true });
    }
    updateMenuLocalization();
})();
</script>
""", unsafe_allow_javascript=True)

