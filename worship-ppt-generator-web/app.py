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

# ==========================================
# 0. 찬양팀 접속 인증 비밀번호 게이트 (4자리 0000 즉시 해제)
# ==========================================
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
    
    entered_pin = render_pin_gate(key="logos_pin_gate")
    if entered_pin == "0000":
        st.session_state.authenticated = True
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

# State initialization
if "pending_conti_update" in st.session_state:
    st.session_state.conti_text = st.session_state.pending_conti_update
    st.session_state.conti_editor = st.session_state.pending_conti_update
    del st.session_state.pending_conti_update

if "active_preset_name" not in st.session_state:
    initial_presets = get_preset_ppts()
    st.session_state.active_preset_name = initial_presets[0] if initial_presets else None

if "conti_text" not in st.session_state or not st.session_state.conti_text:
    st.session_state.conti_text = get_default_conti_template(st.session_state.active_preset_name)
if "conti_editor" not in st.session_state:
    st.session_state.conti_editor = st.session_state.conti_text
elif "가사 1행을 입력하세요" in st.session_state.conti_editor:
    new_skel = get_default_conti_template(st.session_state.active_preset_name)
    st.session_state.conti_text = new_skel
    st.session_state.conti_editor = new_skel
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
        try:
            check_and_process_channel_updates()
            fetch_remote_catalog()
            st.session_state.last_cloud_sync_time = time.time()
        except Exception:
            pass

cloud_badge = "🟢 텔레그램 클라우드 연동됨" if st.session_state.get("cloud_connected", True) else "⚪ 오프라인 모드"
badge_bg = "rgba(16, 185, 129, 0.25)" if st.session_state.get("cloud_connected", True) else "rgba(148, 163, 184, 0.25)"
badge_border = "#10b981" if st.session_state.get("cloud_connected", True) else "#94a3b8"

# Header
st.markdown(f"""
<div class="main-title-container">
    <div>
        <div class="main-title">🎵 LOGOS 찬양 PPT 제작 스튜디오</div>
        <div class="main-subtitle">1. 콘티 작성 ➔ 2. 기획안 검토 ➔ 3. PPT 제작 & 클라우드 보관소</div>
    </div>
    <div style="display: flex; align-items: center; gap: 10px;">
        <span style="background: {badge_bg}; border: 1px solid {badge_border}; padding: 4px 12px; border-radius: 16px; font-size: 0.85rem; font-weight: 700; color: white;">
            {cloud_badge}
        </span>
        <span style="background: rgba(255,255,255,0.2); padding: 4px 12px; border-radius: 16px; font-size: 0.85rem; font-weight: 700; letter-spacing: 0.5px;">
            v1.5.0
        </span>
    </div>
</div>
""", unsafe_allow_html=True)

# Sidebar: Global Settings & File Handlers
with st.sidebar:
    st.header("📁 기준 PPT 설정 [필수]")
    st.caption("텍스트 위치, 배율, 폰트 임베딩의 기준이 되는 원본 PPT입니다.")
    
    template_mode = st.radio(
        "기준 PPT 지정 방식",
        ["내 PC에서 직접 업로드", "교회 표준 프리셋 선택"],
        index=0
    )
    
    active_template_source = None
    active_template_name = None
    preset_info_dict = get_preset_info_dict()
    
    if template_mode == "교회 표준 프리셋 선택":
        preset_names = list(preset_info_dict.keys())
        if preset_names:
            selected_preset = st.radio(
                "표준 프리셋 선택",
                preset_names,
                index=0,
                help="프로젝트에 등록된 교회 표준 찬양 PPT 목록입니다."
            )
            active_template_source = preset_info_dict[selected_preset]["template_path"]
            active_template_name = selected_preset
            st.caption(f"📌 적용 중: **{selected_preset}**")
            
            # 프리셋 변경 시 해당 프리셋 맞춤 기본 콘티 서식 자동 삽입
            if st.session_state.get("active_preset_name") != selected_preset:
                st.session_state.active_preset_name = selected_preset
                new_template = get_default_conti_template(selected_preset)
                st.session_state.conti_text = new_template
                st.session_state.conti_editor = new_template
                st.rerun()
        else:
            st.info("💡 아직 등록된 공식 프리셋이 없습니다.\n\n텔레그램 채널에 일반 텍스트로 `#추가 : [프리셋 이름]`을 전송하여 프리셋을 생성해 보세요.\n\n(또는 '내 PC에서 직접 업로드'를 이용하세요)")
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
    
    # 분기 1: 교회 표준 프리셋 모드
    if template_mode == "교회 표준 프리셋 선택" and active_template_name:
        preset_info = preset_info_dict.get(active_template_name, {})
        detected_cover_path = preset_info.get("cover_path")
        detected_cover_name = preset_info.get("cover_name")
        
        if detected_cover_path:
            use_preset_cover = st.checkbox(
                f"프리셋 기본 표지 사용 ({detected_cover_name})",
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
                help="업로드 시 프리셋 기본 표지 대신 이 이미지가 우선 적용됩니다."
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
    st.header("📚 찬양 라이브러리 (아카이브)")
    st.caption("자주 부르는 찬양을 검색하여 기존 곡 자리에 교체하거나 끝에 추가할 수 있습니다.")
    lib_search = st.text_input("찬양 검색", placeholder="예: 꽃들도, 새 힘...", label_visibility="collapsed")
    matched_songs = search_library_with_version_limit(lib_search, max_versions=3)
    if matched_songs:
        st.caption(f"검색 결과 {len(matched_songs)}곡 (곡별 최신 3개 버전 우선 표시)")
        current_c = st.session_state.get("conti_editor", st.session_state.conti_text).strip()
        parsed_c = parse_user_conti(current_c)
        cur_songs = parsed_c.get("songs", [])
        
        for s_title, versions in matched_songs.items():
            with st.expander(f"🎵 {s_title} ({len(versions)}개 버전)"):
                for v in versions:
                    st.markdown(f"**날짜**: `{v['date']}`")
                    if v['routine']:
                        st.caption(f"루틴: {v['routine']}")
                    
                    # 기존 곡 자리에 교체 버튼들
                    if cur_songs:
                        st.caption("🔄 기존 곡 자리에 교체:")
                        cols_rep = st.columns(min(len(cur_songs), 3))
                        for idx, s in enumerate(cur_songs):
                            target_num = idx + 1
                            col_t = cols_rep[idx % len(cols_rep)]
                            if col_t.button(f"{target_num}번 교체", key=f"btn_rep_side_{s_title}_{v['date']}_{target_num}", help=f"{target_num}번 곡({s.get('title')}) 자리에 교체", use_container_width=True):
                                new_c = replace_song_in_conti(current_c, target_num, v['data'])
                                st.session_state.pending_conti_update = new_c
                                st.rerun()
                                
                    next_num = len(cur_songs) + 1
                    if st.button(f"➕ 콘티 끝에 추가 ({next_num}번)", key=f"btn_add_lib_{s_title}_{v['date']}", use_container_width=True):
                        new_c = append_song_to_conti(current_c, v['data'])
                        st.session_state.pending_conti_update = new_c
                        st.rerun()
    else:
        st.caption("일치하는 찬양곡이 없습니다.")

# Stateful Tabs (100% Cross-origin Cloud Safe)
tab_options = ["1. 콘티 작성", "2. PPT 기획안", "3. 클라우드 보관소", "4. 프로그램 & 설명서 다운로드"]

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
    st.subheader("1단계: 콘티 서식 작성")
    
    cur_conti = st.session_state.get("conti_editor", st.session_state.conti_text)
    parsed_preview = parse_user_conti(cur_conti)
    
    c_m1, c_m2, c_m3, c_m4 = st.columns(4)
    with c_m1:
        st.metric("예배 날짜", parsed_preview.get("date_str", "미입력"))
    with c_m2:
        st.metric("모임명", parsed_preview.get("title", "미입력"))
    with c_m3:
        st.metric("대표기도자", parsed_preview.get("prayer_person", "미입력"))
    with c_m4:
        songs_cnt = len(parsed_preview.get("songs", []))
        st.metric("등록 찬양곡", f"{songs_cnt} 곡")
        
    st.divider()
    
    col_edit, col_guide = st.columns([7.2, 4.8])
    
    with col_edit:
        user_conti = st.text_area(
            "콘티 본문",
            height=540,
            key="conti_editor",
            placeholder="# 2026.09.13 LOGOS 청년 모임\n기도: ㅇㅇㅇ 청년\n\n## 1. \n루틴: \n\n[V]\n\n[P]\n\n[C]\n\n## 2. \n루틴: \n\n[V]\n\n[P]\n\n[C]",
            label_visibility="collapsed"
        )
        st.session_state.conti_text = user_conti
            
        if st.button("PPT 기획안 생성 ➔", type="primary", use_container_width=True):
            if not st.session_state.conti_text.strip():
                st.warning("⚠️ 콘티 내용을 입력하거나 .txt 파일을 업로드한 후 버튼을 눌러주세요!")
            else:
                # 기획안 생성 시 콘티의 찬양곡들을 현재 선택된 부서 라이브러리에 자동 아카이빙/저장
                dept_for_save = "기타" if template_mode == "내 PC에서 직접 업로드" else normalize_department_name(active_template_name or "청년부")
                saved_files = decompose_and_save_songs_from_conti(st.session_state.conti_text, department=dept_for_save)
                if saved_files:
                    st.toast(f"💾 찬양곡 {len(saved_files)}곡이 [{dept_for_save}] 라이브러리에 자동 저장/업데이트되었습니다.")
                    
                parsed = parse_user_conti(st.session_state.conti_text)
                new_plan = generate_plan_text_from_conti(parsed)
                st.session_state.plan_text = new_plan
                st.session_state.plan_editor = new_plan
                st.session_state.switch_to_tab = "2. PPT 기획안"
                st.rerun()
            
    with col_guide:
        lint_res = validate_conti_text(parsed_preview)
        
        st.markdown("##### ⚡ 실시간 서식 검사")
        if lint_res["warnings"]:
            for w in lint_res["warnings"]:
                st.warning(f"⚠️ {w}")
        else:
            st.success("✅ 콘티 서식이 올바릅니다.")
            
        # 비간섭형 찬양 라이브러리 자동 감지
        all_lib = get_all_library_songs()
        matched_in_conti = []
        for s_idx, s in enumerate(parsed_preview.get("songs", [])):
            s_raw = s.get("title", "").strip()
            clean_s = re.sub(r'^\d+[\.\s]+', '', s_raw).strip()
            for lib_title, vers in all_lib.items():
                if (clean_s and clean_s.lower() == lib_title.lower()) or (s_raw and s_raw.lower() == lib_title.lower()):
                    matched_in_conti.append((s_idx + 1, lib_title, vers))
                    break
                    
        if matched_in_conti:
            st.divider()
            st.markdown("##### 📚 라이브러리 일치 찬양 감지")
            st.caption("기존에 불렀던 버전의 가사와 송폼으로 본문을 1:1 교체하거나 추가할 수 있습니다.")
            for s_order, lib_title, vers in matched_in_conti:
                with st.expander(f"🎵 곡 {s_order}번: {lib_title} ({len(vers)}개 버전)", expanded=True):
                    for v in vers:
                        st.markdown(f"**날짜**: `{v['date']}`")
                        if v.get('routine'):
                            st.caption(f"루틴: `{v['routine']}`")
                        col_btn_rep, col_btn_add = st.columns([1.1, 0.9])
                        with col_btn_rep:
                            if st.button(f"🔄 {s_order}번 곡 전체 교체", key=f"rec_rep_{lib_title}_{v['date']}_{s_order}", help=f"콘티의 {s_order}번 곡 내용을 이 버전으로 교체합니다.", use_container_width=True):
                                new_c = replace_song_in_conti(cur_conti, s_order, v['data'])
                                st.session_state.pending_conti_update = new_c
                                st.rerun()
                        with col_btn_add:
                            if st.button(f"➕ 끝에 새 곡 추가", key=f"rec_add_{lib_title}_{v['date']}_{s_order}", help="이 찬양을 콘티 맨 끝에 새로운 순번으로 추가합니다.", use_container_width=True):
                                new_c = append_song_to_conti(cur_conti, v['data'])
                                st.session_state.pending_conti_update = new_c
                                st.rerun()

        st.divider()
        with st.expander("📌 콘티 작성 규칙 안내", expanded=False):
            st.markdown("""
            고정 정보(# 모임명, 기도자), 곡 제목(##), 루틴 문자열, 파트([V], [C] 등)를 표준 서식에 맞춰 작성하세요. 슬라이드는 **빈 줄(더블 엔터)**로 자동 분할됩니다.
            
            - **표지 정보**: `# 2026.09.13 LOGOS 청년 모임`
            - **대표기도자**: `기도: ㅇㅇㅇ 청년`
            - **찬양곡 시작**: `## 1. 새 힘 얻으리`
            - **루틴(송폼)**: `루틴: Intro(8) - V - C - V - C...` (악보 기준)
            - **파트 선언**: `[V]`, `[V1]`, `[V2]`, `[P]`, `[C]`, `[B]` 등
            - **슬라이드 분할**: **빈 줄 하나(더블 엔터)**를 넣으면 다음 슬라이드로 넘어갑니다. (1슬라이드 당 1~2줄 권장)
            """)

# ==========================================
# TAB 2: PPT PLAN TEXT (HUMAN GATE)
# ==========================================
elif active_tab == "2. PPT 기획안":
    st.subheader("2단계: PPT 기획안 검토 및 수정")
    st.caption("생성된 슬라이드 텍스트를 검토하고 오탈자나 [BLANK] 암전 위치를 자유롭게 수정하세요.")
    
    # 상단 메트릭 자리 예약 (실시간 동기화)
    metric_placeholder = st.empty()
        
    st.divider()
    
    st.markdown("##### ✏️ 슬라이드 기획안 편집 (IDE 코드 에디터 스타일)")
    st.caption("좌측 넘버링 거터에 슬라이드 번호(1, 2, 3...)가 실시간으로 표시됩니다. 빈 줄(더블 엔터)로 슬라이드가 구분되며 스크롤이 완벽히 동기화됩니다.")
    
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
    
    # 상단 메트릭에 최신 수치를 1사이클 지연 없이 즉시 반영
    with metric_placeholder.container():
        c_stat1, c_stat2, c_stat3 = st.columns([3, 5, 4])
        with c_stat1:
            st.metric("총 슬라이드 수", f"{total_slide_count} 장")
        with c_stat2:
            template_display = active_template_name if active_template_name else "미지정 (사이드바 설정 필요)"
            st.metric("적용 기준 PPT", template_display)
        with c_stat3:
            st.metric("적용 표지 이미지", active_cover_name)
        
    st.divider()
    
    is_template_ready = (active_template_source is not None)
    
    if not is_template_ready:
        st.warning("⚠️ **[필수] 기준 PPT가 설정되지 않았습니다.**\n\n좌측 사이드바의 **「📁 기준 PPT 설정」**에서 [교회 표준 프리셋]을 선택하거나 [직접 파일 업로드]를 완료해야 PPT를 생성할 수 있습니다.")
    else:
        st.info(f"🎯 **적용 중인 기준 PPT**: `{active_template_name}`")
        
    if st.button("PPT 파일 생성", type="primary", use_container_width=True, disabled=not is_template_ready):
        if not current_slides:
            st.warning("⚠️ 슬라이드 기획안 내용이 비어있습니다. 1단계에서 콘티를 작성 후 생성해 주세요.")
        else:
            with st.spinner("기준 PPT 서식 및 폰트 임베딩을 1:1 딥클론하여 PPTX 빌드 중..."):
                prs = build_praise_pptx(current_slides, template_source=active_template_source, cover_image=active_cover_source)
                
                parsed = parse_user_conti(st.session_state.get("conti_editor", st.session_state.conti_text))
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
                    conti_content = st.session_state.get("conti_editor", st.session_state.conti_text)
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
# TAB 3: CLOUD STORAGE & ARCHIVE
# ==========================================
elif active_tab == "3. 클라우드 보관소":
    st.subheader("3단계: 텔레그램 클라우드 보관소")
    st.caption("교회 텔레그램 비공개 채널에 영구 보존된 찬양 PPT 및 찬양곡 자산을 실시간 열람하고 복원할 수 있습니다.")

    # 상단 상태 및 컨트롤 바
    c_hdr1, c_hdr2 = st.columns([8, 4])
    with c_hdr1:
        st.markdown("""
        <div class="cloud-channel-info-box" style="background: rgba(37, 99, 235, 0.05); border: 1px solid rgba(37, 99, 235, 0.2); padding: 10px 16px; border-radius: 8px; height: 74px; box-sizing: border-box; display: flex; flex-direction: column; justify-content: center;">
            <div style="font-weight: 700; color: #1e3a8a; font-size: 0.92rem; line-height: 1.35;">
                📢 연동 채널: <code>찬양 PPT 라이브러리</code> &nbsp;|&nbsp; 봇: <code>@praise_PPT_library_manager_bot</code>
            </div>
            <div style="font-size: 0.81rem; color: #475569; margin-top: 4px; line-height: 1.35;">
                🔄 자동 동기화: 프로그램 시작 시 및 매 30분마다 원격 갱신 사항을 자동으로 확인합니다.
            </div>
        </div>
        """, unsafe_allow_html=True)
    with c_hdr2:
        if st.button("🔄 클라우드 즉시 동기화", use_container_width=True, help="텔레그램 채널의 최신 카탈로그 및 변경사항을 지금 즉시 새로고침합니다."):
            with st.spinner("원격 클라우드 카탈로그 동기화 중..."):
                sync_summary = sync_on_startup(active_template_name)
                st.session_state.last_cloud_sync_time = time.time()
                ota_cnt = sync_summary.get("ota_presets_count", 0)
                if ota_cnt > 0:
                    st.toast(f"🎉 텔레그램에서 원격 표지/프리셋 {ota_cnt}건이 최신으로 갱신되었습니다!")
                else:
                    st.toast("✅ 클라우드 동기화 완료!")
                st.rerun()

        st.link_button(
            "✈️ 텔레그램 채널 바로가기",
            TELEGRAM_CHANNEL_LINK,
            use_container_width=True,
            help="텔레그램 공식 클라우드 채널로 이동하여 업로드된 PPT 및 찬양곡을 확인합니다."
        )

    st.divider()

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
            [f"📊 [{current_dept}] 찬양 PPT", "🌐 전체 부서 찬양 PPT", "🎵 찬양곡 라이브러리"],
            horizontal=True,
            label_visibility="collapsed"
        )
    with col_f2:
        search_kw = st.text_input("보관소 검색", placeholder="제목 또는 날짜 검색 (예: 2026.09, 꽃들도)", label_visibility="collapsed")

    catalog = load_local_catalog()
    all_items = catalog.get("items", [])

    # 1. 찬양 PPT 뷰
    if "찬양 PPT" in archive_view:
        is_dept_filtered = f"[{current_dept}]" in archive_view
        ppt_items = [it for it in all_items if it.get("type") == "PPT"]
        
        # 로컬 폴더(DIR_PPTX)에만 존재하는 파일도 카탈로그에 없는 경우 목록에 포함
        catalog_filenames = {it.get("filename") for it in ppt_items}
        if os.path.exists(DIR_PPTX):
            for f in os.listdir(DIR_PPTX):
                if f.endswith(".pptx") and f not in catalog_filenames:
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
        st.caption(f"총 {len(ppt_items)}개의 PPT 등록됨 (🟢 내 PC 보관: {local_count}개 | ☁️ 클라우드 전용: {cloud_only_count}개)")

        if not ppt_items:
            st.info(f"ℹ️ 등록된 [{current_dept if is_dept_filtered else '전체'}] 찬양 PPT가 없습니다. 2단계에서 PPT를 생성하면 이곳에 자동으로 백업됩니다.")
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
                    "<span style='background:#dcfce7; color:#15803d; border:1px solid #86efac; padding:2px 8px; border-radius:12px; font-size:0.8rem; font-weight:700;'>🟢 내 PC 보관 중</span>"
                    if is_local else
                    "<span style='background:#f1f5f9; color:#475569; border:1px solid #cbd5e1; padding:2px 8px; border-radius:12px; font-size:0.8rem; font-weight:600;'>☁️ 클라우드 전용 (PC 미보유)</span>"
                )

                with st.container():
                    c_card_info, c_card_act = st.columns([7, 5])
                    with c_card_info:
                        st.markdown(f"**📊 {ftitle}** &nbsp; <span style='background:#e0e7ff; color:#3730a3; padding:2px 8px; border-radius:12px; font-size:0.8rem;'>{dept_badge}</span> &nbsp; {status_badge} &nbsp; <span style='color:#64748b; font-size:0.85rem;'>일자: {fdate} | {fsize_str}</span>", unsafe_allow_html=True)
                        st.caption(f"파일명: `{fname}`")
                    with c_card_act:
                        if is_local:
                            st.markdown("<div style='color:#15803d; font-weight:700; font-size:0.85rem; margin-bottom:6px;'>✅ 보관함에 저장되어 있습니다</div>", unsafe_allow_html=True)
                            if os.path.exists(local_fpath):
                                with open(local_fpath, "rb") as f_dl:
                                    st.download_button(
                                        "📥 내 기기로 다운로드",
                                        data=f_dl.read(),
                                        file_name=fname,
                                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                        key=f"dl_local_{fname}",
                                        use_container_width=True
                                    )
                        else:

                            file_id = it.get("file_id")
                            if file_id:
                                if st.button(f"📥 클라우드에서 PC로 내려받기", key=f"dl_cloud_{fname}", use_container_width=True, type="primary"):
                                    with st.spinner("클라우드에서 다운로드 중..."):
                                        dl_ok, dl_err = download_file_by_id(file_id, local_fpath)
                                        if dl_ok:
                                            st.toast(f"✅ PC 보관함으로 내려받기 완료: {fname}")
                                            st.rerun()
                                        else:
                                            st.error(f"다운로드 실패: {dl_err}")
                            else:
                                st.caption("클라우드 파일 ID 없음")
                    st.divider()

    # 2. 찬양곡 라이브러리 뷰
    else:
        song_items = [it for it in all_items if it.get("type") == "찬양곡"]
        if search_kw.strip():
            kw = search_kw.strip().lower()
            song_items = [it for it in song_items if kw in it.get("title", "").lower() or kw in it.get("filename", "").lower()]

        local_song_count = sum(1 for it in song_items if os.path.exists(os.path.join(DIR_SONGS, sanitize_filename_part(it.get("title", "")), it.get("filename", ""))))
        cloud_only_song_count = len(song_items) - local_song_count
        st.caption(f"총 {len(song_items)}개의 찬양곡 버전 등록됨 (🟢 내 PC 보관: {local_song_count}개 | ☁️ 클라우드 전용: {cloud_only_song_count}개)")

        if not song_items:
            st.info("ℹ️ 클라우드에 보관된 찬양곡이 없습니다.")
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
                    "<span style='background:#dcfce7; color:#15803d; border:1px solid #86efac; padding:2px 8px; border-radius:12px; font-size:0.8rem; font-weight:700;'>🟢 내 PC 보관 중</span>"
                    if is_local else
                    "<span style='background:#f1f5f9; color:#475569; border:1px solid #cbd5e1; padding:2px 8px; border-radius:12px; font-size:0.8rem; font-weight:600;'>☁️ 클라우드 전용 (PC 미보유)</span>"
                )

                with st.container():
                    c_sinfo, c_sact = st.columns([7, 5])
                    with c_sinfo:
                        st.markdown(f"**🎵 {stitle}** &nbsp; <span style='background:#fef3c7; color:#92400e; padding:2px 8px; border-radius:12px; font-size:0.8rem;'>{dept_badge}</span> &nbsp; {status_badge} &nbsp; <span style='color:#64748b; font-size:0.85rem;'>버전일자: {sdate}</span>", unsafe_allow_html=True)
                        st.caption(f"파일명: `{fname}`")
                    with c_sact:
                        if is_local:
                            st.markdown("<div style='color:#15803d; font-weight:700; font-size:0.85rem; margin-bottom:6px;'>✅ 내 PC 라이브러리에 보유 중</div>", unsafe_allow_html=True)
                        else:

                            file_id = it.get("file_id")
                            if file_id:
                                if st.button("📥 클라우드에서 PC로 내려받기", key=f"dl_song_{fname}", use_container_width=True, type="primary"):
                                    with st.spinner("찬양곡 다운로드 중..."):
                                        os.makedirs(local_dir, exist_ok=True)
                                        dl_ok, dl_err = download_file_by_id(file_id, local_spath)
                                        if dl_ok:
                                            st.toast(f"✅ 라이브러리에 저장 완료: {stitle}")
                                            st.rerun()
                                        else:
                                            st.error(f"다운로드 실패: {dl_err}")
                    st.divider()

# ==========================================
# TAB 4: DESKTOP APP & MANUAL DOWNLOAD
# ==========================================
elif active_tab == "4. 프로그램 & 설명서 다운로드":
    st.subheader("4단계: 데스크톱 전용 프로그램 & 사용 설명서")
    st.caption("고성능 윈도우 단독 프로그램을 다운로드하거나, 웹에서 즉시 최신 공식 설명서를 열람할 수 있습니다.")

    # 1. 윈도우 전용 데스크톱 프로그램 (.exe) 다운로드 카드
    with st.container():
        st.markdown("""
        <div style="background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%); border: 1.5px solid #93c5fd; border-radius: 12px; padding: 22px 26px; margin-bottom: 24px;">
            <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px;">
                <div style="font-size: 1.35rem; font-weight: 800; color: #1e3a8a;">
                    💻 LOGOS 찬양 PPT 제작 스튜디오 v1.5.0 (Windows 전용 설치기)
                </div>
                <span style="background: #2563eb; color: white; padding: 4px 12px; border-radius: 20px; font-size: 0.82rem; font-weight: 700;">공식 배포판</span>
            </div>
            <div style="font-size: 0.95rem; color: #334155; line-height: 1.6; margin-bottom: 16px;">
                인터넷 브라우저 없이 PC에서 단독 실행되는 초경량 전용 소프트웨어 창입니다.<br>
                생성된 PPT를 윈도우 탐색기로 자동 연결하며, 완전 오프라인 환경에서도 로컬 라이브러리를 통해 즉시 PPT를 제작할 수 있습니다.
            </div>
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px; margin-bottom: 16px; font-size: 0.88rem; color: #1e40af; font-weight: 600;">
                <div>✨ 브라우저 없는 단독 독립 창 (pywebview)</div>
                <div>📂 완성 PPT 파일 위치 탐색기 자동 열기</div>
                <div>🔒 텔레그램 클라우드 자동 양방향 동기화</div>
                <div>⚡ 초고속 네이티브 PPTX 렌더링 엔진</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        col_dl1, col_dl2 = st.columns([7, 5])
        with col_dl1:
            st.link_button(
                "🚀 Windows 전용 설치기 (.exe) 다운로드 (GitHub Releases)",
                "https://github.com/nimba789/worship-ppt-generator/releases",
                type="primary",
                use_container_width=True,
                help="GitHub Releases 공식 저장소에서 최신 설치기 파일을 바로 내려받습니다."
            )
        with col_dl2:
            st.caption("ℹ️ 설치기 파일 용량: 약 175 MB | 지원 OS: Windows 10 / 11 (64bit)")

    st.divider()

    # 2. 프로그램 공식 사용 설명서
    st.markdown("#### 📖 LOGOS 찬양 PPT 제작 스튜디오 공식 설명서")
    
    tab_m1, tab_m2 = st.tabs(["🚀 [1부] 3분 완성 빠른 사용 흐름 가이드", "📚 [2부] 상세 기능 및 부록 (FAQ)"])
    
    with tab_m1:
        st.markdown("""
### 1단계: 프로그램 실행 및 기준 PPT 프리셋 선택
- 사이드바의 **[📁 기준 PPT 설정]**에서 사용할 프리셋(예: **[청년부] 기준 PPT**)을 선택합니다.
- 특정 부서 프리셋이 없는 경우 **[내 PC에서 직접 업로드]**를 선택하여 교회의 기존 찬양 PPT를 등록합니다.

---

### 2단계: 1. 찬양 콘티 작성
- 우측 사이드바의 **[🎵 찬양곡 라이브러리 검색]**에서 곡 제목을 검색한 후 **[➕ 콘티 끝에 추가]** 또는 **[🔄 n번 교체]**를 누르면 가사와 루틴이 자동으로 채워집니다.
- 콘티 작성이 완료되면 하단의 **[➡️ 2단계: PPT 기획안 생성]** 버튼을 누릅니다.

---

### 3단계: 2. PPT 기획안 검토 및 수정
- 생성된 슬라이드 기획안이 **IDE 다크 에디터**에 슬라이드 번호(`1 |`, `2 |` ...)와 함께 정렬됩니다.
- 오탈자를 교정하거나, 필요시 **엔터 두 번(빈 줄)**을 입력하여 새 슬라이드로 분할할 수 있습니다.
- 준비가 끝나면 하단의 **[🎨 찬양 PPT 자동 생성 및 백업]** 버튼을 누릅니다.

---

### 4단계: 3. PPT 다운로드 및 보관
- 생성이 완료되면 **[💾 PPT 파일 다운로드]** 버튼으로 내 기기에 즉시 저장합니다.
- 동시에 완성된 PPT와 분해된 찬양곡 가사는 **교회 텔레그램 클라우드 공식 채널에 자동 백업**되어 영구 보존됩니다.
        """)

    with tab_m2:
        st.markdown("""
### 📌 부록 A: 텔레그램 클라우드 봇 명령어
찬양팀 텔레그램 채널에서 봇을 통해 언제든 자료를 추가하거나 관리할 수 있습니다:
- **`#추가`**: `.pptx` 파일과 함께 `#추가 #청년부` 형태로 올리면 교회 공용 프리셋으로 즉시 등록됩니다.
- **`#곡등록`**: 가사 JSON 파일이나 텍스트를 첨부하여 찬양곡 라이브러리에 원격 등록합니다.
- **`#목록`**: 현재 클라우드에 등록된 프리셋 및 최신 PPT 목록을 실시간 조회합니다.

---

### 📌 부록 B: 기준 PPT 프리셋 권장 규격
- **화면 비율**: 16:9 와이드스크린 권장
- **텍스트 박스**: 제목용 텍스트 상자와 본문 가사용 텍스트 상자가 명확히 구분된 단일 마스터 슬라이드 구조 권장
- **폰트**: 나눔스퀘어라운드, 프리텐다드, 맑은 고딕 등 가독성이 높은 폰트 사용

---

### 📌 부록 C: 자주 묻는 질문 (FAQ)
- **Q. 폰트가 깨지거나 크기가 맞지 않나요?**
  - A. 기준 PPT로 등록한 원본 PPT의 텍스트 상자 서식과 폰트 크기, 줄 간격을 시스템이 1:1로 정확하게 복제하므로 기준 PPT 파일의 텍스트 상자를 원하는 서식으로 맞춰두시면 항상 동일하게 생성됩니다.
- **Q. 아이패드나 모바일에서도 사용할 수 있나요?**
  - A. 네! 현재 보고 계시는 웹 프로그램을 통해 아이패드, 갤럭시탭, 스마트폰 어디서든 콘티 작성과 PPT 생성이 가능하며 생성 즉시 기기로 다운로드됩니다.
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

