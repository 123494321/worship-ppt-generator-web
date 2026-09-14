"""
core/admin_component.py
독립형 '관리자 모드(Admin Console)' UI 및 컨트롤러
1. 표준 양식 설정 관리 (추가, 수정/교체, 삭제 모달 다이얼로그)
2. 자료실 관리 (전체 찬양 PPT, 찬양곡 자료실, 🗑️ 휴지통 소프트/하드 삭제)
3. 클라우드 관리 (스토리지 용량 분석 대시보드, 텔레그램 봇 유용한 제어 기능)
"""

import os
import shutil
from datetime import datetime
import streamlit as st

from core.cloud_config import (
    DIR_PPTX,
    DIR_SONGS,
    DIR_PRESETS,
    CHAT_ID,
    ensure_data_directories,
    sanitize_filename_part,
    normalize_department_name,
    parse_standard_filename
)
from core.cloud_sync import (
    load_local_catalog,
    save_local_catalog,
    publish_catalog,
    sync_on_startup,
    ensure_channel_guide_message,
    telegram_api_call
)
from core.pptx_builder import get_preset_info_dict
from core.song_library import get_all_library_songs
from core.trash_manager import (
    get_trash_catalog,
    is_pptx_trashed,
    is_song_trashed,
    move_pptx_to_trash,
    restore_pptx_from_trash,
    hard_delete_pptx,
    move_song_to_trash,
    restore_song_from_trash,
    hard_delete_song,
    empty_all_trash
)
from core.preset_manager import (
    add_preset_department,
    update_preset_department,
    delete_preset_department
)


# ==============================================================================
# 1. 모달 다이얼로그 (@st.dialog)
# ==============================================================================

@st.dialog("➕ 새 표준 양식(부서) 등록")
def show_add_preset_dialog():
    st.caption("새로운 교회 부서나 모임의 공식 기준 PPT 서식 및 표지를 등록합니다.")
    dept_name = st.text_input("부서 / 양식 이름", placeholder="예: 중고등부, 초등부, 대학부")
    pptx_file = st.file_uploader("기준 PPTX 파일 (.pptx) *필수", type=["pptx"])
    cover_file = st.file_uploader("표지 이미지 파일 (.png, .jpg) (선택)", type=["png", "jpg", "jpeg"])

    c_cancel, c_save = st.columns([1, 1])
    with c_cancel:
        if st.button("취소", use_container_width=True):
            st.rerun()
    with c_save:
        if st.button("저장", type="primary", use_container_width=True):
            if not dept_name or not dept_name.strip():
                st.error("부서 이름을 입력해 주세요.")
                return
            if not pptx_file:
                st.error("기준 PPTX 파일을 업로드해야 합니다.")
                return
            with st.spinner("새 표준 양식을 클라우드에 등록 중..."):
                p_bytes = pptx_file.read()
                c_bytes = cover_file.read() if cover_file else None
                c_name = cover_file.name if cover_file else None
                ok, msg = add_preset_department(
                    dept_name.strip(),
                    p_bytes,
                    pptx_filename=pptx_file.name,
                    cover_bytes=c_bytes,
                    cover_filename=c_name
                )
                if ok:
                    st.toast(f"✅ {msg}")
                    st.rerun()
                else:
                    st.error(msg)


@st.dialog("⚙️ 표준 양식 설정 및 교체")
def show_edit_preset_dialog(preset_name):
    st.markdown(f"### 📌 **[{preset_name}]** 표준 양식 설정")
    st.caption("서식 이름 변경, 기준 PPT 파일 교체, 표지 이미지 교체 및 삭제를 진행합니다.")

    new_name = st.text_input("표준 양식 이름", value=preset_name)
    new_pptx = st.file_uploader("새 기준 PPTX 파일로 교체 (.pptx)", type=["pptx"], key=f"edit_pptx_{preset_name}")
    new_cover = st.file_uploader("새 표지 이미지로 교체 (.png, .jpg)", type=["png", "jpg", "jpeg"], key=f"edit_cover_{preset_name}")
    remove_cover = st.checkbox("현재 표지 이미지 삭제 (기본 텍스트 표지로 전환)", key=f"edit_rmcover_{preset_name}")

    st.divider()

    # 삭제 옵션 (청년부 제외)
    if preset_name != "청년부":
        with st.expander("🚨 표준 양식 삭제", expanded=False):
            st.warning(f"'{preset_name}' 표준 양식을 완전히 삭제하면 복구할 수 없습니다.")
            confirm_del = st.checkbox("정말로 이 표준 양식을 삭제합니다.", key=f"confirm_del_{preset_name}")
            if st.button("🗑️ 표준 양식 영구 삭제", type="secondary", disabled=not confirm_del, use_container_width=True):
                with st.spinner("표준 양식 삭제 중..."):
                    ok, msg = delete_preset_department(preset_name)
                    if ok:
                        st.toast(f"✅ {msg}")
                        st.rerun()
                    else:
                        st.error(msg)

    c_cancel, c_change = st.columns([1, 1])
    with c_cancel:
        if st.button("취소", use_container_width=True, key=f"btn_cancel_{preset_name}"):
            st.rerun()
    with c_change:
        if st.button("변경 저장", type="primary", use_container_width=True, key=f"btn_save_{preset_name}"):
            p_bytes = new_pptx.read() if new_pptx else None
            p_name = new_pptx.name if new_pptx else None
            c_bytes = new_cover.read() if new_cover else None
            c_name = new_cover.name if new_cover else None

            with st.spinner("표준 양식 설정을 업데이트 중..."):
                ok, msg = update_preset_department(
                    old_dept_name=preset_name,
                    new_dept_name=new_name.strip() if new_name else None,
                    pptx_bytes=p_bytes,
                    pptx_filename=p_name,
                    cover_bytes=c_bytes,
                    cover_filename=c_name,
                    remove_cover=remove_cover
                )
                if ok:
                    st.toast(f"✅ {msg}")
                    st.rerun()
                else:
                    st.error(msg)


# ==============================================================================
# 2. 메인 관리자 콘솔 렌더링 함수
# ==============================================================================

def render_admin_console():
    """
    독립된 관리자 모드 전용 대시보드를 렌더링합니다.
    """
    # 1. 관리자 전용 상단 헤더 및 모드 전환 바
    st.markdown("""
    <div style="background: linear-gradient(135deg, #1e1b4b 0%, #312e81 50%, #4338ca 100%); padding: 18px 24px; border-radius: 12px; color: white; margin-bottom: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.15);">
        <div style="display: flex; align-items: center; justify-content: space-between;">
            <div>
                <div style="font-size: 1.45rem; font-weight: 800; letter-spacing: -0.5px; display: flex; align-items: center; gap: 8px;">
                    🛡️ LOGOS 관리자 모드 <span style="font-size: 0.75rem; background: rgba(255,255,255,0.2); padding: 2px 8px; border-radius: 12px; font-weight: 600;">Admin Console</span>
                </div>
                <div style="font-size: 0.85rem; color: #c7d2fe; margin-top: 4px;">
                    교회 표준 양식, 찬양곡 자료실 및 클라우드 스토리지를 중앙 집중 관리합니다.
                </div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # 상단 컨트롤 버튼 (일반 모드로 전환 / 로그아웃)
    c_sw1, c_sw2, c_sp = st.columns([2.5, 2, 7.5])
    with c_sw1:
        if st.button("👤 일반 사용자 모드로 전환", use_container_width=True):
            st.session_state.app_mode = "general"
            st.rerun()
    with c_sw2:
        if st.button("🔒 로그아웃", use_container_width=True):
            st.session_state.authenticated = False
            st.session_state.app_mode = "general"
            st.rerun()

    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

    # 관리자 모드 3개 탭
    admin_tab_options = ["1. 표준 양식 설정 관리", "2. 자료실 관리", "3. 클라우드 관리"]
    if "admin_selected_tab" not in st.session_state:
        st.session_state.admin_selected_tab = admin_tab_options[0]

    admin_active_tab = st.radio(
        "관리자 메뉴 선택",
        admin_tab_options,
        key="admin_selected_tab",
        horizontal=True,
        label_visibility="collapsed"
    )

    st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)

    # ==========================================================================
    # TAB 1: 표준 양식 설정 관리
    # ==========================================================================
    if admin_active_tab == "1. 표준 양식 설정 관리":
        col_t1_title, col_t1_btn = st.columns([8, 2.5])
        with col_t1_title:
            st.markdown("### 📋 표준 양식(부서) 목록")
            st.caption("등록된 표준 양식은 일반 사용자 모드의 사이드바에 즉시 반영되며, 찬양 PPT 생성 서식으로 활용됩니다.")
        with col_t1_btn:
            if st.button("➕ 표준 양식 추가", type="primary", use_container_width=True):
                show_add_preset_dialog()

        preset_info = get_preset_info_dict()

        if not preset_info:
            st.info("ℹ️ 등록된 표준 양식이 없습니다. [➕ 표준 양식 추가] 버튼을 눌러 첫 양식을 등록해 보세요.")
        else:
            for p_name, p_data in preset_info.items():
                t_path = p_data.get("template_path")
                c_path = p_data.get("cover_path")
                c_name = p_data.get("cover_name")
                fname = os.path.basename(t_path) if t_path else "기준 PPT 없음"

                fsize_str = "0 KB"
                if t_path and os.path.exists(t_path):
                    sz = os.path.getsize(t_path)
                    fsize_str = f"{sz / (1024*1024):.1f} MB" if sz >= 1024*1024 else f"{sz / 1024:.0f} KB"

                with st.container():
                    st.markdown("""<div style="background:#ffffff; border:1px solid #e2e8f0; border-radius:10px; padding:12px; margin-bottom:12px; box-shadow: 0 1px 3px rgba(0,0,0,0.03);">""", unsafe_allow_html=True)
                    c_n, c_thumb, c_file, c_btn = st.columns([2.5, 2.2, 5, 2.3])
                    
                    with c_n:
                        st.markdown(f"**📌 {p_name}**")
                        if p_name == "청년부":
                            st.caption("기본 공식 양식")
                        else:
                            st.caption("사용자 정의 양식")

                    with c_thumb:
                        if c_path and os.path.exists(c_path):
                            st.image(c_path, use_container_width=True)
                        else:
                            st.caption("🖼️ 표지 이미지 없음\n(기본 텍스트 표지)")

                    with c_file:
                        st.markdown(f"**기준 파일:** `{fname}`")
                        st.caption(f"파일 크기: {fsize_str}")

                    with c_btn:
                        if st.button("⚙️ 설정", key=f"btn_edit_p_{p_name}", use_container_width=True):
                            show_edit_preset_dialog(p_name)

                    st.markdown("</div>", unsafe_allow_html=True)

    # ==========================================================================
    # TAB 2: 자료실 관리 (3개 하위 탭)
    # ==========================================================================
    elif admin_active_tab == "2. 자료실 관리":
        subtab_ppt, subtab_song, subtab_trash = st.tabs(["📊 전체 찬양 PPT", "🎵 찬양곡 자료실", "🗑️ 휴지통"])

        # ----------------------------------------------------------------------
        # 하위 탭 1: 전체 찬양 PPT
        # ----------------------------------------------------------------------
        with subtab_ppt:
            col_search, _ = st.columns([6, 6])
            with col_search:
                kw_ppt = st.text_input("PPT 검색", placeholder="제목 또는 일자 검색", key="admin_ppt_kw", label_visibility="collapsed")

            catalog = load_local_catalog()
            all_items = catalog.get("items", [])
            ppt_items = [it for it in all_items if it.get("type") == "PPT" and not is_pptx_trashed(it.get("filename"))]

            # 로컬 파일 보강
            cat_fnames = {it.get("filename") for it in ppt_items}
            if os.path.exists(DIR_PPTX):
                for f in os.listdir(DIR_PPTX):
                    if f.endswith(".pptx") and f not in cat_fnames and not is_pptx_trashed(f):
                        meta = parse_standard_filename(f)
                        ppt_items.append({
                            "filename": f,
                            "file_size": os.path.getsize(os.path.join(DIR_PPTX, f)),
                            "title": meta["title"] if meta else f,
                            "date": meta["date"] if meta else "",
                            "preset": meta["preset"] if meta else "기타"
                        })

            if kw_ppt.strip():
                k = kw_ppt.strip().lower()
                ppt_items = [it for it in ppt_items if k in it.get("title", "").lower() or k in it.get("filename", "").lower() or k in it.get("date", "").lower()]

            st.caption(f"총 {len(ppt_items)}개의 PPT 등록됨 (삭제 시 휴지통으로 이동하며 일반 모드에서 즉시 숨겨집니다)")

            if not ppt_items:
                st.info("ℹ️ 관리할 찬양 PPT가 없습니다.")
            else:
                for it in ppt_items:
                    fname = it.get("filename")
                    ftitle = it.get("title", fname)
                    fdate = it.get("date", "")
                    dept_badge = normalize_department_name(it.get("preset", "기타"))
                    fsize = it.get("file_size", 0)
                    fsize_str = f"{fsize / (1024*1024):.1f} MB" if fsize >= 1024*1024 else f"{fsize / 1024:.0f} KB"

                    with st.container():
                        c_pinfo, c_pact = st.columns([8.5, 1.5])
                        with c_pinfo:
                            st.markdown(f"**📊 {ftitle}** &nbsp; <span style='background:#e0e7ff; color:#3730a3; padding:2px 8px; border-radius:12px; font-size:0.8rem;'>{dept_badge}</span> &nbsp; <span style='color:#64748b; font-size:0.85rem;'>일자: {fdate} | {fsize_str}</span>", unsafe_allow_html=True)
                            st.caption(f"파일명: `{fname}`")
                        with c_pact:
                            if st.button("🗑️ 삭제", key=f"trash_ppt_{fname}", type="secondary", use_container_width=True):
                                move_pptx_to_trash(fname)
                                st.toast(f"🗑️ '{ftitle}' 파일이 휴지통으로 이동되었습니다.")
                                st.rerun()
                        st.divider()

        # ----------------------------------------------------------------------
        # 하위 탭 2: 찬양곡 자료실 (곡명별 아코디언 묶음)
        # ----------------------------------------------------------------------
        with subtab_song:
            col_search_s, _ = st.columns([6, 6])
            with col_search_s:
                kw_song = st.text_input("찬양곡 검색", placeholder="곡 제목 검색", key="admin_song_kw", label_visibility="collapsed")

            all_songs = get_all_library_songs(include_trashed=False)

            if kw_song.strip():
                k_s = kw_song.strip().lower()
                all_songs = {t: v for t, v in all_songs.items() if k_s in t.lower()}

            st.caption(f"총 {len(all_songs)}곡 등록됨 (곡을 클릭하면 버전별 목록이 열리며, 개별 버전을 삭제할 수 있습니다)")

            if not all_songs:
                st.info("ℹ️ 등록된 찬양곡이 없습니다.")
            else:
                for song_title, versions in all_songs.items():
                    with st.expander(f"🎵 **{song_title}** ({len(versions)}개 버전)", expanded=False):
                        for v in versions:
                            c_vinfo, c_vact = st.columns([8.5, 1.5])
                            with c_vinfo:
                                r_text = v.get("routine") or "루틴 없음"
                                st.markdown(f"**일자: {v['date']}** &nbsp; <span style='background:#e0e7ff; color:#3730a3; padding:2px 8px; border-radius:12px; font-size:0.8rem;'>{v.get('department', '청년부')}</span> &nbsp; <span style='color:#475569; font-size:0.85rem;'>루틴: `{r_text}`</span>", unsafe_allow_html=True)
                            with c_vact:
                                if st.button("🗑️ 삭제", key=f"trash_song_{song_title}_{v['date']}", type="secondary", use_container_width=True):
                                    move_song_to_trash(song_title, v['date'], routine=v.get("routine", ""), department=v.get("department", "청년부"))
                                    st.toast(f"🗑️ '{song_title} ({v['date']})' 버전이 휴지통으로 이동되었습니다.")
                                    st.rerun()

        # ----------------------------------------------------------------------
        # 하위 탭 3: 🗑️ 휴지통 (Soft / Hard Delete & Restore)
        # ----------------------------------------------------------------------
        with subtab_trash:
            trash = get_trash_catalog()
            trash_pptx = trash.get("pptx", [])
            trash_songs = trash.get("songs", [])

            col_t_stat, col_t_empty = st.columns([8, 2.5])
            with col_t_stat:
                st.caption(f"휴지통 보관 중: PPT {len(trash_pptx)}개 | 찬양곡 {len(trash_songs)}개 (복구 시 일반 모드에 다시 노출되며, 영구 삭제 시 텔레그램에서도 완전 파기됩니다)")
            with col_t_empty:
                if trash_pptx or trash_songs:
                    if st.button("💥 휴지통 전체 영구 비우기", type="secondary", use_container_width=True):
                        dp, ds = empty_all_trash()
                        st.toast(f"💥 휴지통의 모든 항목(PPT {dp}개, 찬양곡 {ds}개)이 영구 삭제되었습니다.")
                        st.rerun()

            if not trash_pptx and not trash_songs:
                st.info("✨ 휴지통이 비어 있습니다.")
            else:
                if trash_pptx:
                    st.markdown("#### 📊 임시 삭제된 찬양 PPT")
                    for it in trash_pptx:
                        fname = it.get("filename")
                        ftitle = it.get("title", fname)
                        del_at = it.get("deleted_at", "")
                        with st.container():
                            c_t_info, c_t_res, c_t_del = st.columns([7, 1.5, 1.5])
                            with c_t_info:
                                st.markdown(f"**{ftitle}** &nbsp; <span style='color:#ef4444; font-size:0.8rem;'>삭제일시: {del_at}</span>", unsafe_allow_html=True)
                                st.caption(f"파일명: `{fname}`")
                            with c_t_res:
                                if st.button("↩️ 복구", key=f"res_ppt_{fname}", use_container_width=True):
                                    restore_pptx_from_trash(fname)
                                    st.toast(f"✅ '{ftitle}' 파일이 복구되었습니다.")
                                    st.rerun()
                            with c_t_del:
                                if st.button("❌ 영구 삭제", key=f"hard_ppt_{fname}", type="secondary", use_container_width=True):
                                    hard_delete_pptx(fname)
                                    st.toast(f"❌ '{ftitle}' 파일이 텔레그램 클라우드에서 영구 삭제되었습니다.")
                                    st.rerun()
                            st.divider()

                if trash_songs:
                    st.markdown("#### 🎵 임시 삭제된 찬양곡 버전")
                    for it in trash_songs:
                        stitle = it.get("title")
                        sdate = it.get("date")
                        del_at = it.get("deleted_at", "")
                        with st.container():
                            c_ts_info, c_ts_res, c_ts_del = st.columns([7, 1.5, 1.5])
                            with c_ts_info:
                                st.markdown(f"**{stitle} ({sdate})** &nbsp; <span style='color:#ef4444; font-size:0.8rem;'>삭제일시: {del_at}</span>", unsafe_allow_html=True)
                            with c_ts_res:
                                if st.button("↩️ 복구", key=f"res_s_{stitle}_{sdate}", use_container_width=True):
                                    restore_song_from_trash(stitle, sdate)
                                    st.toast(f"✅ '{stitle} ({sdate})' 버전이 복구되었습니다.")
                                    st.rerun()
                            with c_ts_del:
                                if st.button("❌ 영구 삭제", key=f"hard_s_{stitle}_{sdate}", type="secondary", use_container_width=True):
                                    hard_delete_song(stitle, sdate)
                                    st.toast(f"❌ '{stitle} ({sdate})' 버전이 텔레그램 클라우드에서 영구 삭제되었습니다.")
                                    st.rerun()
                            st.divider()

    # ==========================================================================
    # TAB 3: 클라우드 관리
    # ==========================================================================
    elif admin_active_tab == "3. 클라우드 관리":
        st.markdown("### ☁️ 클라우드 스토리지 현황")
        st.caption("텔레그램 공식 채널 클라우드 저장소의 자산 점유율 및 상태를 모니터링합니다.")

        catalog = load_local_catalog()
        items = catalog.get("items", [])
        presets = catalog.get("presets", {})
        trash = catalog.get("trash", {})

        # 파일 유형별 통계
        ppt_items = [it for it in items if it.get("type") == "PPT"]
        song_items = [it for it in items if it.get("type") == "찬양곡"]

        total_files = len(items) + len(presets)
        total_ppt_bytes = sum(it.get("file_size", 0) for it in ppt_items)
        total_song_bytes = sum(it.get("file_size", 0) for it in song_items)
        total_bytes = total_ppt_bytes + total_song_bytes

        total_mb = total_bytes / (1024 * 1024)
        ppt_mb = total_ppt_bytes / (1024 * 1024)
        song_kb = total_song_bytes / 1024

        ppt_pct = (total_ppt_bytes / total_bytes * 100) if total_bytes > 0 else 0
        song_pct = (total_song_bytes / total_bytes * 100) if total_bytes > 0 else 0

        c_m1, c_m2, c_m3 = st.columns(3)
        with c_m1:
            st.metric("📁 총 클라우드 파일 수", f"{total_files}개", help="PPTX 완성본, 찬양곡 JSON 및 표준 양식을 합산한 총 파일 수입니다.")
        with c_m2:
            st.metric("💾 총 저장 용량", f"{total_mb:.1f} MB", help="텔레그램 채널에 보관 중인 전체 자산의 용량입니다.")
        with c_m3:
            trash_total = len(trash.get("pptx", [])) + len(trash.get("songs", []))
            st.metric("🗑️ 휴지통 보관 항목", f"{trash_total}개", help="휴지통에 임시 보관 중인 항목 수입니다.")

        st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)
        st.markdown("#### 📊 파일 유형별 스토리지 점유율")
        
        col_p1, col_p2 = st.columns([1, 1])
        with col_p1:
            st.markdown(f"**📊 찬양 PPTX:** `{ppt_mb:.1f} MB` ({ppt_pct:.1f}%)")
            st.progress(min(max(ppt_pct / 100, 0.0), 1.0))
        with col_p2:
            st.markdown(f"**🎵 찬양곡 JSON:** `{song_kb:.1f} KB` ({song_pct:.1f}%)")
            st.progress(min(max(song_pct / 100, 0.0), 1.0))

        st.divider()

        st.markdown("### 🤖 텔레그램 봇 및 채널 제어 도구")
        st.caption("클라우드 채널의 무결성을 검사하고 수동 작업을 지시합니다.")

        # 봇 연결 상태 실시간 진단
        bot_check = telegram_api_call("getMe")
        if bot_check.get("ok"):
            b_info = bot_check.get("result", {})
            st.success(f"🟢 **텔레그램 봇 정상 연결됨:** @{b_info.get('username')} ({b_info.get('first_name')})")
        else:
            st.warning(f"🟡 **텔레그램 봇 연결 확인 필요:** {bot_check.get('error', '응답 없음')}")

        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

        c_act1, c_act2, c_act3 = st.columns(3)
        with c_act1:
            with st.container():
                st.markdown("**🔄 클라우드 강제 동기화**")
                st.caption("채널의 최신 카탈로그를 즉시 내려받아 로컬과 일치시킵니다.")
                if st.button("지금 전체 동기화", key="admin_force_sync", use_container_width=True):
                    with st.spinner("클라우드 동기화 중..."):
                        s_res = sync_on_startup()
                        st.toast("✅ 클라우드 전체 동기화가 완료되었습니다!")
                        st.rerun()

        with c_act2:
            with st.container():
                st.markdown("**📌 공식 안내 공지문 최신화**")
                st.caption("채널 상단의 고정 안내 공지문을 최신 카탈로그로 갱신합니다.")
                if st.button("안내문 최신화 및 고정", key="admin_update_guide", use_container_width=True):
                    with st.spinner("공식 안내문 갱신 중..."):
                        ok, g_res = ensure_channel_guide_message(force=True)
                        if ok:
                            st.toast("✅ 채널 상단 공식 안내문이 성공적으로 최신화되었습니다!")
                        else:
                            st.error(f"안내문 갱신 실패: {g_res}")

        with c_act3:
            with st.container():
                st.markdown("**🗑️ 휴지통 전체 비우기**")
                st.caption("휴지통에 남아있는 모든 PPT와 곡을 클라우드에서 일괄 영구 파기합니다.")
                if st.button("휴지통 일괄 파기", key="admin_empty_all", type="secondary", use_container_width=True):
                    with st.spinner("휴지통 전체 비우는 중..."):
                        dp, ds = empty_all_trash()
                        st.toast(f"💥 휴지통의 모든 항목(PPT {dp}개, 곡 {ds}개)이 영구 파기되었습니다.")
                        st.rerun()
