import os
import sys
import json
import logging
import streamlit as st
from streamlit import config
from streamlit.runtime import get_instance
import streamlit.components.v1.custom_component as cc_module
from streamlit.components.v1.custom_component import CustomComponent

_COMPONENT_DIR = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "components", "conti_editor"))
_conti_editor = None

def _get_or_register_conti_editor_component():
    """
    Directly instantiates and registers the conti_editor CustomComponent in Streamlit's ComponentRegistry.
    Guarantees seamless execution in both web and frozen environments.
    """
    global _conti_editor
    if _conti_editor is None:
        try:
            module_name = "core.conti_editor_component"
            name = "conti_editor"
            component_name = f"{module_name}.{name}"

            url = None
            if component_base_path := config.get_option("server.customComponentBaseUrlPath"):
                url = f"{component_base_path}/{component_name}/"

            _conti_editor = CustomComponent(
                name=component_name,
                path=str(_COMPONENT_DIR),
                url=url,
                module_name=module_name
            )
        except Exception as e:
            logging.error(f"[conti_editor] Direct CustomComponent creation failed: {e}")
            _conti_editor = None

    if _conti_editor is not None:
        try:
            inst = get_instance()
            if inst and hasattr(inst, "component_registry") and inst.component_registry is not None:
                inst.component_registry.register_component(_conti_editor)
        except Exception as e:
            logging.debug(f"[conti_editor] Runtime registration warning: {e}")

    return _conti_editor

def render_conti_editor(value="", height=540, key=None):
    """
    Renders the live conti editor component with:
    1. Real-time debounced auto-sync (no Ctrl+Enter required)
    2. Instant client-side localStorage auto-save & 1-click restore
    3. Seamless fallback to standard st.text_area if unavailable
    """
    editor = _get_or_register_conti_editor_component()
    if editor is not None:
        try:
            res = editor(value=value, height=height, key=key, default=value)
            return res if res is not None else value
        except Exception as e:
            logging.error(f"[conti_editor] Custom component rendering exception: {e}", exc_info=True)

    # 2중 안전망: 예외 발생 시 표준 텍스트 영역으로 대체
    return st.text_area(
        label="콘티 본문",
        value=value,
        height=height,
        key=f"{key}_fallback" if key else None,
        label_visibility="collapsed"
    )
