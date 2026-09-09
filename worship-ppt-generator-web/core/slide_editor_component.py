import os
import sys
import json
import logging
import streamlit as st
from streamlit import config
from streamlit.runtime import get_instance
import streamlit.components.v1.custom_component as cc_module
from streamlit.components.v1.custom_component import CustomComponent

# ---------------------------------------------------------------------------
# Robust CustomComponent.create_instance patch
# ---------------------------------------------------------------------------
# Streamlit's default create_instance unconditionally executes `import pyarrow`.
# In frozen Windows environments (PyInstaller), pyarrow's C++ DLLs (arrow.dll,
# aws-cpp, etc.) may fail to load, throwing ImportError and triggering fallback.
# Since slide_editor only exchanges JSON strings/integers and never passes
# Arrow dataframes, we provide a safe create_instance that bypasses pyarrow
# unless a dataframe is actually provided.
# ---------------------------------------------------------------------------
def _safe_create_instance(
    self,
    *args,
    default=None,
    key=None,
    on_change=None,
    tab_index=None,
    **kwargs,
):
    if len(args) > 0:
        raise cc_module.MarshallComponentException(f"Argument '{args[0]}' needs a label")

    if tab_index is not None and not (
        isinstance(tab_index, int)
        and not isinstance(tab_index, bool)
        and tab_index >= -1
    ):
        raise cc_module.StreamlitAPIException(
            "tab_index must be None, -1, or a non-negative integer."
        )

    cc_module.check_cache_replay_rules()
    all_args = dict(kwargs, default=default, key=key)

    json_args = {}
    special_args = []
    for arg_name, arg_val in all_args.items():
        if cc_module.is_bytes_like(arg_val):
            bytes_arg = cc_module.SpecialArg()
            bytes_arg.key = arg_name
            bytes_arg.bytes = cc_module.to_bytes(arg_val)
            special_args.append(bytes_arg)
        elif cc_module.is_dataframe_like(arg_val):
            try:
                from streamlit.components.v1 import component_arrow
                dataframe_arg = cc_module.SpecialArg()
                dataframe_arg.key = arg_name
                component_arrow.marshall(dataframe_arg.arrow_dataframe.data, arg_val)
                special_args.append(dataframe_arg)
            except Exception as ex:
                raise cc_module.MarshallComponentException("DataFrame marshalling failed", ex)
        else:
            json_args[arg_name] = arg_val

    try:
        serialized_json_args = cc_module.json.dumps(json_args)
    except Exception as ex:
        raise cc_module.MarshallComponentException(
            "Could not convert component args to JSON", ex
        )

    def marshall_component(dg, element):
        element.component_instance.component_name = self.name
        element.component_instance.form_id = cc_module.current_form_id(dg)
        if self.url is not None:
            element.component_instance.url = self.url
        if tab_index is not None:
            element.component_instance.tab_index = tab_index

        element.component_instance.json_args = serialized_json_args
        element.component_instance.special_args.extend(special_args)

        computed_id = cc_module.compute_and_register_element_id(
            "component_instance",
            user_key=key,
            key_as_main_identity={"name", "url"},
            dg=dg,
            name=self.name,
            url=self.url,
            json_args=serialized_json_args,
            special_args=special_args,
        )
        element.component_instance.id = computed_id

        def deserialize_component(ui_value):
            return ui_value

        component_state = cc_module.register_widget(
            element.component_instance.id,
            deserializer=deserialize_component,
            serializer=lambda x: x,
            ctx=cc_module.get_script_run_ctx(),
            on_change_handler=on_change,
            value_type="json_value",
        )
        widget_value = component_state.value

        if widget_value is None:
            widget_value = default
        elif isinstance(widget_value, cc_module.ArrowTableProto):
            try:
                from streamlit.components.v1 import component_arrow
                widget_value = component_arrow.arrow_proto_to_dataframe(widget_value)
            except Exception:
                pass
        return widget_value

    dg = cc_module.get_dg_singleton_instance().main_dg
    element = cc_module.Element()
    return_value = marshall_component(dg, element)
    dg._enqueue("component_instance", element.component_instance)
    return return_value

cc_module.CustomComponent.create_instance = _safe_create_instance

_COMPONENT_DIR = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "components", "slide_editor"))
_slide_ide_editor = None

def _get_or_register_editor_component():
    """
    Directly instantiates and registers the CustomComponent in Streamlit's ComponentRegistry.
    This completely bypasses inspect.getmodule() which fails in PyInstaller frozen environments.
    """
    global _slide_ide_editor
    if _slide_ide_editor is None:
        try:
            module_name = "core.slide_editor_component"
            name = "slide_ide_editor"
            component_name = f"{module_name}.{name}"

            url = None
            if component_base_path := config.get_option("server.customComponentBaseUrlPath"):
                url = f"{component_base_path}/{component_name}/"

            _slide_ide_editor = CustomComponent(
                name=component_name,
                path=str(_COMPONENT_DIR),
                url=url,
                module_name=module_name
            )
        except Exception as e:
            logging.error(f"[slide_editor] Direct CustomComponent creation failed: {e}")
            _slide_ide_editor = None

    if _slide_ide_editor is not None:
        try:
            inst = get_instance()
            if inst and hasattr(inst, "component_registry") and inst.component_registry is not None:
                inst.component_registry.register_component(_slide_ide_editor)
        except Exception as e:
            logging.debug(f"[slide_editor] Runtime registration warning: {e}")

    return _slide_ide_editor

def render_slide_ide_editor(value="", height=580, key=None):
    """
    Renders an IDE-style code editor with dynamic slide numbering gutter.
    Guaranteed native rendering with 1:1 synchronized scrolling.
    Falls back gracefully to standard st.text_area if custom component is unavailable.
    """
    editor = _get_or_register_editor_component()
    if editor is not None:
        try:
            res = editor(value=value, height=height, key=key, default=value)
            return res if res is not None else value
        except Exception as e:
            logging.error(f"[slide_editor] Custom component rendering exception: {e}", exc_info=True)

    # 2중 안전망: 커스텀 컴포넌트 렌더링 예외 발생 시 표준 텍스트 영역으로 매끄럽게 대체
    return st.text_area(
        label="슬라이드 기획안 편집 (텍스트 모드)",
        value=value,
        height=height,
        key=f"{key}_fallback" if key else None,
        label_visibility="collapsed"
    )
