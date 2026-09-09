import os
import streamlit as st
import streamlit.components.v1 as components

_COMPONENT_DIR = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "components", "pin_gate"))
_pin_gate_func = components.declare_component("pin_gate", path=_COMPONENT_DIR)

def render_pin_gate(key="pin_gate_auth"):
    """
    Renders an interactive, auto-submitting 4-digit PIN authentication gate.
    When the user enters 0000, it immediately notifies Python without pressing Enter or a button.
    """
    return _pin_gate_func(key=key, default=None)
