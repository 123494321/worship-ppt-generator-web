import os
import streamlit as st
import streamlit.components.v1 as components

_COMPONENT_DIR = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "components", "pin_gate"))
_pin_gate_func = components.declare_component("pin_gate", path=_COMPONENT_DIR)

def render_pin_gate(key="pin_gate_auth", allowed_pins=None):
    """
    Renders an interactive, auto-submitting 4-digit PIN authentication gate.
    Supports general user (0000) and admin PIN (default: 7777).
    """
    if allowed_pins is None:
        allowed_pins = ["0000", "7777"]
    return _pin_gate_func(key=key, allowed_pins=allowed_pins, default=None)
