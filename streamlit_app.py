"""Unified entry point for the complete ERP suite."""

import streamlit as st

from ui_theme import apply_theme


st.set_page_config(
    page_title="Northstar ERP",
    page_icon="◆",
    layout="wide",
    initial_sidebar_state="expanded",
)
apply_theme()

with st.sidebar:
    st.markdown(
        """
        <div class="erp-brand">
          <div class="erp-brand-mark">N</div>
          <div><strong>Northstar ERP</strong><small>OPERATIONS CLOUD</small></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

navigation = st.navigation(
    {
        "Workspace": [
            st.Page("home.py", title="Command Center", icon=":material/dashboard:", default=True),
            st.Page("agent_console.py", title="AI Operations", icon=":material/auto_awesome:"),
        ],
        "Business modules": [
            st.Page("erp_ui.py", title="Source to Pay", icon=":material/shopping_cart:"),
            st.Page("mfg_ui.py", title="Manufacturing", icon=":material/precision_manufacturing:"),
            st.Page("o2c_ui.py", title="Order to Cash", icon=":material/payments:"),
        ],
    },
    position="sidebar",
)
navigation.run()
