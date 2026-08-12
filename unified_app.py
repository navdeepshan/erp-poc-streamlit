"""Single Streamlit Cloud entry point; domain apps remain separate modules."""

import streamlit as st


pages = {
    "Core Operations": [
        st.Page("erp_ui.py", title="Source to Pay", icon="📋"),
        st.Page("mfg_ui.py", title="Manufacturing", icon="⚙️"),
        st.Page("o2c_ui.py", title="Order to Cash", icon="🛒"),
    ],
    "Intelligence": [
        st.Page("control_tower_ui.py", title="Control Tower", icon="🎯"),
        st.Page("agent_console.py", title="Agent Console", icon="🤖"),
    ],
}

navigation = st.navigation(pages, position="sidebar", expanded=True)
navigation.run()
