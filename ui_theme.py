"""Shared visual theme for the ERP Streamlit surfaces."""

import streamlit as st


THEME_CSS = """
<style>
:root {
  --erp-navy: #102a43;
  --erp-blue: #2563eb;
  --erp-blue-dark: #1d4ed8;
  --erp-cyan: #06b6d4;
  --erp-ink: #172033;
  --erp-muted: #64748b;
  --erp-line: #dbe4ee;
  --erp-surface: #ffffff;
  --erp-canvas: #f3f7fb;
}

.stApp {
  background:
    radial-gradient(circle at 92% 0%, rgba(37, 99, 235, .08), transparent 24rem),
    var(--erp-canvas);
}

[data-testid="stHeader"] { background: rgba(243, 247, 251, .82); }
[data-testid="stAppViewContainer"] > .main .block-container {
  max-width: 1500px;
  padding-top: 2rem;
  padding-bottom: 4rem;
}

[data-testid="stSidebar"] {
  background: linear-gradient(180deg, #102a43 0%, #163d61 58%, #102a43 100%);
  border-right: 0;
  box-shadow: 12px 0 32px rgba(15, 42, 67, .12);
}
[data-testid="stSidebar"] * { color: #e8f1f8 !important; }
[data-testid="stSidebar"] hr { border-color: rgba(255,255,255,.14) !important; }
[data-testid="stSidebar"] [role="radiogroup"] label {
  padding: .62rem .7rem;
  border-radius: .65rem;
  margin: .1rem 0;
  transition: background .16s ease, transform .16s ease;
}
[data-testid="stSidebar"] [role="radiogroup"] label:hover {
  background: rgba(255,255,255,.1);
  transform: translateX(2px);
}
[data-testid="stSidebar"] [data-baseweb="radio"] > div:first-child {
  border-color: #67e8f9 !important;
}

h1, h2, h3 { color: var(--erp-navy) !important; letter-spacing: -.02em; }
h4, h5, h6 { color: #243b53 !important; }
p, label, [data-testid="stCaptionContainer"] { color: var(--erp-muted); }

div[data-testid="stMetric"] {
  background: linear-gradient(145deg, #ffffff 0%, #f8fbff 100%);
  border: 1px solid var(--erp-line);
  border-radius: 14px;
  padding: 1rem 1.1rem;
  box-shadow: 0 8px 24px rgba(15, 42, 67, .055);
}
div[data-testid="stMetricLabel"] { color: var(--erp-muted) !important; }
div[data-testid="stMetricValue"] { color: var(--erp-navy) !important; font-weight: 750; }

.stButton > button, .stDownloadButton > button {
  border-radius: 9px !important;
  min-height: 2.55rem;
  font-weight: 650 !important;
  transition: transform .15s ease, box-shadow .15s ease, border-color .15s ease;
}
.stButton > button[kind="primary"] {
  color: white !important;
  border: 0 !important;
  background: linear-gradient(135deg, var(--erp-blue), #0e7490) !important;
  box-shadow: 0 6px 16px rgba(37, 99, 235, .2);
}
.stButton > button:hover, .stDownloadButton > button:hover {
  transform: translateY(-1px);
  box-shadow: 0 8px 20px rgba(15, 42, 67, .15);
}
.stButton > button[kind="secondary"], .stDownloadButton > button {
  background: white !important;
  color: var(--erp-blue) !important;
  border: 1px solid #b9cbe0 !important;
}

[data-baseweb="input"] > div, [data-baseweb="select"] > div,
[data-baseweb="textarea"] > div {
  background: white !important;
  border-color: #c8d5e3 !important;
  border-radius: 9px !important;
}
input, textarea { color: var(--erp-ink) !important; }

div[data-baseweb="tab-list"] {
  gap: .35rem;
  background: #e8eff7;
  border-radius: 11px;
  padding: .28rem;
}
button[data-baseweb="tab"] { border-radius: 8px; padding: .35rem .9rem; }
button[data-baseweb="tab"][aria-selected="true"] {
  background: white;
  color: var(--erp-blue) !important;
  box-shadow: 0 2px 8px rgba(15, 42, 67, .09);
}

[data-testid="stDataFrame"], details {
  background: white;
  border: 1px solid var(--erp-line) !important;
  border-radius: 12px !important;
  overflow: hidden;
  box-shadow: 0 5px 18px rgba(15, 42, 67, .04);
}
[data-testid="stAlert"] { border-radius: 10px; }
hr { border-color: var(--erp-line) !important; }

@media (max-width: 768px) {
  [data-testid="stAppViewContainer"] > .main .block-container { padding-top: 1rem; }
  div[data-testid="stMetric"] { padding: .8rem; }
}
</style>
"""


def apply_theme():
    """Inject the shared app styling after ``st.set_page_config``."""
    st.markdown(THEME_CSS, unsafe_allow_html=True)
