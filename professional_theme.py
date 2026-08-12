"""Shared visual system for every AutonoVerse ERP Streamlit surface."""

import streamlit as st


def apply_theme(area="Enterprise Operations"):
    st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Manrope:wght@600;700;800&display=swap');
    :root{{--av-ink:#132238;--av-muted:#64748b;--av-line:#dce5ef;--av-blue:#155eef;--av-navy:#09182c;--av-bg:#f4f7fb}}
    html,body,[class*="css"]{{font-family:Inter,system-ui,sans-serif}}
    .stApp{{background:linear-gradient(180deg,#f8fafc 0,#f3f6fa 100%);color:var(--av-ink)}}
    [data-testid="stHeader"]{{background:rgba(248,250,252,.88);backdrop-filter:blur(12px);border-bottom:1px solid #e6edf5}}
    [data-testid="stMainBlockContainer"]{{padding-top:2rem;max-width:1480px}}
    h1,h2,h3,h4,h5{{font-family:Manrope,Inter,sans-serif!important;color:#10213a!important;letter-spacing:-.025em}}
    h1{{font-weight:800!important}} h2,h3{{font-weight:700!important}}
    p,label,[data-testid="stCaptionContainer"]{{color:#53657b}}
    [data-testid="stSidebar"]{{background:linear-gradient(180deg,#08192d 0%,#0d2340 65%,#102b4a 100%);border-right:1px solid #183a60;box-shadow:8px 0 28px rgba(13,32,55,.08)}}
    [data-testid="stSidebar"] *{{color:#dbe8f7!important}}
    [data-testid="stSidebar"] h3{{color:#fff!important;font-size:1.08rem!important;letter-spacing:-.01em}}
    [data-testid="stSidebar"] [data-testid="stCaptionContainer"]{{color:#86a4c5!important;text-transform:uppercase;letter-spacing:.12em;font-size:.67rem}}
    [data-testid="stSidebar"] hr{{border-color:#264667!important}}
    [data-testid="stSidebar"] [role="radiogroup"] label{{border-radius:9px;padding:8px 9px;margin:2px 0;transition:.15s}}
    [data-testid="stSidebar"] [role="radiogroup"] label:hover{{background:#183a5d}}
    [data-testid="stSidebar"] [data-testid="stSidebarNav"] a{{border-radius:9px;margin:2px 8px;padding:9px 12px}}
    [data-testid="stSidebar"] [data-testid="stSidebarNav"] a:hover{{background:#173a5e}}
    [data-testid="stSidebar"] [aria-current="page"]{{background:linear-gradient(90deg,#155eef,#2476ff)!important;color:#fff!important;box-shadow:0 6px 18px #07182b88}}
    div[data-testid="stMetric"]{{background:#fff;border:1px solid var(--av-line);border-radius:14px;padding:17px 19px;box-shadow:0 5px 18px rgba(15,35,60,.055)}}
    div[data-testid="stMetricLabel"]{{color:#6a7b91!important;font-size:.72rem!important;text-transform:uppercase;letter-spacing:.07em;font-weight:700}}
    div[data-testid="stMetricValue"]{{color:#10213a!important;font-family:Manrope,sans-serif;font-weight:800!important}}
    .stButton>button,.stDownloadButton>button,[data-testid="stLinkButton"] a{{border-radius:9px!important;font-weight:700!important;min-height:40px;border:1px solid #155eef!important;background:#155eef!important;color:#fff!important;box-shadow:0 4px 12px rgba(21,94,239,.18);transition:.16s}}
    .stButton>button:hover,.stDownloadButton>button:hover,[data-testid="stLinkButton"] a:hover{{background:#0d4fd8!important;border-color:#0d4fd8!important;transform:translateY(-1px);box-shadow:0 7px 18px rgba(21,94,239,.24)}}
    .stButton>button[kind="secondary"]{{background:#fff!important;color:#243b56!important;border-color:#cad6e4!important;box-shadow:0 2px 7px rgba(15,35,60,.05)}}
    input,textarea,[data-baseweb="select"]>div{{background:#fff!important;border-color:#cbd8e6!important;border-radius:9px!important;color:#1c2e44!important}}
    input:focus,textarea:focus{{border-color:#3b82f6!important;box-shadow:0 0 0 3px rgba(59,130,246,.12)!important}}
    div[data-baseweb="tab-list"]{{gap:5px;background:#eaf0f6;border:1px solid #d9e3ed;border-radius:11px;padding:5px}}
    button[data-baseweb="tab"]{{border-radius:8px;padding:8px 14px}}
    button[data-baseweb="tab"][aria-selected="true"]{{background:#fff;color:#174a80!important;box-shadow:0 2px 8px rgba(20,45,75,.1)}}
    [data-testid="stDataFrame"]{{border:1px solid #dbe5ef;border-radius:12px;overflow:hidden;box-shadow:0 4px 14px rgba(15,35,60,.045)}}
    details{{background:#fff;border:1px solid #dce5ef!important;border-radius:11px!important;box-shadow:0 2px 9px rgba(15,35,60,.035)}}
    [data-testid="stAlert"]{{border-radius:11px;border-width:1px;box-shadow:0 2px 8px rgba(15,35,60,.04)}}
    hr{{border-color:#dfe7f0!important;margin:1.25rem 0!important}}
    .av-suite-tag{{display:flex;align-items:center;gap:9px;padding:9px 12px;margin:0 0 13px;border:1px solid #dbe5ef;border-radius:10px;background:#fff;color:#51647a;font-size:.72rem;font-weight:700;letter-spacing:.08em;text-transform:uppercase}}
    .av-suite-tag i{{width:8px;height:8px;border-radius:50%;background:#19b979;box-shadow:0 0 0 4px #dff8ee}}
    </style>
    <div class="av-suite-tag"><i></i> AutonoVerse ERP &nbsp;·&nbsp; {area}</div>
    """, unsafe_allow_html=True)
