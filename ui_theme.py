"""Shared visual theme and embedded chatbot for the ERP suite."""

import streamlit as st


THEME_CSS = """
<style>
:root {
  --erp-navy:#101828; --erp-sidebar:#0f172a; --erp-blue:#2563eb;
  --erp-blue-hover:#1d4ed8; --erp-ink:#101828; --erp-muted:#667085;
  --erp-line:#e4e7ec; --erp-canvas:#f8fafc; --erp-surface:#ffffff;
}
html,body,[class*="css"] { font-family:Inter,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif; }
.stApp { background:var(--erp-canvas) !important; color:var(--erp-ink); }
[data-testid="stHeader"] { background:rgba(248,250,252,.92) !important; border-bottom:1px solid rgba(228,231,236,.65); }
[data-testid="stAppViewContainer"] > .main .block-container { max-width:1440px; padding:2.25rem 2.5rem 4rem; }
[data-testid="stSidebar"] { background:var(--erp-sidebar) !important; border-right:1px solid #1e293b !important; box-shadow:none !important; }
[data-testid="stSidebarContent"] { padding:.65rem .7rem 1rem; }
[data-testid="stSidebar"] * { color:#cbd5e1 !important; }
[data-testid="stSidebar"] h3 { color:#fff !important; font-size:.9rem !important; font-weight:650 !important; letter-spacing:0 !important; }
[data-testid="stSidebar"] [data-testid="stCaptionContainer"] { color:#94a3b8 !important; font-size:.71rem; text-transform:uppercase; letter-spacing:.09em; }
[data-testid="stSidebar"] hr { border-color:#25334a !important; margin:.75rem 0 !important; }
[data-testid="stSidebar"] [role="radiogroup"] { gap:.15rem; }
[data-testid="stSidebar"] [role="radiogroup"] label { padding:.58rem .66rem !important; border-radius:7px !important; margin:0 !important; transition:background .14s ease; }
[data-testid="stSidebar"] [role="radiogroup"] label:hover { background:#1e293b !important; transform:none !important; }
[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) { background:#243451 !important; box-shadow:inset 3px 0 0 #60a5fa; }
[data-testid="stSidebar"] [role="radiogroup"] label p { color:#e2e8f0 !important; font-size:.84rem !important; }
[data-testid="stSidebar"] [data-baseweb="radio"] > div:first-child { display:none !important; }
[data-testid="stSidebar"] [data-testid="stSidebarNav"] a { border-radius:7px; margin:2px 0; }
[data-testid="stSidebar"] [data-testid="stSidebarNav"] a:hover { background:#1e293b; }
[data-testid="stSidebar"] [data-testid="stSidebarNav"] a[aria-current="page"] { background:#243451; }
h1,h2,h3 { color:var(--erp-navy) !important; letter-spacing:-.025em; font-weight:680 !important; }
h1 { font-size:2rem !important; line-height:1.2 !important; }
h2 { font-size:1.45rem !important; line-height:1.3 !important; }
h3 { font-size:1.12rem !important; }
h4,h5,h6 { color:#344054 !important; font-weight:650 !important; }
p,label,[data-testid="stCaptionContainer"] { color:var(--erp-muted) !important; line-height:1.55; }
div[data-testid="stMetric"] { background:var(--erp-surface) !important; border:1px solid var(--erp-line) !important; border-radius:10px !important; padding:1rem 1.1rem !important; box-shadow:0 1px 2px rgba(16,24,40,.04) !important; }
div[data-testid="stMetricLabel"] { color:#667085 !important; font-size:.74rem !important; font-weight:550; }
div[data-testid="stMetricValue"] { color:var(--erp-navy) !important; font-size:1.55rem !important; font-weight:700 !important; letter-spacing:-.025em; }
.stButton > button,.stDownloadButton > button { border-radius:7px !important; min-height:2.45rem; padding:.45rem .9rem !important; font-size:.84rem !important; font-weight:600 !important; box-shadow:none !important; transition:background .14s ease,border-color .14s ease !important; }
.stButton > button[kind="primary"] { color:#fff !important; border:1px solid var(--erp-blue) !important; background:var(--erp-blue) !important; }
.stButton > button[kind="primary"] * { color:#fff !important; }
.stButton > button[kind="primary"]:hover { background:var(--erp-blue-hover) !important; border-color:var(--erp-blue-hover) !important; transform:none !important; }
.stButton > button[kind="secondary"],.stDownloadButton > button { background:#fff !important; color:#344054 !important; border:1px solid #d0d5dd !important; }
.stButton > button[kind="secondary"]:hover,.stDownloadButton > button:hover { background:#f9fafb !important; border-color:#98a2b3 !important; transform:none !important; }
[data-baseweb="input"] > div,[data-baseweb="select"] > div,[data-baseweb="textarea"] > div { background:#fff !important; border:1px solid #d0d5dd !important; border-radius:7px !important; box-shadow:0 1px 2px rgba(16,24,40,.04) !important; }
[data-baseweb="input"] > div:focus-within,[data-baseweb="select"] > div:focus-within,[data-baseweb="textarea"] > div:focus-within { border-color:#84adff !important; box-shadow:0 0 0 3px rgba(37,99,235,.1) !important; }
input,textarea { color:var(--erp-ink) !important; }
div[data-baseweb="tab-list"] { gap:.2rem; background:transparent !important; border:0 !important; border-bottom:1px solid var(--erp-line) !important; border-radius:0 !important; padding:0 !important; }
button[data-baseweb="tab"] { border-radius:0 !important; padding:.6rem .85rem !important; border-bottom:2px solid transparent !important; }
button[data-baseweb="tab"][aria-selected="true"] { background:transparent !important; color:var(--erp-blue) !important; border-bottom-color:var(--erp-blue) !important; box-shadow:none !important; }
[data-testid="stDataFrame"] { background:#fff !important; border:1px solid var(--erp-line) !important; border-radius:9px !important; overflow:hidden; box-shadow:none !important; }
details { background:#fff !important; border:1px solid var(--erp-line) !important; border-radius:8px !important; margin-bottom:.5rem !important; box-shadow:none !important; }
summary { padding:.75rem .9rem !important; font-size:.88rem; }
[data-testid="stAlert"] { border-radius:8px !important; border-width:1px !important; box-shadow:none !important; }
hr { border-color:var(--erp-line) !important; margin:1.25rem 0 !important; }
.erp-brand { display:flex; align-items:center; gap:.75rem; padding:.35rem .15rem 1rem; }
.erp-brand-mark { width:2rem; height:2rem; display:grid; place-items:center; border-radius:7px; background:#2563eb; color:#fff!important; font-weight:750; }
.erp-brand strong { display:block; color:#fff!important; }
.erp-brand small { display:block; color:#98a2b3!important; font-size:.58rem; letter-spacing:.12em; margin-top:.12rem; }
.erp-eyebrow { color:#315bea!important; font-size:.67rem; letter-spacing:.13em; font-weight:750; margin-bottom:.5rem; }
[data-testid="stSidebarNav"] span { font-size:.88rem; }
[data-testid="stVerticalBlockBorderWrapper"] { background:#fff; border-color:var(--erp-line) !important; border-radius:10px !important; box-shadow:0 1px 2px rgba(16,24,40,.035); }
@media (max-width:768px) { [data-testid="stAppViewContainer"] > .main .block-container { padding:1.15rem 1rem 3rem; } h1{font-size:1.65rem!important;} div[data-testid="stMetric"]{padding:.8rem!important;} }
</style>
"""


def apply_theme():
    st.markdown(THEME_CSS, unsafe_allow_html=True)
    _install_chatbot()


def _install_chatbot():
    script = """
    <script>
    (() => {
      const install = () => {
        try {
          const doc = window.top.document;
          if (doc.getElementById('erp-chat-launcher')) return;
          const launcher = doc.createElement('button');
          launcher.id = 'erp-chat-launcher';
          launcher.textContent = window.top.innerWidth < 768 ? '✦' : '✦  Ask Northstar';
          launcher.setAttribute('aria-label', 'Open ERP assistant');
          launcher.style.cssText = `position:fixed;right:${window.top.innerWidth < 768 ? '16px' : '24px'};bottom:${window.top.innerWidth < 768 ? '16px' : '24px'};z-index:2147483647;height:52px;min-width:52px;padding:0 ${window.top.innerWidth < 768 ? '0' : '22px'};border:0;border-radius:26px;background:#2563eb;color:#fff;font:600 15px system-ui;box-shadow:0 10px 28px rgba(17,24,39,.24);cursor:pointer`;
          launcher.onclick = () => {
            let panel = doc.getElementById('erp-chat-panel');
            if (panel) { panel.style.display = panel.style.display === 'none' ? 'flex' : 'none'; return; }
            panel = doc.createElement('section');
            panel.id = 'erp-chat-panel';
            panel.setAttribute('aria-label', 'Northstar ERP assistant');
            panel.style.cssText = 'position:fixed;right:24px;bottom:88px;z-index:2147483646;width:min(430px,calc(100vw - 32px));height:min(680px,calc(100vh - 120px));display:flex;flex-direction:column;overflow:hidden;background:#fff;border:1px solid #d9e2ec;border-radius:18px;box-shadow:0 24px 70px rgba(17,24,39,.3)';
            const header = doc.createElement('div');
            header.style.cssText = 'height:56px;flex:0 0 56px;display:flex;align-items:center;justify-content:space-between;padding:0 14px 0 18px;background:#111827;color:#fff;font:600 15px system-ui';
            header.innerHTML = '<span>✦ &nbsp;Northstar Assistant</span>';
            const close = doc.createElement('button');
            close.textContent = '×'; close.setAttribute('aria-label','Close ERP assistant');
            close.style.cssText = 'border:0;background:transparent;color:#fff;font-size:26px;cursor:pointer;line-height:1';
            close.onclick = () => { panel.style.display = 'none'; };
            header.appendChild(close);
            const frame = doc.createElement('iframe');
            frame.title = 'Northstar ERP Chat';
            frame.src = window.top.location.origin + '/agent_console?embedded_chat=1';
            frame.style.cssText = 'width:100%;height:100%;border:0;background:#fff';
            panel.append(header, frame); doc.body.appendChild(panel);
          };
          doc.body.appendChild(launcher);
        } catch (_) {}
      };
      install(); const timer = setInterval(install,500); setTimeout(() => clearInterval(timer),30000);
    })();
    </script>
    """
    st.components.v1.html(script, height=0)


def embed_html(content, *, height, scrolling=False):
    st.components.v1.html(content, height=height, scrolling=scrolling)
