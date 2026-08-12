"""Single professional launch point for the separately deployed ERP apps."""

import streamlit as st


st.set_page_config(page_title="AutonoVerse ERP", page_icon="◈", layout="wide")

st.markdown("""
<style>
  #MainMenu, header, footer, [data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"]{display:none!important}
  [data-testid="stAppViewContainer"]{background:radial-gradient(circle at 80% 0,#15345a 0,#091525 38%,#060d18 100%);color:#edf5ff}
  [data-testid="stMainBlockContainer"]{max-width:1180px;padding-top:2rem;padding-bottom:3rem}
  .hero{padding:30px 34px;border:1px solid #284362;border-radius:22px;background:linear-gradient(135deg,#11253ddd,#0b1728dd);box-shadow:0 24px 70px #0006;margin-bottom:24px}
  .eyebrow{color:#54c7ff;font-weight:750;font-size:12px;letter-spacing:.16em;text-transform:uppercase}.hero h1{font-size:40px;line-height:1.1;margin:10px 0 8px;color:#fff}.hero p{color:#a9bbd1;font-size:16px;margin:0;max-width:720px}
  .online{display:inline-flex;align-items:center;gap:8px;margin-top:20px;padding:7px 11px;border-radius:30px;background:#123124;color:#72e5ad;font-size:12px}.dot{width:8px;height:8px;border-radius:50%;background:#40dc91;box-shadow:0 0 12px #40dc91}
  .section{font-size:13px;color:#89a2bf;letter-spacing:.1em;text-transform:uppercase;font-weight:750;margin:25px 2px 12px}
  .module{padding:20px 20px 12px;border:1px solid #263d58;border-radius:16px 16px 0 0;background:#0e1c2dcc;min-height:144px}.module .icon{display:grid;place-items:center;width:42px;height:42px;border-radius:12px;background:linear-gradient(135deg,#225986,#173b62);font-size:21px;margin-bottom:15px}.module strong{display:block;font-size:18px;margin-bottom:7px;color:#edf5ff}.module span{color:#93a9c2;line-height:1.5}.agent{background:linear-gradient(110deg,#10283a,#13233f);border-color:#315678}.agent .icon{background:linear-gradient(135deg,#4d55d9,#2b91bd)}
  [data-testid="stLinkButton"]{margin-top:-1px;margin-bottom:16px}[data-testid="stLinkButton"] a{border-radius:0 0 14px 14px!important;border-color:#31577a!important;background:#142b43!important;color:#dff5ff!important;font-weight:700!important;min-height:44px}[data-testid="stLinkButton"] a:hover{background:#1b4264!important;border-color:#45aee8!important;color:#fff!important}
  .note{margin-top:10px;color:#7188a2;font-size:12px;text-align:center}
</style>
<section class="hero"><div class="eyebrow">Unified Operations Cloud</div><h1>AutonoVerse ERP</h1><p>One operations launchpad for procurement, manufacturing, order fulfilment, exceptions, and your conversational ERP agent.</p><div class="online"><i class="dot"></i>All five cloud systems online</div></section>
<div class="section">Core business operations</div>
""", unsafe_allow_html=True)

modules = [
    ("📋", "Source to Pay", "Purchase requests, consolidation, RFx, vendors, contracts, and purchase orders.", "https://navdeep-erp-s2p.streamlit.app/"),
    ("⚙️", "Manufacturing", "Goods receipt, quality, BOM planning, production, inventory, transfers, and traceability.", "https://navdeep-erp-mfg.streamlit.app/"),
    ("🛒", "Order to Cash", "Customers, quotations, orders, fulfilment, billing, collections, returns, and accounting.", "https://erp-poc-app-joy6baexqfpzvxfffy4bc2.streamlit.app/"),
    ("◉", "Exception Control Tower", "A prioritized worklist of live operational exceptions across the platform.", "https://navdeep-erp-control-tower.streamlit.app/"),
]

for start in range(0, len(modules), 2):
    for column, (icon, title, description, url) in zip(st.columns(2), modules[start:start + 2]):
        with column:
            st.markdown(f'<div class="module"><div class="icon">{icon}</div><strong>{title}</strong><span>{description}</span></div>', unsafe_allow_html=True)
            st.link_button(f"Open {title}  ↗", url, use_container_width=True)

st.markdown('<div class="section">Conversational workspace</div><div class="module agent"><div class="icon">✦</div><strong>Agent Console</strong><span>Ask operational questions and navigate the ERP through the conversational interface.</span></div>', unsafe_allow_html=True)
st.link_button("Open Agent Console  ↗", "https://navdeep-erp-agent.streamlit.app/", use_container_width=True)
st.markdown('<div class="note">All workspaces open securely in a new browser tab, keeping this launchpad available.</div>', unsafe_allow_html=True)
