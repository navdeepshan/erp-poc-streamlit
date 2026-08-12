"""Single professional launch point for the separately deployed ERP apps."""

import streamlit as st


st.set_page_config(page_title="AutonoVerse ERP", page_icon="◈", layout="wide")

st.markdown("""
<style>
  #MainMenu, header, footer, [data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"]{display:none!important}
  [data-testid="stAppViewContainer"]{background:radial-gradient(circle at 80% 0,#15345a 0,#091525 38%,#060d18 100%);color:#edf5ff}
  [data-testid="stMainBlockContainer"]{max-width:1180px;padding-top:2rem;padding-bottom:3rem}
  .hero{padding:30px 34px;border:1px solid #284362;border-radius:22px;background:linear-gradient(135deg,#11253ddd,#0b1728dd);box-shadow:0 24px 70px #0006;margin-bottom:24px}
  .eyebrow{color:#54c7ff;font-weight:750;font-size:12px;letter-spacing:.16em;text-transform:uppercase}
  .hero h1{font-size:40px;line-height:1.1;margin:10px 0 8px;color:#fff}.hero p{color:#a9bbd1;font-size:16px;margin:0;max-width:720px}
  .online{display:inline-flex;align-items:center;gap:8px;margin-top:20px;padding:7px 11px;border-radius:30px;background:#123124;color:#72e5ad;font-size:12px}.dot{width:8px;height:8px;border-radius:50%;background:#40dc91;box-shadow:0 0 12px #40dc91}
  .section{font-size:13px;color:#89a2bf;letter-spacing:.1em;text-transform:uppercase;font-weight:750;margin:25px 2px 12px}
  .grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:15px}
  a.card{display:block;text-decoration:none!important;padding:22px;border:1px solid #263d58;border-radius:16px;background:#0e1c2dcc;color:#edf5ff!important;transition:.18s;min-height:150px}
  a.card:hover{transform:translateY(-3px);border-color:#3d8ec4;background:#12263d;box-shadow:0 15px 34px #0005}
  .icon{display:grid;place-items:center;width:42px;height:42px;border-radius:12px;background:linear-gradient(135deg,#225986,#173b62);font-size:21px;margin-bottom:17px}.card strong{display:block;font-size:18px;margin-bottom:7px}.card span{color:#93a9c2;line-height:1.5}.arrow{float:right;color:#55c7ff;font-size:20px}
  .agent{grid-column:span 2;background:linear-gradient(110deg,#10283a,#13233f)!important;border-color:#315678!important}.agent .icon{background:linear-gradient(135deg,#4d55d9,#2b91bd)}
  .note{margin-top:23px;color:#7188a2;font-size:12px;text-align:center}
  @media(max-width:700px){.grid{grid-template-columns:1fr}.agent{grid-column:auto}.hero{padding:24px}.hero h1{font-size:31px}}
</style>

<section class="hero">
  <div class="eyebrow">Unified Operations Cloud</div>
  <h1>AutonoVerse ERP</h1>
  <p>One secure operations launchpad for procurement, manufacturing, order fulfilment, exceptions, and your conversational ERP agent.</p>
  <div class="online"><i class="dot"></i>All five cloud systems online</div>
</section>

<div class="section">Core business operations</div>
<div class="grid">
  <a class="card" href="https://navdeep-erp-s2p.streamlit.app/" target="_self"><span class="arrow">↗</span><div class="icon">📋</div><strong>Source to Pay</strong><span>Purchase requests, consolidation, RFx, vendors, contracts, and purchase orders.</span></a>
  <a class="card" href="https://navdeep-erp-mfg.streamlit.app/" target="_self"><span class="arrow">↗</span><div class="icon">⚙️</div><strong>Manufacturing</strong><span>Goods receipt, quality, BOM planning, production, inventory, transfers, and traceability.</span></a>
  <a class="card" href="https://erp-poc-app-joy6baexqfpzvxfffy4bc2.streamlit.app/" target="_self"><span class="arrow">↗</span><div class="icon">🛒</div><strong>Order to Cash</strong><span>Customers, quotations, orders, fulfilment, billing, collections, returns, and accounting.</span></a>
  <a class="card" href="https://navdeep-erp-control-tower.streamlit.app/" target="_self"><span class="arrow">↗</span><div class="icon">◉</div><strong>Exception Control Tower</strong><span>A single prioritized worklist of live operational exceptions across the platform.</span></a>
  <a class="card agent" href="https://navdeep-erp-agent.streamlit.app/" target="_self"><span class="arrow">↗</span><div class="icon">✦</div><strong>Agent Console</strong><span>Ask operational questions and navigate the ERP through the conversational interface.</span></a>
</div>
<div class="note">Select a workspace to continue · Use your browser Back button to return to this launchpad</div>
""", unsafe_allow_html=True)
