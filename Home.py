import streamlit as st

st.set_page_config(
    page_title="Transport Analytics",
    page_icon="🚛",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("🚛 Transport Analytics Platform")
st.markdown("### Select a module from the sidebar to begin")
st.markdown("---")

c1, c2, c3, c4, c5 = st.columns(5)

with c1:
    st.markdown("""
    <div style='background:#1F4E79;padding:20px;border-radius:12px;text-align:center;min-height:220px'>
        <div style='font-size:36px'>📥📤</div>
        <h4 style='color:white;margin:8px 0'>TAT Analysis</h4>
        <p style='color:#BDD7EE;font-size:12px'>
        Inbound & Outbound<br>TAT calculation<br>in HH:MM:SS format<br><br>
        YI-GI · GI-GW · GW-TW<br>TW-GO · GI-GO
        </p>
    </div>""", unsafe_allow_html=True)

with c2:
    st.markdown("""
    <div style='background:#375623;padding:20px;border-radius:12px;text-align:center;min-height:220px'>
        <div style='font-size:36px'>🏗️</div>
        <h4 style='color:white;margin:8px 0'>Loader Analysis</h4>
        <p style='color:#E2EFDA;font-size:12px'>
        Analyse loading<br>performance by<br>loader / shift /<br>material category
        </p>
    </div>""", unsafe_allow_html=True)

with c3:
    st.markdown("""
    <div style='background:#7B2D8B;padding:20px;border-radius:12px;text-align:center;min-height:220px'>
        <div style='font-size:36px'>📦</div>
        <h4 style='color:white;margin:8px 0'>Packer Analysis</h4>
        <p style='color:#E9D5F5;font-size:12px'>
        Analyse packing<br>performance by<br>packer / shift /<br>material category
        </p>
    </div>""", unsafe_allow_html=True)

with c4:
    st.markdown("""
    <div style='background:#833C00;padding:20px;border-radius:12px;text-align:center;min-height:220px'>
        <div style='font-size:36px'>⚖️</div>
        <h4 style='color:white;margin:8px 0'>Weighbridge<br>Congestion</h4>
        <p style='color:#FCE4D6;font-size:12px'>
        Analyse weighbridge<br>utilisation &<br>congestion by<br>time / shift / bridge
        </p>
    </div>""", unsafe_allow_html=True)

with c5:
    st.markdown("""
    <div style='background:#243F60;padding:20px;border-radius:12px;text-align:center;min-height:220px'>
        <div style='font-size:36px'>📊</div>
        <h4 style='color:white;margin:8px 0'>Category Analysis</h4>
        <p style='color:#BDD7EE;font-size:12px'>
        Upload processed file<br>& analyse TAT by<br>transporter / shift /<br>material / vehicle
        </p>
    </div>""", unsafe_allow_html=True)

st.markdown("---")
st.markdown("""
### 📋 How to Use
1. Click any module in the **left sidebar**
2. Upload your Excel file
3. Map columns using dropdowns
4. Click **Calculate / Analyse**
5. Preview results in table
6. Click **Download** for formatted Excel
""")
st.markdown("---")
st.markdown("""
<p style='color:gray;font-size:12px'>
🟢 Green cells = Calculated TAT &nbsp;|&nbsp;
🔵 Blue header = Input columns &nbsp;|&nbsp;
🟢 Green header = TAT columns &nbsp;|&nbsp;
All TAT values in <b>HH:MM:SS</b> format
</p>""", unsafe_allow_html=True)
