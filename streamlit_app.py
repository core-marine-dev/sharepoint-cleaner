import streamlit as st

st.set_page_config(page_title="SharePoint Version Cleaner", page_icon="🧹")

st.markdown(
    """
    <style>
        .stAppDeployButton {display:none;}
    </style>
""",
    unsafe_allow_html=True,
)

pg = st.navigation(
    [
        st.Page("pages/sharepoint.py", title="SharePoint Version Cleaner", icon="🧹"),
    ]
)

pg.run()
