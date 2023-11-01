import streamlit as st

demoText = 'I am Apurba Khanra'

with open("excel_files.zip", "rb") as fp:
    btn = st.download_button(
        label="Download ZIP",
        data=fp,
        file_name="myfile.zip",
        mime="application/zip"
    )