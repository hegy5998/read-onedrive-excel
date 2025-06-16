import streamlit as st
import pandas as pd
import requests
from io import BytesIO

def convert_onedrive_to_download_url(share_url):
    if "1drv.ms" in share_url:
        resp = requests.get(share_url)
        return resp.url.replace("redir?", "download?").split("&")[0]
    elif "onedrive.live.com" in share_url:
        return share_url.replace("redir?", "download?").split("&")[0]
    else:
        raise ValueError("不支援的 OneDrive 分享連結格式")

st.title("從 OneDrive 載入 Excel")

share_url = st.text_input("請貼上 OneDrive 分享連結")

if share_url:
    try:
        download_url = convert_onedrive_to_download_url(share_url)
        response = requests.get(download_url)
        if response.status_code == 200:
            excel_data = BytesIO(response.content)
            df = pd.read_excel(excel_data, engine='openpyxl')
            st.success("成功讀取 Excel 檔案！")
            st.dataframe(df)
        else:
            st.error("下載檔案失敗，請檢查連結是否正確")
    except Exception as e:
        st.error(f"讀取過程發生錯誤: {e}")
