import streamlit as st
from streamlit_option_menu import option_menu
import plotly.express as px
import pandas as pd
from pandas import Series, DataFrame
import numpy as np
import os
import plotly.graph_objects as go
import matplotlib.pyplot as plt
from streamlit_extras.metric_cards import style_metric_cards
import seaborn as sns
import base64
from io import BytesIO
import plotly.graph_objects as go
from bs4 import BeautifulSoup
import locale
import re
from lxml import etree
from PIL import Image
from streamlit.components.v1 import html
######################################################################################################
# emojis https://streamlit-emoji-shortcodes-streamlit-app-gwckff.streamlit.app/
#Webpage config& tab name& Icon
st.set_page_config(page_title="Sales Dashboard",page_icon=":rainbow:",layout="wide")

title_row1, title_row2, title_row3, title_row4 = st.columns(4)


st.title(':world_map: C66 Sales Dashboard')
#Text Credit
st.write("by Arthur Chan")

#Move the title higher
st.markdown('<style>div.block-container{padding-top:1rem;}</style>',unsafe_allow_html=True)

#Move the title higher
st.markdown('<style>div.block-container{padding-top:1rem;}</style>',unsafe_allow_html=True)
######################################################################################################
#Create a browser for user to upload+ 
#@st.cache_data
#def load_data(file):
#       data = pd.read_excel(file)
#       return data
#uploaded_file = st.sidebar.file_uploader(":file_folder: Upload monthly report here")
#if uploaded_file is not None:
#        df = load_data(uploaded_file)
#        st.dataframe(df)
#唔show 17/18, cancel, tba資料
#else:
#os.chdir(r"/Users/arthurchan/Downloads/Sample")
#os.chdir(r"C:\Users\ArthurChan\OneDrive\VS Code\PythonProject_ESE\Sample Excel")

df = pd.read_excel(
               io='C66_All_AR_Summary-new_version.xlsm',engine= 'openpyxl',sheet_name='Summary', skiprows=0, usecols='A:AF',nrows=20000,)

#Sidebar Filter
# Create FY Invoice filter
st.sidebar.header(":point_down:Filter:")

# Define sidebar filters and create corresponding DataFrames for each filter
invoice_fy_filter = st.sidebar.multiselect("INVOICE_FY", df["INVOICE_FY"].unique(),default=["FY23/24"])
df_invoice_fy = df[df["INVOICE_FY"].isin(invoice_fy_filter)]

invoice_yr_filter = st.sidebar.multiselect("INVOICE_YR", df["INVOICE_YR"].unique())
df_invoice_yr = df[df["INVOICE_YR"].isin(invoice_yr_filter)]

invoice_fq_filter = st.sidebar.multiselect("INVOICE_FQ", df["INVOICE_FQ"].unique())
df_invoice_fq = df[df["INVOICE_FQ"].isin(invoice_fq_filter)]

invoice_month_filter = st.sidebar.multiselect("INVOICE_MONTH", df["INVOICE_MONTH"].unique())
df_invoice_month = df[df["INVOICE_MONTH"].isin(invoice_month_filter)]

# Add region filter
region_filter = st.sidebar.multiselect("REGION", df["REGION"].unique())
df_region = df[df["REGION"].isin(region_filter)]

branch_filter = st.sidebar.multiselect("Branch", df["Branch"].unique())
df_branch = df[df["Branch"].isin(branch_filter)]

type_filter = st.sidebar.multiselect("TYPE", df["TYPE"].unique())
df_type = df[df["TYPE"].isin(type_filter)]

brand_filter = st.sidebar.multiselect("BRAND", df["BRAND"].unique())
df_brand = df[df["BRAND"].isin(brand_filter)]

# Handle different filter combinations
if invoice_fy_filter and invoice_yr_filter and invoice_fq_filter and invoice_month_filter and region_filter and branch_filter and type_filter and brand_filter:
    # All filters are selected
    filtered_df = df_invoice_fy[df_invoice_yr["INVOICE_YR"].isin(invoice_yr_filter)]
    filtered_df = filtered_df[filtered_df["INVOICE_FQ"].isin(invoice_fq_filter)]
    filtered_df = filtered_df[filtered_df["INVOICE_MONTH"].isin(invoice_month_filter)]
    filtered_df = filtered_df[filtered_df["REGION"].isin(region_filter)]
    filtered_df = filtered_df[filtered_df["Branch"].isin(branch_filter)]
    filtered_df = filtered_df[filtered_df["TYPE"].isin(type_filter)]
    filtered_df = filtered_df[filtered_df["BRAND"].isin(brand_filter)]
elif not invoice_fy_filter and not invoice_yr_filter and not invoice_fq_filter and not invoice_month_filter and not region_filter and not branch_filter and not type_filter and not brand_filter:
    # No filters are selected
    filtered_df = df
else:
    # Other filter combinations
    filtered_df = pd.DataFrame(columns=df.columns)  # Create an empty DataFrame

    if invoice_fy_filter:
        filtered_df = pd.concat([filtered_df, df_invoice_fy])  # Use concat instead of append
    
    if invoice_yr_filter:
        if not invoice_fy_filter:
            filtered_df = pd.concat([filtered_df, df_invoice_yr])
        else:
            filtered_df = filtered_df[filtered_df["INVOICE_YR"].isin(invoice_yr_filter)]
    
    if invoice_fq_filter:
        if not invoice_fy_filter and not invoice_yr_filter:
            filtered_df = pd.concat([filtered_df, df_invoice_fq])
        else:
            filtered_df = filtered_df[filtered_df["INVOICE_FQ"].isin(invoice_fq_filter)]
    
    if invoice_month_filter:
        if not invoice_fy_filter and not invoice_yr_filter and not invoice_fq_filter:
            filtered_df = pd.concat([filtered_df, df_invoice_month])
        else:
            filtered_df = filtered_df[filtered_df["INVOICE_MONTH"].isin(invoice_month_filter)]
    
    if region_filter:
        if not invoice_fy_filter and not invoice_yr_filter and not invoice_fq_filter and not invoice_month_filter:
            filtered_df = pd.concat([filtered_df, df_region])
        else:
            filtered_df = filtered_df[filtered_df["REGION"].isin(region_filter)]
    
    if branch_filter:
        if not invoice_fy_filter and not invoice_yr_filter and not invoice_fq_filter and not invoice_month_filter and not region_filter:
            filtered_df = pd.concat([filtered_df, df_branch])
        else:
            filtered_df = filtered_df[filtered_df["Branch"].isin(branch_filter)]
    
    if type_filter:
        if not invoice_fy_filter and not invoice_yr_filter and not invoice_fq_filter and not invoice_month_filter and not region_filter and not branch_filter:
            filtered_df = pd.concat([filtered_df, df_type])
        else:
            filtered_df = filtered_df[filtered_df["TYPE"].isin(type_filter)]
    
    if brand_filter:
        if not invoice_fy_filter and not invoice_yr_filter and not invoice_fq_filter and not invoice_month_filter and not region_filter and not branch_filter and not type_filter:
            filtered_df = pd.concat([filtered_df, df_brand])
        else:
            filtered_df = filtered_df[filtered_df["BRAND"].isin(brand_filter)]

############################################################################################################################################################################################################        
#Create tabs after overall summary

# Make the tab font bigger
font_css = """
<style>
button[data-baseweb="tab"] > div[data-testid="stMarkdownContainer"] > p {
  font-size: 24px;
}
</style>
"""

st.write(font_css, unsafe_allow_html=True)
tab1, tab2, tab3 ,tab4,tab5= st.tabs([":wedding: Overview",":earth_asia: Branch",":blue_book: Invoice Details",":package: Brand",":handshake: Customer"])

#TAB 1: Overall category
################################################################################################################################################
with tab1:

#LINE CHART of Overall Invoice Amount
       st.subheader(":chart_with_upwards_trend: 月份:orange[同比]:")
       InvoiceAmount_df2 = filtered_df.round(0).groupby(by = ["INVOICE_FY","INVOICE_FQ","INVOICE_MONTH"
                          ], as_index= False)["Functional Amount(HKD)"].sum()

       fig3 = go.Figure()
# 添加每个INVOICE_FY的折线
       fy_inv_values = InvoiceAmount_df2['INVOICE_FY'].unique()
       for fy_inv in fy_inv_values:
                   fy_inv_data = InvoiceAmount_df2[InvoiceAmount_df2['INVOICE_FY'] == fy_inv]
                   fig3.add_trace(go.Scatter(
                         x=fy_inv_data['INVOICE_MONTH'],
                         y=fy_inv_data['Functional Amount(HKD)'],
                         mode='lines+markers+text',
                         name=fy_inv,
                         text=fy_inv_data['Functional Amount(HKD)'],
                         textposition="bottom center",
                         texttemplate='%{text:.3s}',
                         hovertemplate='%{x}<br>%{y:.2f}',
                         marker=dict(size=10)))
       fig3.update_layout(xaxis=dict(
                         type='category',
                         categoryorder='array',
                         ),
                         yaxis=dict(showticklabels=True),
                         font=dict(family="Arial, Arial", size=12, color="Black"),
                         hovermode='x', showlegend=True,
                         legend=dict(orientation="h",font=dict(size=14)))
       fig3.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
       st.plotly_chart(fig3.update_layout(yaxis_showticklabels = True), use_container_width=True)
#############################################################################################################
#FY to FY Quarter Invoice Details:
       pvt6 = filtered_df.round(0).pivot_table(values="Functional Amount(HKD)",index=["INVOICE_FY"],columns=["INVOICE_FQ"],
                            aggfunc="sum",fill_value=0, margins=True,margins_name="Total")
       html11 = pvt6.applymap('HKD{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             #st.dataframe(pvt6.style.highlight_max(color = 'yellow', axis = 0)
             #                       .format("HKD{:,}"), use_container_width=True)   
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
       html12 = html11.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
       html13 = html12.replace('<th>Q1</th>', '<th style="background-color: lightgrey">Q1</th>')
       html14 = html13.replace('<th>Q2</th>', '<th style="background-color: lightgreen">Q2</th>')
       html15 = html14.replace('<th>Q3</th>', '<th style="background-color: lightgrey">Q3</th>')
       html16 = html15.replace('<th>Q4</th>', '<th style="background-color: lightgreen">Q4</th>')
       html117 = html16.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
       html_with_style = str(f'<div style="zoom: 1.2;">{html117}</div>')
       st.markdown(html_with_style, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
       csv2 = pvt6.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
       st.download_button(label='Download Table', data=csv2, file_name='Monthly_Sales.csv', mime='text/csv')
######################################################################################################################
with tab2:
        one_column, two_column= st.columns(2)
        with one_column:
#All Regional total inv amount BAR CHART
              st.subheader(":bar_chart: Invoice Amount_:orange[FY](Available to show :orange[Multiple FY]):")
              category2_df = filtered_df.round(0).groupby(by=["INVOICE_FY","Branch"], 
                       as_index=False)["Functional Amount(HKD)"].sum().sort_values(by="Functional Amount(HKD)", ascending=False)
              df_contract_vs_invoice = px.bar(category2_df, x="INVOICE_FY", y="Functional Amount(HKD)", color="Branch", text_auto='.3s')

# 更改顏色
              colors = {"SZX": "orange","SHA": "lightblue","BJS": "Khaki","CTU": "lightgreen","XIY": "grey"}
              for trace in df_contract_vs_invoice.data:
                    Branch = trace.name.split("=")[-1]
                    trace.marker.color = colors.get(Branch, "blue")

# 更改字體
              df_contract_vs_invoice.update_layout(font=dict(family="Arial", size=14))
              df_contract_vs_invoice.update_traces(marker_line_color='black', marker_line_width=2,opacity=1)

# 將barmode設置為"group"以顯示多條棒形圖
              df_contract_vs_invoice.update_layout(barmode='group')

# 将图例放在底部
              df_contract_vs_invoice.update_layout(legend=dict(orientation="h",font=dict(size=14), yanchor="bottom", y=1.03, xanchor="right", x=1))
              st.plotly_chart(df_contract_vs_invoice, use_container_width=True) 
##############################################################################################################################          
# LINE CHART of Regional Comparision
        with two_column:
              st.subheader(":chart_with_upwards_trend: Invoice Amount Trend_:orange[All Branch in one]:")
              InvoiceAmount_df2 = filtered_df.round(0).groupby(by = ["INVOICE_FQ","Branch"], as_index= False)["Functional Amount(HKD)"].sum()
        # 使用pivot_table函數來重塑數據，使每個Region成為一個列
              InvoiceAmount_df2 = InvoiceAmount_df2.pivot_table(index="INVOICE_FQ", columns="Branch", values="Functional Amount(HKD)", fill_value=0).reset_index()
        # 使用melt函數來恢復原來的長格式，並保留0值
              InvoiceAmount_df2 = InvoiceAmount_df2.melt(id_vars="INVOICE_FQ", value_name="Functional Amount(HKD)", var_name="Branch")
              fig2 = px.line(InvoiceAmount_df2,
                       x = "INVOICE_FQ",
                       y = "Functional Amount(HKD)",
                       color='Branch',
                       markers=True,
                       text="Functional Amount(HKD)",
                       color_discrete_map={'SZX': 'orange','SHA': 'lightblue',
                                           'BJS': 'Khaki','CTU': 'lightgreen','XIY': 'grey'})
              # 更新圖表的字體大小和粗細
              fig2.update_layout(font=dict(
                    family="Arial, Arial",
                    size=12,
                    color="Black"))
              fig2.update_layout(legend=dict(orientation="h",font=dict(size=14), yanchor="bottom", y=1.02, xanchor="right", x=1))
              fig2.update_traces(marker_size=9, textposition="bottom center", texttemplate='%{text:.2s}')
              st.plotly_chart(fig2.update_layout(yaxis_showticklabels = True), use_container_width=True)
