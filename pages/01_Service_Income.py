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

#Title& #Text Credit
with title_row1:
     st.title(':toolbox: C66 Sales Dashboard')
     st.write("by Arthur Chan")

#FIRST PIC
#image_path = '/Users/arthurchan/Downloads/Sample/SERVICE.png'
#image_path = "/Users/ArthurChan/OneDrive/VS Code/PythonProject_ESE/SERVICE.png"
image_path = 'SERVICE.png'
image = Image.open(image_path)
target_width = 600# 設置目標寬度和高度
target_height = 300
resized_image = image.resize((target_width, target_height))# 縮小圖片

with title_row2:
# 在Streamlit應用程式中顯示縮小後的圖片
# 加載圖片
     st.image(resized_image, use_column_width=False, output_format='PNG')

#Second PIC
#image_path2 = '/Users/arthurchan/Downloads/Sample/SERVICE2.png'
#image_path = "/Users/ArthurChan/OneDrive/VS Code/PythonProject_ESE/SERVICE.png"
image_path2 = 'SERVICE2.png'
image2 = Image.open(image_path2)
target_width2 = 600# 設置目標寬度和高度
target_height2 = 300
resized_image2 = image2.resize((target_width2, target_height2))# 縮小圖片

with title_row4:
# 在Streamlit應用程式中顯示縮小後的圖片
# 加載圖片
     st.image(resized_image2, use_column_width=False, output_format='PNG')


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

branch_filter = st.sidebar.multiselect("BRANCH", df["Branch"].unique())
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
#Change the number of month into string
filtered_df["INVOICE_MONTH"] = filtered_df["INVOICE_MONTH"].astype(str)
#filtered_df["INVOICE_YR"] = filtered_df["INVOICE_YR"].astype(str)



st.write(font_css, unsafe_allow_html=True)
tab1, tab2, tab3 ,tab4,tab5,tab6= st.tabs([":wedding: Overview",":earth_asia: Region",":hammer_and_wrench: SERVICE Type",":books: Brand",":handshake: Customer& Project",":memo: Customer& Project"])

#TAB 1: Overall category
################################################################################################################################################
with tab1:

#LINE CHART of Overall Invoice Amount
       st.subheader(":chart_with_upwards_trend: 月份:orange[同比]:")
       InvoiceAmount_df2 = filtered_df.round(0).groupby(by = ["INVOICE_FY","INVOICE_YR","INVOICE_FQ","INVOICE_MONTH"
                          ], as_index= False)["Functional Amount(HKD)"].sum()

# 确保 "Inv Month" 列中的所有值都出现
       sort_Month_order = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
       InvoiceAmount_df2 = InvoiceAmount_df2.groupby(["INVOICE_FY", "INVOICE_MONTH"]).sum().reindex(pd.MultiIndex.from_product([InvoiceAmount_df2['INVOICE_FY'].unique(), sort_Month_order],
                                   names=['INVOICE_FY','INVOICE_MONTH'])).fillna(0).reset_index()
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
              colors = {"SZX": "orange","SHA": "lightblue","BJS": "Khaki","CTU": "lightgreen","XIY": "purple"}
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

#All Region PIE CHART
        with two_column:
             st.subheader(":round_pushpin: Invoice Percentage_:orange[Accumulate]:")
# 创建示例数据框
             brand_data = filtered_df.round(0).groupby(by=["INVOICE_FY","Branch"],
                     as_index=False)["Functional Amount(HKD)"].sum().sort_values(by="Functional Amount(HKD)", ascending=False)         
             brandinvpie_df = pd.DataFrame(brand_data)
# 按照指定順序排序 
             brandinvpie_df["Branch"] = brandinvpie_df["Branch"].replace(to_replace=[x for x in brandinvpie_df["Branch"
                                       ].unique() if x not in ["SZX","SHA", "BJS", "XIY", "CTU"]], value="OTHERS")
             brandinvpie_df["Branch"] = pd.Categorical(brandinvpie_df["Branch"], ["SZX","SHA", "BJS", "XIY", "CTU","OTHERS"])
# 创建饼状图
             df_pie = px.pie(brandinvpie_df, values="Functional Amount(HKD)", names="Branch", color="Branch", color_discrete_map={
                      "SHA": "lightblue", "SZX": "orange", "CTU": "lightgreen","BJS": "khaki", "XIY":"purple"})
# 设置字体和样式
             df_pie.update_layout(
                   font=dict(family="Arial", size=14, color="black"),
                   legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
# 显示百分比标签
             df_pie.update_traces(textposition='outside', textinfo='label+percent', marker_line_width=2,opacity=1)
# 在Streamlit中显示图表
             st.plotly_chart(df_pie, use_container_width=True)
             
              
##############################################################################################################################                   
        tab2row2one_column, tab2row2two_column= st.columns(2)
        with tab2row2one_column:
# LINE CHART of SOUTH CHINA FY/FY
              st.divider()
              st.subheader(":chart_with_upwards_trend: :orange[SOUTH CHINA] Inv Amt Trend_MONTHLY(Available to Show :orange[Multiple FY]):")
              df_Single_south = filtered_df.query('REGION == "SOUTH"').round(0).groupby(by = ["INVOICE_FY",
                                 "INVOICE_MONTH"], as_index= False)["Functional Amount(HKD)"].sum()
# 确保 "INVOICE_FQ" 列中的所有值都出现在 df_Single_region 中
              all_fq_invoice_values = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
              df_Single_south = df_Single_south.groupby(["INVOICE_FY", "INVOICE_MONTH"]).sum().reindex(pd.MultiIndex.from_product([df_Single_south['INVOICE_FY'].unique(), all_fq_invoice_values],
                                   names=['INVOICE_FY', 'INVOICE_MONTH'])).fillna(0).reset_index()
              fig4 = go.Figure()

# 添加每个INVOICE_FY的折线
              fy_inv_values = df_Single_south['INVOICE_FY'].unique()
              for fy_inv in fy_inv_values:
               fy_inv_data = df_Single_south[df_Single_south['INVOICE_FY'] == fy_inv]
               fig4.add_trace(go.Scatter(
                          x=fy_inv_data['INVOICE_MONTH'],
                          y=fy_inv_data['Functional Amount(HKD)'],
                          mode='lines+markers+text',
                          name=fy_inv,
                          text=fy_inv_data['Functional Amount(HKD)'],
                          textposition="bottom center",
                          texttemplate='%{text:.3s}',
                          hovertemplate='%{x}<br>%{y:.2f}',
                          marker=dict(size=10)))
               fig4.update_layout(xaxis=dict(
                          type='category',
                          categoryorder='array',
                          categoryarray=all_fq_invoice_values),
                          yaxis=dict(showticklabels=True),
                          font=dict(family="Arial, Arial", size=12, color="Black"),
                          hovermode='x', showlegend=True,
                          legend=dict(orientation="h",font=dict(size=14)),paper_bgcolor='rgba(255,165,0,0.3)')
             
              fig4.update_layout(legend=dict(orientation="h",font=dict(size=14), yanchor="bottom", y=1.02, xanchor="right", x=1))

              st.plotly_chart(fig4.update_layout(yaxis_showticklabels = True), use_container_width=True)           
#SOUTH Region Invoice Details FQ_FQ:
              pvt8 = filtered_df.query('REGION == "SOUTH"').round(0).pivot_table(values="Functional Amount(HKD)",
                     index=['INVOICE_FY'],columns=["INVOICE_FQ"],aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_values(by='INVOICE_FY',ascending=True)
            
              html76 = pvt8.applymap('HKD{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
              html77 = html76.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
              html78 = html77.replace('<th>Q1</th>', '<th style="background-color: orange">Q1</th>')
              html79 = html78.replace('<th>Q2</th>', '<th style="background-color: orange">Q2</th>')
              html80 = html79.replace('<th>Q3</th>', '<th style="background-color: orange">Q3</th>')
              html81 = html80.replace('<th>Q4</th>', '<th style="background-color: orange">Q4</th>')
              html822 = html81.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
# 放大pivot table
              html833 = f'<div style="zoom: 0.7;">{html822}</div>'

              st.markdown(html833, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
              csv11 = pvt8.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
              st.download_button(label='Download Table', data=csv11, file_name='South_Sales.csv', mime='text/csv')
              st.divider()

 
 
        with tab2row2two_column:
# LINE CHART of EAST CHINA FY/FY
              st.divider()
              st.subheader(":chart_with_upwards_trend: :orange[EAST CHINA] Inv Amt Trend_MONTHLY(Available to Show :orange[Multiple FY)]:")
              df_Single_region = filtered_df.query('REGION == "EAST"').round(0).groupby(by = ["INVOICE_FY","INVOICE_MONTH"], as_index= False)["Functional Amount(HKD)"].sum()
# 确保 "INVOICE_FQ" 列中的所有值都出现在 df_Single_region 中
              all_fq_invoice_values = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
              df_Single_region = df_Single_region.groupby(["INVOICE_FY", "INVOICE_MONTH"]).sum().reindex(pd.MultiIndex.from_product([df_Single_region['INVOICE_FY'].unique(), all_fq_invoice_values],
                                   names=['INVOICE_FY', 'INVOICE_MONTH'])).fillna(0).reset_index()
              fig5 = go.Figure()

# 添加每个FY_INV的折线
              fy_inv_values = df_Single_region['INVOICE_FY'].unique()
              for fy_inv in fy_inv_values:
               fy_inv_data = df_Single_region[df_Single_region['INVOICE_FY'] == fy_inv]
               fig5.add_trace(go.Scatter(
                          x=fy_inv_data['INVOICE_MONTH'],
                          y=fy_inv_data['Functional Amount(HKD)'],
                          mode='lines+markers+text',
                          name=fy_inv,
                          text=fy_inv_data['Functional Amount(HKD)'],
                          textposition="bottom center",
                          texttemplate='%{text:.2s}',
                          hovertemplate='%{x}<br>%{y:.2f}',
                          marker=dict(size=10)))
               fig5.update_layout(xaxis=dict(
                          type='category',
                          categoryorder='array',
                          categoryarray=all_fq_invoice_values),
                          yaxis=dict(showticklabels=True),
                          font=dict(family="Arial, Arial", size=12, color="Black"),
                          hovermode='x', showlegend=True,
                          legend=dict(orientation="h",font=dict(size=14)),
                          paper_bgcolor='rgba(0,150,255,0.1)')
             
              fig5.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
              st.plotly_chart(fig5.update_layout(yaxis_showticklabels = True), use_container_width=True)

########################             
#EAST Region Invoice Details FQ_FQ:
              pvt9 = filtered_df.query('REGION == "EAST"').round(0).pivot_table(values="Functional Amount(HKD)",index=['INVOICE_FY'],columns=["INVOICE_FQ"],
                            aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_values(by='INVOICE_FY',ascending=True)
            
              html83 = pvt9.applymap('HKD{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
              html84 = html83.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
              html85 = html84.replace('<th>Q1</th>', '<th style="background-color: lightblue">Q1</th>')
              html86 = html85.replace('<th>Q2</th>', '<th style="background-color: lightblue">Q2</th>')
              html87 = html86.replace('<th>Q3</th>', '<th style="background-color: lightblue">Q3</th>')
              html88 = html87.replace('<th>Q4</th>', '<th style="background-color: lightblue">Q4</th>')
              html89 = html88.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
# 放大pivot table
              html900 = f'<div style="zoom: 0.7;">{html89}</div>'
             
              st.markdown(html900, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
              csv12 = pvt9.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
              st.download_button(label='Download Table', data=csv12, file_name='East_Sales.csv', mime='text/csv')

              st.divider()
################################################# 
        three_column, four_column= st.columns(2)
  
        with three_column:
# LINE CHART of NORTH CHINA FY/FY
             st.subheader(":chart_with_upwards_trend: :orange[NORTH CHINA] Inv Amt Trend_MONTHLY(Available to Show :orange[Multiple FY]):")

             df_Single_north = filtered_df.query('REGION == "NORTH"').round(0).groupby(by=["INVOICE_FY", "INVOICE_MONTH"],
                                as_index=False)["Functional Amount(HKD)"].sum()
# 确保 "INVOICE_FQ" 列中的所有值都出现在 df_Single_region 中
             all_fq_invoice_values = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             df_Single_north = df_Single_north.groupby(["INVOICE_FY", "INVOICE_MONTH"]).sum().reindex(pd.MultiIndex.from_product([df_Single_north['INVOICE_FY'].unique(), all_fq_invoice_values],
                               names=['INVOICE_FY', 'INVOICE_MONTH'])).fillna(0).reset_index()
             fig7 = go.Figure()

# 添加每个INVOICE_FY的折线
             fy_inv_values = df_Single_north['INVOICE_FY'].unique()
             for fy_inv in fy_inv_values:
               fy_inv_data = df_Single_north[df_Single_north['INVOICE_FY'] == fy_inv]
               fig7.add_trace(go.Scatter(
                          x=fy_inv_data['INVOICE_MONTH'],
                          y=fy_inv_data['Functional Amount(HKD)'],
                          mode='lines+markers+text',
                          name=fy_inv,
                          text=fy_inv_data['Functional Amount(HKD)'],
                          textposition="bottom center",
                          texttemplate='%{text:.3s}',
                          hovertemplate='%{x}<br>%{y:.2f}',
                          marker=dict(size=10)))
               fig7.update_layout(xaxis=dict(
                          type='category',
                          categoryorder='array',
                          categoryarray=all_fq_invoice_values),
                          yaxis=dict(showticklabels=True),
                          font=dict(family="Arial, Arial", size=12, color="Black"),
                          hovermode='x', showlegend=True,
                          legend=dict(orientation="h",font=dict(size=14)),
                          paper_bgcolor='khaki')
              
             fig7.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
             st.plotly_chart(fig7.update_layout(yaxis_showticklabels = True), use_container_width=True)

#NORTH Region Invoice Details FQ_FQ:
             pvt10 = filtered_df.query('REGION == "NORTH"').round(0).pivot_table(values="Functional Amount(HKD)",index=['INVOICE_FY'],columns=["INVOICE_FQ"],
                     aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_values(by='INVOICE_FY',ascending=True)

             html62 = pvt10.applymap('HKD{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
             html63 = html62.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
             html64 = html63.replace('<th>Q1</th>', '<th style="background-color: khaki">Q1</th>')
             html65 = html64.replace('<th>Q2</th>', '<th style="background-color: khaki">Q2</th>')
             html66 = html65.replace('<th>Q3</th>', '<th style="background-color: khaki">Q3</th>')
             html67 = html66.replace('<th>Q4</th>', '<th style="background-color: khaki">Q4</th>')
             html68 = html67.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')

# 放大pivot table
             html699 = f'<div style="zoom: 0.7;">{html68}</div>'
             st.markdown(html699, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
             csv9 = pvt10.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
             st.download_button(label='Download Table', data=csv9, file_name='North_Sales.csv', mime='text/csv')

             st.divider()
##################################################
        with four_column:
# LINE CHART of WEST CHINA FY/FY
             st.subheader(":chart_with_upwards_trend: :orange[WEST CHINA] Inv Amt Trend_MONTHLY(Available to Show :orange[Multiple FY]):")
             df_Single_west = filtered_df.query('REGION == "WEST"').round(0).groupby(by=["INVOICE_FY", "INVOICE_MONTH"],
                                as_index=False)["Functional Amount(HKD)"].sum()
# 确保 "INVOICE_FQ" 列中的所有值都出现在 df_Single_region 中
             all_fq_invoice_values = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             df_Single_west = df_Single_west.groupby(["INVOICE_FY", "INVOICE_MONTH"]).sum().reindex(pd.MultiIndex.from_product([df_Single_west['INVOICE_FY'].unique(), all_fq_invoice_values],
                               names=['INVOICE_FY', 'INVOICE_MONTH'])).fillna(0).reset_index()
             fig6 = go.Figure()

# 添加每个FY_INV的折线
             fy_inv_values = df_Single_west['INVOICE_FY'].unique()
             for fy_inv in fy_inv_values:
               fy_inv_data = df_Single_west[df_Single_west['INVOICE_FY'] == fy_inv]
               fig6.add_trace(go.Scatter(
                          x=fy_inv_data['INVOICE_MONTH'],
                          y=fy_inv_data['Functional Amount(HKD)'],
                          mode='lines+markers+text',
                          name=fy_inv,
                          text=fy_inv_data['Functional Amount(HKD)'],
                          textposition="bottom center",
                          texttemplate='%{text:.3s}',
                          hovertemplate='%{x}<br>%{y:.2f}',
                          marker=dict(size=10)))
               fig6.update_layout(xaxis=dict(
                          type='category',
                          categoryorder='array',
                          categoryarray=all_fq_invoice_values),
                          yaxis=dict(showticklabels=True),
                          font=dict(family="Arial, Arial", size=12, color="Black"),
                          hovermode='x', showlegend=True,
                          legend=dict(orientation="h",font=dict(size=14)),
                          paper_bgcolor='lightgreen')
              
             fig6.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
             st.plotly_chart(fig6.update_layout(yaxis_showticklabels = True), use_container_width=True)
#WEST Region Invoice Details FQ_FQ:
             pvt18 = filtered_df.query('REGION == "WEST"').round(0).pivot_table(values="Functional Amount(HKD)",index=['INVOICE_FY'],columns=["INVOICE_FQ"],
                            aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_values(by='INVOICE_FY',ascending=True)
            
             html69 = pvt18.applymap('HKD{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
             html70 = html69.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
             html71 = html70.replace('<th>Q1</th>', '<th style="background-color: lightgreen">Q1</th>')
             html72 = html71.replace('<th>Q2</th>', '<th style="background-color: lightgreen">Q2</th>')
             html73 = html72.replace('<th>Q3</th>', '<th style="background-color: lightgreen">Q3</th>')
             html74 = html73.replace('<th>Q4</th>', '<th style="background-color: lightgreen">Q4</th>')
             html75 = html74.replace('<th>Total</th>', '<th style="background-color: lightgreen">Total</th>')
# 放大pivot table
             html766 = f'<div style="zoom: 0.7;">{html75}</div>'

             st.markdown(html766, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
             csv10 = pvt18.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
             st.download_button(label='Download Table', data=csv10, file_name='West_Sales.csv', mime='text/csv')
             st.divider()

            
# LINE CHART of Regional Comparision              
        st.subheader(":chart_with_upwards_trend: Invoice Amount Trend_:orange[Quarterly Accumulation]:")
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
#######################################################################################################################################

##############################################################################################################################  
with tab3:
        tab3one_column, tab3two_column= st.columns(2)
        with tab3one_column:
#All Regional total inv amount BAR CHART
              st.subheader(":bar_chart: Invoice Amount_:orange[FY](Available to show :orange[Multiple FY]):")
              category2_df = filtered_df.round(0).groupby(by=["INVOICE_FY","TYPE"], 
                       as_index=False)["Functional Amount(HKD)"].sum().sort_values(by="Functional Amount(HKD)", ascending=False)
              df_contract_vs_invoice = px.bar(category2_df, x="INVOICE_FY", y="Functional Amount(HKD)", color="TYPE", text_auto='.3s')

# 更改顏色
              colors = {"SPARES/OTHER": "pink","CONTRACT_FEE": "blue","SERVICE_CHARGE": "yellow","FEEDER": "lightgreen"}
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

#All Region PIE CHART
        with tab3two_column:
             st.subheader(":round_pushpin: Invoice Percentage_:orange[Accumulate]:")
# 创建示例数据框
             brand_data = filtered_df.round(0).groupby(by=["INVOICE_FY","TYPE"],
                     as_index=False)["Functional Amount(HKD)"].sum().sort_values(by="Functional Amount(HKD)", ascending=False)         
             brandinvpie_df = pd.DataFrame(brand_data)
# 按照指定順序排序 
             brandinvpie_df["TYPE"] = brandinvpie_df["TYPE"].replace(to_replace=[x for x in brandinvpie_df["TYPE"
                                       ].unique() if x not in ["SPARES/OTHER","FEEDER", "SERVICE_CHARGE", "CONTRACT_FEE"]], value="OTHERS")
             brandinvpie_df["TYPE"] = pd.Categorical(brandinvpie_df["TYPE"], ["SPARES/OTHER","FEEDER", "SERVICE_CHARGE", "CONTRACT_FEE","OTHERS"])
# 创建饼状图
             df_pie = px.pie(brandinvpie_df, values="Functional Amount(HKD)", names="TYPE", color="TYPE", color_discrete_map={
                      "SPARES/OTHER": "pink","CONTRACT_FEE": "blue","SERVICE_CHARGE": "yellow","FEEDER": "lightgreen"})
# 设置字体和样式
             df_pie.update_layout(
                   font=dict(family="Arial", size=14, color="black"),
                   legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
# 显示百分比标签
             df_pie.update_traces(textposition='outside', textinfo='label+percent', marker_line_width=2,opacity=1)
# 在Streamlit中显示图表
             st.plotly_chart(df_pie, use_container_width=True)
##############################################################################################################################                   
        tab3row2one_column, tab3row2two_column= st.columns(2)
        with tab3row2one_column:
# LINE CHART of SPARES/OTHER FY/FY
              st.divider()
              st.subheader(":chart_with_upwards_trend: :orange[SPARES/OTHER] Inv Amt Trend_MONTHLY(Available to Show :orange[Multiple FY]):")
              df_Single_SPARES = filtered_df.query('TYPE == "SPARES/OTHER"').round(0).groupby(by = ["INVOICE_FY",
                                 "INVOICE_MONTH"], as_index= False)["Functional Amount(HKD)"].sum()
# 确保 "INVOICE_FQ" 列中的所有值都出现在 df_Single_region 中
              all_fq_invoice_values = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
              df_Single_SPARES = df_Single_SPARES.groupby(["INVOICE_FY", "INVOICE_MONTH"]).sum().reindex(pd.MultiIndex.from_product([df_Single_SPARES['INVOICE_FY'].unique(), all_fq_invoice_values],
                                   names=['INVOICE_FY', 'INVOICE_MONTH'])).fillna(0).reset_index()
              fig4 = go.Figure()

# 添加每个INVOICE_FY的折线
              fy_inv_values = df_Single_SPARES['INVOICE_FY'].unique()
              for fy_inv in fy_inv_values:
               fy_inv_data = df_Single_SPARES[df_Single_SPARES['INVOICE_FY'] == fy_inv]
               fig4.add_trace(go.Scatter(
                          x=fy_inv_data['INVOICE_MONTH'],
                          y=fy_inv_data['Functional Amount(HKD)'],
                          mode='lines+markers+text',
                          name=fy_inv,
                          text=fy_inv_data['Functional Amount(HKD)'],
                          textposition="bottom center",
                          texttemplate='%{text:.3s}',
                          hovertemplate='%{x}<br>%{y:.2f}',
                          marker=dict(size=10)))
               fig4.update_layout(xaxis=dict(
                          type='category',
                          categoryorder='array',
                          categoryarray=all_fq_invoice_values),
                          yaxis=dict(showticklabels=True),
                          font=dict(family="Arial, Arial", size=12, color="Black"),
                          hovermode='x', showlegend=True,
                          legend=dict(orientation="h",font=dict(size=14)),paper_bgcolor='pink')
             
              fig4.update_layout(legend=dict(orientation="h",font=dict(size=14), yanchor="bottom", y=1.02, xanchor="right", x=1))

              st.plotly_chart(fig4.update_layout(yaxis_showticklabels = True), use_container_width=True)           
#SPARES/ OTHER Invoice Details FQ_FQ:
              pvt1 = filtered_df.query('TYPE == "SPARES/OTHER"').round(0).pivot_table(values="Functional Amount(HKD)",
                     index=['INVOICE_FY'],columns=["INVOICE_FQ"],aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_values(by='INVOICE_FY',ascending=True)
            
              html76 = pvt1.applymap('HKD{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
              html77 = html76.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
              html78 = html77.replace('<th>Q1</th>', '<th style="background-color: pink">Q1</th>')
              html79 = html78.replace('<th>Q2</th>', '<th style="background-color: pink">Q2</th>')
              html80 = html79.replace('<th>Q3</th>', '<th style="background-color: pink">Q3</th>')
              html81 = html80.replace('<th>Q4</th>', '<th style="background-color: pink">Q4</th>')
              html822 = html81.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
# 放大pivot table
              html833 = f'<div style="zoom: 0.7;">{html822}</div>'

              st.markdown(html833, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
              csv11 = pvt1.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
              st.download_button(label='Download Table', data=csv11, file_name='SPARES/OTHER.csv', mime='text/csv')
              st.divider()

        with tab3row2two_column:
# LINE CHART of SERVICE_CHARGE FY/FY
              st.divider()
              st.subheader(":chart_with_upwards_trend: :orange[SERVICE_CHARGE] Inv Amt Trend_MONTHLY(Available to Show :orange[Multiple FY)]:")
              df_Single_SERVICE_CHARGE = filtered_df.query('TYPE == "SERVICE_CHARGE"').round(0).groupby(by = ["INVOICE_FY","INVOICE_MONTH"], as_index= False)["Functional Amount(HKD)"].sum()
# 确保 "INVOICE_FQ" 列中的所有值都出现在 df_Single_SERVICE_CHARGE 中
              all_fq_invoice_values = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
              df_Single_SERVICE_CHARGE = df_Single_SERVICE_CHARGE.groupby(["INVOICE_FY", "INVOICE_MONTH"]).sum().reindex(pd.MultiIndex.from_product([df_Single_SERVICE_CHARGE['INVOICE_FY'].unique(), all_fq_invoice_values],
                                   names=['INVOICE_FY', 'INVOICE_MONTH'])).fillna(0).reset_index()
              fig5 = go.Figure()

# 添加每个FY_INV的折线
              fy_inv_values = df_Single_SERVICE_CHARGE['INVOICE_FY'].unique()
              for fy_inv in fy_inv_values:
               fy_inv_data = df_Single_SERVICE_CHARGE[df_Single_SERVICE_CHARGE['INVOICE_FY'] == fy_inv]
               fig5.add_trace(go.Scatter(
                          x=fy_inv_data['INVOICE_MONTH'],
                          y=fy_inv_data['Functional Amount(HKD)'],
                          mode='lines+markers+text',
                          name=fy_inv,
                          text=fy_inv_data['Functional Amount(HKD)'],
                          textposition="bottom center",
                          texttemplate='%{text:.2s}',
                          hovertemplate='%{x}<br>%{y:.2f}',
                          marker=dict(size=10)))
               fig5.update_layout(xaxis=dict(
                          type='category',
                          categoryorder='array',
                          categoryarray=all_fq_invoice_values),
                          yaxis=dict(showticklabels=True),
                          font=dict(family="Arial, Arial", size=12, color="Black"),
                          hovermode='x', showlegend=True,
                          legend=dict(orientation="h",font=dict(size=14)),
                          paper_bgcolor='yellow')
             
              fig5.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
              st.plotly_chart(fig5.update_layout(yaxis_showticklabels = True), use_container_width=True)

########################             
#SERVICE_CHARGE Invoice Details FQ_FQ:
              pvt2 = filtered_df.query('TYPE == "SERVICE_CHARGE"').round(0).pivot_table(values="Functional Amount(HKD)",index=['INVOICE_FY'],columns=["INVOICE_FQ"],
                            aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_values(by='INVOICE_FY',ascending=True)
            
              html83 = pvt2.applymap('HKD{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
              html84 = html83.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
              html85 = html84.replace('<th>Q1</th>', '<th style="background-color: khaki">Q1</th>')
              html86 = html85.replace('<th>Q2</th>', '<th style="background-color: khaki">Q2</th>')
              html87 = html86.replace('<th>Q3</th>', '<th style="background-color: khaki">Q3</th>')
              html88 = html87.replace('<th>Q4</th>', '<th style="background-color: khaki">Q4</th>')
              html89 = html88.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
# 放大pivot table
              html900 = f'<div style="zoom: 0.7;">{html89}</div>'
             
              st.markdown(html900, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
              csv12 = pvt2.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
              st.download_button(label='Download Table', data=csv12, file_name='SERVICE_CHARGE.csv', mime='text/csv')

              st.divider()
################################################# 
        three_column, four_column= st.columns(2)
  
        with three_column:
# LINE CHART of CONTRACT_FEE FY/FY
             st.subheader(":chart_with_upwards_trend: :orange[CONTRACT_FEE] Inv Amt Trend_MONTHLY(Available to Show :orange[Multiple FY]):")

             df_Single_CONTRACT_FEE = filtered_df.query('TYPE == "CONTRACT_FEE"').round(0).groupby(by=["INVOICE_FY", "INVOICE_MONTH"],
                                as_index=False)["Functional Amount(HKD)"].sum()
# 确保 "INVOICE_FQ" 列中的所有值都出现在 df_Single_region 中
             all_fq_invoice_values = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             df_Single_CONTRACT_FEE = df_Single_CONTRACT_FEE.groupby(["INVOICE_FY", "INVOICE_MONTH"]).sum().reindex(pd.MultiIndex.from_product([df_Single_CONTRACT_FEE['INVOICE_FY'].unique(), all_fq_invoice_values],
                               names=['INVOICE_FY', 'INVOICE_MONTH'])).fillna(0).reset_index()
             fig7 = go.Figure()

# 添加每个INVOICE_FY的折线
             fy_inv_values = df_Single_CONTRACT_FEE['INVOICE_FY'].unique()
             for fy_inv in fy_inv_values:
               fy_inv_data = df_Single_CONTRACT_FEE[df_Single_CONTRACT_FEE['INVOICE_FY'] == fy_inv]
               fig7.add_trace(go.Scatter(
                          x=fy_inv_data['INVOICE_MONTH'],
                          y=fy_inv_data['Functional Amount(HKD)'],
                          mode='lines+markers+text',
                          name=fy_inv,
                          text=fy_inv_data['Functional Amount(HKD)'],
                          textposition="bottom center",
                          texttemplate='%{text:.3s}',
                          hovertemplate='%{x}<br>%{y:.2f}',
                          marker=dict(size=10)))
               fig7.update_layout(xaxis=dict(
                          type='category',
                          categoryorder='array',
                          categoryarray=all_fq_invoice_values),
                          yaxis=dict(showticklabels=True),
                          font=dict(family="Arial, Arial", size=12, color="Black"),
                          hovermode='x', showlegend=True,
                          legend=dict(orientation="h",font=dict(size=14)),
                          paper_bgcolor='lightblue')
              
             fig7.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
             st.plotly_chart(fig7.update_layout(yaxis_showticklabels = True), use_container_width=True)

#CONTRACT_FEE Invoice Details FQ_FQ:
             pvt3 = filtered_df.query('TYPE == "CONTRACT_FEE"').round(0).pivot_table(values="Functional Amount(HKD)",index=['INVOICE_FY'],columns=["INVOICE_FQ"],
                     aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_values(by='INVOICE_FY',ascending=True)

             html62 = pvt3.applymap('HKD{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
             html63 = html62.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
             html64 = html63.replace('<th>Q1</th>', '<th style="background-color: lightblue">Q1</th>')
             html65 = html64.replace('<th>Q2</th>', '<th style="background-color: lightblue">Q2</th>')
             html66 = html65.replace('<th>Q3</th>', '<th style="background-color: lightblue">Q3</th>')
             html67 = html66.replace('<th>Q4</th>', '<th style="background-color: lightblue">Q4</th>')
             html68 = html67.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')

# 放大pivot table
             html699 = f'<div style="zoom: 0.7;">{html68}</div>'
             st.markdown(html699, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
             csv9 = pvt3.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
             st.download_button(label='Download Table', data=csv9, file_name='CONTRACT_FEE.csv', mime='text/csv')

             st.divider()
##################################################
        with four_column:
# LINE CHART of FEEDER FY/FY
             st.subheader(":chart_with_upwards_trend: :orange[FEEDER] Inv Amt Trend_MONTHLY(Available to Show :orange[Multiple FY]):")
             df_Single_feeder = filtered_df.query('TYPE == "FEEDER"').round(0).groupby(by=["INVOICE_FY", "INVOICE_MONTH"],
                                as_index=False)["Functional Amount(HKD)"].sum()
# 确保 "INVOICE_FQ" 列中的所有值都出现在 df_Single_region 中
             all_fq_invoice_values = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             df_Single_feeder = df_Single_feeder.groupby(["INVOICE_FY", "INVOICE_MONTH"]).sum().reindex(pd.MultiIndex.from_product([df_Single_feeder['INVOICE_FY'].unique(), all_fq_invoice_values],
                               names=['INVOICE_FY', 'INVOICE_MONTH'])).fillna(0).reset_index()
             fig6 = go.Figure()

# 添加每个FY_INV的折线
             fy_inv_values = df_Single_feeder['INVOICE_FY'].unique()
             for fy_inv in fy_inv_values:
               fy_inv_data = df_Single_feeder[df_Single_feeder['INVOICE_FY'] == fy_inv]
               fig6.add_trace(go.Scatter(
                          x=fy_inv_data['INVOICE_MONTH'],
                          y=fy_inv_data['Functional Amount(HKD)'],
                          mode='lines+markers+text',
                          name=fy_inv,
                          text=fy_inv_data['Functional Amount(HKD)'],
                          textposition="bottom center",
                          texttemplate='%{text:.3s}',
                          hovertemplate='%{x}<br>%{y:.2f}',
                          marker=dict(size=10)))
               fig6.update_layout(xaxis=dict(
                          type='category',
                          categoryorder='array',
                          categoryarray=all_fq_invoice_values),
                          yaxis=dict(showticklabels=True),
                          font=dict(family="Arial, Arial", size=12, color="Black"),
                          hovermode='x', showlegend=True,
                          legend=dict(orientation="h",font=dict(size=14)),
                          paper_bgcolor='lightgreen')
              
             fig6.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
             st.plotly_chart(fig6.update_layout(yaxis_showticklabels = True), use_container_width=True)
#FEEDER Invoice Details FQ_FQ:
             pvt4 = filtered_df.query('TYPE == "FEEDER"').round(0).pivot_table(values="Functional Amount(HKD)",index=['INVOICE_FY'],columns=["INVOICE_FQ"],
                            aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_values(by='INVOICE_FY',ascending=True)
            
             html69 = pvt4.applymap('HKD{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
             html70 = html69.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
             html71 = html70.replace('<th>Q1</th>', '<th style="background-color: lightgreen">Q1</th>')
             html72 = html71.replace('<th>Q2</th>', '<th style="background-color: lightgreen">Q2</th>')
             html73 = html72.replace('<th>Q3</th>', '<th style="background-color: lightgreen">Q3</th>')
             html74 = html73.replace('<th>Q4</th>', '<th style="background-color: lightgreen">Q4</th>')
             html75 = html74.replace('<th>Total</th>', '<th style="background-color: lightgreen">Total</th>')
# 放大pivot table
             html766 = f'<div style="zoom: 0.7;">{html75}</div>'

             st.markdown(html766, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
             csv10 = pvt4.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
             st.download_button(label='Download Table', data=csv10, file_name='FEEDER.csv', mime='text/csv')
             st.divider()             
              
# LINE CHART of SERVICE TYPE Comparision              
        st.subheader(":chart_with_upwards_trend: Inv Amt Trend_:orange[Quarterly Accumulation]:")
        InvoiceAmount_df2 = filtered_df.round(0).groupby(by = ["INVOICE_FQ","TYPE"], as_index= False)["Functional Amount(HKD)"].sum()
        # 使用pivot_table函數來重塑數據，使每個Region成為一個列
        InvoiceAmount_df2 = InvoiceAmount_df2.pivot_table(index="INVOICE_FQ", columns="TYPE", values="Functional Amount(HKD)", fill_value=0).reset_index()
        # 使用melt函數來恢復原來的長格式，並保留0值
        InvoiceAmount_df2 = InvoiceAmount_df2.melt(id_vars="INVOICE_FQ", value_name="Functional Amount(HKD)", var_name="TYPE")
        fig2 = px.line(InvoiceAmount_df2,
                       x = "INVOICE_FQ",
                       y = "Functional Amount(HKD)",
                       color='TYPE',
                       markers=True,
                       text="Functional Amount(HKD)",
                       color_discrete_map={'SPARES/OTHER': 'pink','CONTRACT_FEE': 'blue',
                                           'SERVICE_CHARGE': 'yellow','FEEDER': 'lightgreen'})
              # 更新圖表的字體大小和粗細
        fig2.update_layout(font=dict(
                    family="Arial, Arial",
                    size=12,
                    color="Black"))
        fig2.update_layout(legend=dict(orientation="h",font=dict(size=14), yanchor="bottom", y=1.02, xanchor="right", x=1))
        fig2.update_traces(marker_size=9, textposition="bottom center", texttemplate='%{text:.2s}')
        st.plotly_chart(fig2.update_layout(yaxis_showticklabels = True), use_container_width=True)
##############################################################################################################################  
with tab4:
        tab4one_column, tab4two_column= st.columns(2)
        with tab4one_column:
#All BRAND total inv amount BAR CHART
              st.subheader(":bar_chart: Invoice Amount_:orange[FY](Available to show :orange[Multiple FY]):")
              category2_df = filtered_df.round(0).groupby(by=["INVOICE_FY","BRAND"], 
                       as_index=False)["Functional Amount(HKD)"].sum().sort_values(by="Functional Amount(HKD)", ascending=False)
              df_contract_vs_invoice = px.bar(category2_df, x="INVOICE_FY", y="Functional Amount(HKD)", color="BRAND", text_auto='.3s')

# 更改顏色
              colors = {"YAMAHA": "green","HELLER": "orange","PEMTRON": "khaki","OTHERS": "lightblue"}
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

#All BRAND PIE CHART
        with tab4two_column:
             st.subheader(":round_pushpin: Invoice Percentage_:orange[Accumulate]:")
# 创建示例数据框
             brand_data = filtered_df.round(0).groupby(by=["INVOICE_FY","BRAND"],
                     as_index=False)["Functional Amount(HKD)"].sum().sort_values(by="Functional Amount(HKD)", ascending=False)         
             brandinvpie_df = pd.DataFrame(brand_data)
# 按照指定順序排序 
             brandinvpie_df["BRAND"] = brandinvpie_df["BRAND"].replace(to_replace=[x for x in brandinvpie_df["BRAND"
                                       ].unique() if x not in ["YAMAHA","HELLER", "PEMTRON", "OTHER"]], value="OTHERS")
             brandinvpie_df["BRAND"] = pd.Categorical(brandinvpie_df["BRAND"], ["YAMAHA","HELLER", "PEMTRON", "OTHER","OTHERS"])
# 创建饼状图
             df_pie = px.pie(brandinvpie_df, values="Functional Amount(HKD)", names="BRAND", color="BRAND", color_discrete_map={
                      "YAMAHA": "green","HELLER": "orange","PEMTRON": "khaki","OTHERS": "lightblue"})
# 设置字体和样式
             df_pie.update_layout(
                   font=dict(family="Arial", size=14, color="black"),
                   legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
# 显示百分比标签
             df_pie.update_traces(textposition='outside', textinfo='label+percent', marker_line_width=2,opacity=1)
# 在Streamlit中显示图表
             st.plotly_chart(df_pie, use_container_width=True)
##############################################################################################################################                   
        tab4row2one_column, tab4row2two_column= st.columns(2)
        with tab4row2one_column:
# LINE CHART of YAMAHA FY/FY
              st.divider()
              st.subheader(":chart_with_upwards_trend: :orange[YAMAHA] Inv Amt Trend_MONTHLY(Available to Show :orange[Multiple FY]):")
              df_Single_YAMAHA = filtered_df.query('BRAND == "YAMAHA"').round(0).groupby(by = ["INVOICE_FY",
                                 "INVOICE_MONTH"], as_index= False)["Functional Amount(HKD)"].sum()
# 确保 "INVOICE_FQ" 列中的所有值都出现在 df_Single_brand 中
              all_fq_invoice_values = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
              df_Single_YAMAHA = df_Single_YAMAHA.groupby(["INVOICE_FY", "INVOICE_MONTH"]).sum().reindex(pd.MultiIndex.from_product([df_Single_YAMAHA['INVOICE_FY'].unique(), all_fq_invoice_values],
                                   names=['INVOICE_FY', 'INVOICE_MONTH'])).fillna(0).reset_index()
              fig4 = go.Figure()

# 添加每个INVOICE_FY的折线
              fy_inv_values = df_Single_YAMAHA['INVOICE_FY'].unique()
              for fy_inv in fy_inv_values:
               fy_inv_data = df_Single_YAMAHA[df_Single_YAMAHA['INVOICE_FY'] == fy_inv]
               fig4.add_trace(go.Scatter(
                          x=fy_inv_data['INVOICE_MONTH'],
                          y=fy_inv_data['Functional Amount(HKD)'],
                          mode='lines+markers+text',
                          name=fy_inv,
                          text=fy_inv_data['Functional Amount(HKD)'],
                          textposition="bottom center",
                          texttemplate='%{text:.3s}',
                          hovertemplate='%{x}<br>%{y:.2f}',
                          marker=dict(size=10)))
               fig4.update_layout(xaxis=dict(
                          type='category',
                          categoryorder='array',
                          categoryarray=all_fq_invoice_values),
                          yaxis=dict(showticklabels=True),
                          font=dict(family="Arial, Arial", size=12, color="Black"),
                          hovermode='x', showlegend=True,
                          legend=dict(orientation="h",font=dict(size=14)),paper_bgcolor='lightgreen')
             
              fig4.update_layout(legend=dict(orientation="h",font=dict(size=14), yanchor="bottom", y=1.02, xanchor="right", x=1))

              st.plotly_chart(fig4.update_layout(yaxis_showticklabels = True), use_container_width=True)           
#YAMAHA Invoice Details FQ_FQ:
              pvt5 = filtered_df.query('BRAND == "YAMAHA"').round(0).pivot_table(values="Functional Amount(HKD)",
                     index=['INVOICE_FY'],columns=["INVOICE_FQ"],aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_values(by='INVOICE_FY',ascending=True)
            
              html76 = pvt5.applymap('HKD{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
              html77 = html76.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
              html78 = html77.replace('<th>Q1</th>', '<th style="background-color: lightgreen">Q1</th>')
              html79 = html78.replace('<th>Q2</th>', '<th style="background-color: lightgreen">Q2</th>')
              html80 = html79.replace('<th>Q3</th>', '<th style="background-color: lightgreen">Q3</th>')
              html81 = html80.replace('<th>Q4</th>', '<th style="background-color: lightgreen">Q4</th>')
              html822 = html81.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
# 放大pivot table
              html833 = f'<div style="zoom: 0.7;">{html822}</div>'

              st.markdown(html833, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
              csv11 = pvt5.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
              st.download_button(label='Download Table', data=csv11, file_name='YAMAHA_Sales.csv', mime='text/csv')
              st.divider()

        with tab4row2two_column:
# LINE CHART of HELLER FY/FY
              st.divider()
              st.subheader(":chart_with_upwards_trend: :orange[HELLER] Inv Amt Trend_MONTHLY(Available to Show :orange[Multiple FY)]:")
              df_Single_HELLER = filtered_df.query('BRAND == "HELLER"').round(0).groupby(by = ["INVOICE_FY","INVOICE_MONTH"], as_index= False)["Functional Amount(HKD)"].sum()
# 确保 "INVOICE_FQ" 列中的所有值都出现在 df_Single_HELLER 中
              all_fq_invoice_values = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
              df_Single_HELLER = df_Single_HELLER.groupby(["INVOICE_FY", "INVOICE_MONTH"]).sum().reindex(pd.MultiIndex.from_product([df_Single_HELLER['INVOICE_FY'].unique(), all_fq_invoice_values],
                                   names=['INVOICE_FY', 'INVOICE_MONTH'])).fillna(0).reset_index()
              fig5 = go.Figure()

# 添加每个FY_INV的折线
              fy_inv_values = df_Single_HELLER['INVOICE_FY'].unique()
              for fy_inv in fy_inv_values:
               fy_inv_data = df_Single_HELLER[df_Single_HELLER['INVOICE_FY'] == fy_inv]
               fig5.add_trace(go.Scatter(
                          x=fy_inv_data['INVOICE_MONTH'],
                          y=fy_inv_data['Functional Amount(HKD)'],
                          mode='lines+markers+text',
                          name=fy_inv,
                          text=fy_inv_data['Functional Amount(HKD)'],
                          textposition="bottom center",
                          texttemplate='%{text:.2s}',
                          hovertemplate='%{x}<br>%{y:.2f}',
                          marker=dict(size=10)))
               fig5.update_layout(xaxis=dict(
                          type='category',
                          categoryorder='array',
                          categoryarray=all_fq_invoice_values),
                          yaxis=dict(showticklabels=True),
                          font=dict(family="Arial, Arial", size=12, color="Black"),
                          hovermode='x', showlegend=True,
                          legend=dict(orientation="h",font=dict(size=14)),
                          paper_bgcolor='rgba(255,165,0,0.3)')
             
              fig5.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
              st.plotly_chart(fig5.update_layout(yaxis_showticklabels = True), use_container_width=True)

########################             
#HELLER Invoice Details FQ_FQ:
              pvt7 = filtered_df.query('BRAND == "HELLER"').round(0).pivot_table(values="Functional Amount(HKD)",index=['INVOICE_FY'],columns=["INVOICE_FQ"],
                            aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_values(by='INVOICE_FY',ascending=True)
            
              html83 = pvt7.applymap('HKD{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
              html84 = html83.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
              html85 = html84.replace('<th>Q1</th>', '<th style="background-color: orange">Q1</th>')
              html86 = html85.replace('<th>Q2</th>', '<th style="background-color: orange">Q2</th>')
              html87 = html86.replace('<th>Q3</th>', '<th style="background-color: orange">Q3</th>')
              html88 = html87.replace('<th>Q4</th>', '<th style="background-color: orange">Q4</th>')
              html89 = html88.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
# 放大pivot table
              html900 = f'<div style="zoom: 0.7;">{html89}</div>'
             
              st.markdown(html900, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
              csv12 = pvt7.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
              st.download_button(label='Download Table', data=csv12, file_name='HELLER_Sales.csv', mime='text/csv')

              st.divider()
################################################# 
        three_column, four_column= st.columns(2)
  
        with three_column:
# LINE CHART of OTHERS FY/FY
             st.subheader(":chart_with_upwards_trend: :orange[OTHERS] Inv Amt Trend_MONTHLY(Available to Show :orange[Multiple FY]):")

             df_Single_OTHERS = filtered_df.query('BRAND == "OTHERS"').round(0).groupby(by=["INVOICE_FY", "INVOICE_MONTH"],
                                as_index=False)["Functional Amount(HKD)"].sum()
# 确保 "INVOICE_FQ" 列中的所有值都出现在 df_Single_OTHERS 中
             all_fq_invoice_values = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             df_Single_OTHERS = df_Single_OTHERS.groupby(["INVOICE_FY", "INVOICE_MONTH"]).sum().reindex(pd.MultiIndex.from_product([df_Single_OTHERS['INVOICE_FY'].unique(), all_fq_invoice_values],
                               names=['INVOICE_FY', 'INVOICE_MONTH'])).fillna(0).reset_index()
             fig7 = go.Figure()

# 添加每个INVOICE_FY的折线
             fy_inv_values = df_Single_OTHERS['INVOICE_FY'].unique()
             for fy_inv in fy_inv_values:
               fy_inv_data = df_Single_OTHERS[df_Single_OTHERS['INVOICE_FY'] == fy_inv]
               fig7.add_trace(go.Scatter(
                          x=fy_inv_data['INVOICE_MONTH'],
                          y=fy_inv_data['Functional Amount(HKD)'],
                          mode='lines+markers+text',
                          name=fy_inv,
                          text=fy_inv_data['Functional Amount(HKD)'],
                          textposition="bottom center",
                          texttemplate='%{text:.3s}',
                          hovertemplate='%{x}<br>%{y:.2f}',
                          marker=dict(size=10)))
               fig7.update_layout(xaxis=dict(
                          type='category',
                          categoryorder='array',
                          categoryarray=all_fq_invoice_values),
                          yaxis=dict(showticklabels=True),
                          font=dict(family="Arial, Arial", size=12, color="Black"),
                          hovermode='x', showlegend=True,
                          legend=dict(orientation="h",font=dict(size=14)),
                          paper_bgcolor='lightblue')
              
             fig7.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
             st.plotly_chart(fig7.update_layout(yaxis_showticklabels = True), use_container_width=True)

#OTHERS Invoice Details FQ_FQ:
             pvt11 = filtered_df.query('BRAND == "OTHERS"').round(0).pivot_table(values="Functional Amount(HKD)",index=['INVOICE_FY'],columns=["INVOICE_FQ"],
                     aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_values(by='INVOICE_FY',ascending=True)

             html62 = pvt11.applymap('HKD{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
             html63 = html62.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
             html64 = html63.replace('<th>Q1</th>', '<th style="background-color: lightblue">Q1</th>')
             html65 = html64.replace('<th>Q2</th>', '<th style="background-color: lightblue">Q2</th>')
             html66 = html65.replace('<th>Q3</th>', '<th style="background-color: lightblue">Q3</th>')
             html67 = html66.replace('<th>Q4</th>', '<th style="background-color: lightblue">Q4</th>')
             html68 = html67.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')

# 放大pivot table
             html699 = f'<div style="zoom: 0.7;">{html68}</div>'
             st.markdown(html699, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
             csv9 = pvt11.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
             st.download_button(label='Download Table', data=csv9, file_name='OTHERS_Sales.csv', mime='text/csv')

             st.divider()
##################################################
        with four_column:
# LINE CHART of PEMTRON FY/FY
             st.subheader(":chart_with_upwards_trend: :orange[PEMTRON] Inv Amt Trend_MONTHLY(Available to Show :orange[Multiple FY]):")
             df_Single_PEMTRON = filtered_df.query('BRAND == "PEMTRON"').round(0).groupby(by=["INVOICE_FY", "INVOICE_MONTH"],
                                as_index=False)["Functional Amount(HKD)"].sum()
# 确保 "INVOICE_FQ" 列中的所有值都出现在 df_Single_region 中
             all_fq_invoice_values = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             df_Single_PEMTRON = df_Single_PEMTRON.groupby(["INVOICE_FY", "INVOICE_MONTH"]).sum().reindex(pd.MultiIndex.from_product([df_Single_PEMTRON['INVOICE_FY'].unique(), all_fq_invoice_values],
                               names=['INVOICE_FY', 'INVOICE_MONTH'])).fillna(0).reset_index()
             fig6 = go.Figure()

# 添加每个FY_INV的折线
             fy_inv_values = df_Single_PEMTRON['INVOICE_FY'].unique()
             for fy_inv in fy_inv_values:
               fy_inv_data = df_Single_PEMTRON[df_Single_PEMTRON['INVOICE_FY'] == fy_inv]
               fig6.add_trace(go.Scatter(
                          x=fy_inv_data['INVOICE_MONTH'],
                          y=fy_inv_data['Functional Amount(HKD)'],
                          mode='lines+markers+text',
                          name=fy_inv,
                          text=fy_inv_data['Functional Amount(HKD)'],
                          textposition="bottom center",
                          texttemplate='%{text:.3s}',
                          hovertemplate='%{x}<br>%{y:.2f}',
                          marker=dict(size=10)))
               fig6.update_layout(xaxis=dict(
                          type='category',
                          categoryorder='array',
                          categoryarray=all_fq_invoice_values),
                          yaxis=dict(showticklabels=True),
                          font=dict(family="Arial, Arial", size=12, color="Black"),
                          hovermode='x', showlegend=True,
                          legend=dict(orientation="h",font=dict(size=14)),
                          paper_bgcolor='khaki')
              
             fig6.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
             st.plotly_chart(fig6.update_layout(yaxis_showticklabels = True), use_container_width=True)
#PEMTRON Invoice Details FQ_FQ:
             pvt13 = filtered_df.query('BRAND == "PEMTRON"').round(0).pivot_table(values="Functional Amount(HKD)",index=['INVOICE_FY'],columns=["INVOICE_FQ"],
                            aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_values(by='INVOICE_FY',ascending=True)
            
             html69 = pvt13.applymap('HKD{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
             html70 = html69.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
             html71 = html70.replace('<th>Q1</th>', '<th style="background-color: khaki">Q1</th>')
             html72 = html71.replace('<th>Q2</th>', '<th style="background-color: khaki">Q2</th>')
             html73 = html72.replace('<th>Q3</th>', '<th style="background-color: khaki">Q3</th>')
             html74 = html73.replace('<th>Q4</th>', '<th style="background-color: khaki">Q4</th>')
             html75 = html74.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
# 放大pivot table
             html766 = f'<div style="zoom: 0.7;">{html75}</div>'

             st.markdown(html766, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
             csv13 = pvt13.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
             st.download_button(label='Download Table', data=csv13, file_name='PEMTRON.csv', mime='text/csv')
             st.divider()             
              
# LINE CHART of BRAND Comparision              
        st.subheader(":chart_with_upwards_trend: Inv Amt Trend_:orange[Quarterly Accumulation]:")
        InvoiceAmount_df2 = filtered_df.round(0).groupby(by = ["INVOICE_FQ","BRAND"], as_index= False)["Functional Amount(HKD)"].sum()
        # 使用pivot_table函數來重塑數據，使每個brand成為一個列
        InvoiceAmount_df2 = InvoiceAmount_df2.pivot_table(index="INVOICE_FQ", columns="BRAND", values="Functional Amount(HKD)", fill_value=0).reset_index()
        # 使用melt函數來恢復原來的長格式，並保留0值
        InvoiceAmount_df2 = InvoiceAmount_df2.melt(id_vars="INVOICE_FQ", value_name="Functional Amount(HKD)", var_name="BRAND")
        fig2 = px.line(InvoiceAmount_df2,
                       x = "INVOICE_FQ",
                       y = "Functional Amount(HKD)",
                       color='BRAND',
                       markers=True,
                       text="Functional Amount(HKD)",
                       color_discrete_map={'YAMAHA': 'green','OTHERS': 'lightblue',
                                           'HELLER': 'orange','PEMTRON': 'Khaki'})
              # 更新圖表的字體大小和粗細
        fig2.update_layout(font=dict(
                    family="Arial, Arial",
                    size=12,
                    color="Black"))
        fig2.update_layout(legend=dict(orientation="h",font=dict(size=14), yanchor="bottom", y=1.02, xanchor="right", x=1))
        fig2.update_traces(marker_size=9, textposition="bottom center", texttemplate='%{text:.2s}')
        st.plotly_chart(fig2.update_layout(yaxis_showticklabels = True), use_container_width=True)

#############
with tab5:
#Top Down Customer details Table
      left_column, right_column= st.columns(2)
#BAR CHART Customer List
      with left_column:
             st.subheader(":radio: Top 20 Customer_:blue[Inv Amt]:")            
             customer_line = (filtered_df.groupby(
                             by=["Enduser"])[["Functional Amount(HKD)"]].sum().sort_values(by="Functional Amount(HKD)", ascending=False).head(20))
# 生成颜色梯度
             colors = px.colors.sequential.Blues[::-1]  # 将颜色顺序反转为从深到浅
# 创建条形图
             fig_customer = px.bar(
                  customer_line,
                  x="Functional Amount(HKD)",
                  y=customer_line.index,
                  text="Functional Amount(HKD)",
                  orientation="h",
                  color=customer_line.index,
                  color_discrete_sequence=colors[:len(customer_line)],
                  template="plotly_white", text_auto='.3s')

# 更新图表布局和样式
             fig_customer.update_layout(
                  height=1000,
                  yaxis=dict(title="Enduser"),
                  xaxis=dict(title="Functional Amount(HKD)"),)
             fig_customer.update_layout(font=dict(family="Arial", size=15))
             fig_customer.update_traces(
                  textposition="inside",
                  marker_line_color="black",
                  marker_line_width=2,
                  opacity=1,showlegend=False,
                  )
             # 显示图表
             st.plotly_chart(fig_customer, use_container_width=True)

# PIE CHART Customer Type
with right_column:
    st.subheader(":radio: Top 20 Customer Type Percentage:")
    customer_type_data = (filtered_df.groupby(by=["Enduser", "Type"])[["Functional Amount(HKD)"]].sum().reset_index().sort_values(by="Functional Amount(HKD)", ascending=False).head(20))
    customer_type_data["Percentage"] = (customer_type_data["Functional Amount(HKD)"] / customer_type_data.groupby("Enduser")["Functional Amount(HKD)"].transform("sum")) * 100
    customer_type_data = customer_type_data.groupby("Enduser").apply(lambda x: x.sort_values(by="Percentage", ascending=False).head(5)).reset_index(drop=True)
    fig_customer_type = px.pie(customer_type_data, values="Percentage", names="Type", color="Type", hole=.3, template="plotly_white")
    fig_customer_type.update_traces(textposition='inside', textinfo='percent+label')
    fig_customer_type.update_layout(font=dict(family="Arial", size=15))
    st.plotly_chart(fig_customer_type, use_container_width=True)






# Top 20 Customer 買什麼type/ BRAND   
#👇 Slider for Sales Amount
