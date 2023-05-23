import pandas as pd
import streamlit as st
import os
import re
import openpyxl

a = openpyxl.Workbook()
del a

def read_data(file):
    # st.write(file.name)
    # xxxx. pattern in file.name:
    if not re.match(r'\d{4}.', file.name):
        st.write('文件名格式不对')
        return
    df = pd.read_excel(file)
    df = df.drop([0, 1, 2])
    df = df.drop(df.index[len(df)-1])
    df = df.T
    new_header = df.iloc[0]
    df = df[1:]
    df.columns = new_header
    # drop empty rows
    df = df.dropna(axis=0, how='all')
    df = df.fillna(0)
    df.head(10)
    data = {}
    index = ''
    for i, row in df.iterrows():
        if not i.startswith('Unnamed'):
            index = i
            data[index] = {}
            for j, col in row.items():
                data[index][j] = 0
        else:
            for j, col in row.items():
                try:
                    data[index][j] += col
                except:
                    pass
    new_data = {}
    regions = ['ZL条数', '崆峒', '泾川', '灵台', '崇信', '华亭', '庄浪', '静宁']
    for k, v in data.items():
        k = file.name.split('.')[0] + '-' + k.replace('日期：', '').replace('年', '-').replace('月', '-').replace('日', '')
        new_data[k] = {region: 0 for region in regions}
        for k2, v2 in v.items():
            if k2 == '省上下发指令条数':
                new_data[k]['ZL条数'] = v2
                continue
            if '平凉' in k2:
                new_data[k]['崆峒'] += v2
            for region in regions:
                if region in k2:
                    new_data[k][region] += v2
                    break
    new_data = pd.DataFrame(new_data).T
    # make index new column
    new_data['日期'] = new_data.index
    # # add total column to right
    # new_data['总计'] = new_data.sum(axis=1)
    # add file name column to bottom
    # new_data = new_data.append(pd.Series(name='文件名', dtype='object'))
    # new_data.iloc[-1, :] = file.name


    return new_data

with st.form(key='read_data'):
    file_path = st.file_uploader(label='选择文件', accept_multiple_files=False)
    submit_button = st.form_submit_button(label='展示')
    if submit_button:
        data = read_data(file_path)
        if os.path.exists('data.xlsx'):
            local_data = pd.read_excel('data.xlsx')
            # st.write(local_data)
            data = pd.concat([local_data, data], axis=0)


        # sort with 日期
        data = data.sort_values(by='日期', ascending=True)
        # drop duplicate
        data = data.drop_duplicates(subset=['日期'], keep='last')
        data.to_excel('data.xlsx', index=False)
        


with st.form(key='show_data'):

    d1 = st.date_input('开始日期')
    d2 = st.date_input('结束日期')

    # convert to datetime
    d1 = pd.to_datetime(d1)
    d2 = pd.to_datetime(d2)

    
    
    submit_button = st.form_submit_button(label='查询')
    if submit_button:
        data = pd.read_excel('data.xlsx')
        data['日期'] = pd.to_datetime(data['日期'])
        data = data[(data['日期'] >= d1) & (data['日期'] <= d2)]
        st.write(data)
