import streamlit as st
import pandas as pd
from io import BytesIO

# excel 파일 읽어오기

# header 가져오기
# 행 가져오기

@st.cache_data
def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0.00'}) 
    worksheet.set_column('A:A', None, format1)  
    writer.close()
    # writer.save()
    processed_data = output.getvalue()
    return processed_data


st.title("엑셀 텍스트 편집기")

uploaded_file = st.file_uploader("엑셀 파일 업로드", type=["xlsx", "xls"])



if uploaded_file is not None:
    df = pd.read_excel(uploaded_file).fillna("")
    # df_to_download = pd.read_excel(uploaded_file).fillna("")

    if "df_to_download" not in st.session_state:
        st.session_state.df_to_download = pd.read_excel(uploaded_file).fillna("")

    number = st.number_input('Insert a row number', min_value=1, step = 1)

    origin = st.session_state.df_to_download['원문'][number-1].strip()
    opening = str(st.session_state.df_to_download['인사말'][number-1]).strip()
    main1 = str(st.session_state.df_to_download['본문1'][number-1]).strip()
    main2 = str(st.session_state.df_to_download['본문2'][number-1]).strip()
    main3 = str(st.session_state.df_to_download['본문3'][number-1]).strip()
    main4 = str(st.session_state.df_to_download['본문4'][number-1]).strip()
    main5 = str(st.session_state.df_to_download['본문5'][number-1]).strip()
    closing = str(st.session_state.df_to_download['맺음말'][number-1]).strip()
    pass_value = str(st.session_state.df_to_download['PASS'][number-1]).strip()
    not_matched_value = str(st.session_state.df_to_download['본문-맺음말 안 맞음'][number-1]).strip()


    try:


        col_buttons = st.columns(3)
        col_ls = st.columns(2)

        # 텍스트 입력 및 출력
        with col_ls[0]:
            input_text0 = col_ls[0].text_area('원문', value = origin, height=1500, key = 1)
        with col_ls[1]:
            input_text1 = col_ls[1].text_area('인사말', value = opening, height=100, key = 2)

            input_text2 = col_ls[1].text_area('본문1', value = main1, height=100, key = 3)

            input_text3 = col_ls[1].text_area('본문2', value = main2, height=100, key = 4)

            input_text4 = col_ls[1].text_area('본문3', value = main3, height=100, key = 5)

            input_text5 = col_ls[1].text_area('본문4', value = main4, height=100, key = 6)
            
            input_text6 = col_ls[1].text_area('본문5', value = main5, height=100, key = 7)

            input_text7 = col_ls[1].text_area('맺음말', value = closing, height=100, key = 8)

            to_pass = col_ls[1].text_area('PASS', value = pass_value, height=10, key = 9)
            not_matched = col_ls[1].text_area('본문-맺음말 안 맞음', value = not_matched_value, height=10, key = 10)

        
        df_xlsx = to_excel(st.session_state.df_to_download)
        col_buttons[2].download_button(label='📥 Download Current Result',
                            data=df_xlsx ,
                            file_name= '{input}_수정.xlsx'.format(input = uploaded_file.name))
        
        
        if col_buttons[0].button('Apply'):
            
            st.session_state.df_to_download['원문'][number-1] = input_text0
            st.session_state.df_to_download['인사말'][number-1] = input_text1
            st.session_state.df_to_download['본문1'][number-1] = input_text2
            st.session_state.df_to_download['본문2'][number-1] = input_text3
            st.session_state.df_to_download['본문3'][number-1] = input_text4
            st.session_state.df_to_download['본문4'][number-1] = input_text5
            st.session_state.df_to_download['본문5'][number-1] = input_text6
            st.session_state.df_to_download['맺음말'][number-1] = input_text7
            st.session_state.df_to_download['PASS'][number-1] = to_pass
            st.session_state.df_to_download['본문-맺음말 안 맞음'][number-1] = not_matched



        if col_buttons[1].button('Get Origin'):
            # 텍스트 입력 및 출력
            input_text0 = col_ls[0].text_area
        
            input_text1 = col_ls[1].text_area('인사말', value = df['인사말'][number-1], height=100, key = 2)

            input_text2 = col_ls[1].text_area('본문1', value = main1, height=100, key = 3)

            input_text3 = col_ls[1].text_area('본문2', value = main2, height=100, key = 4)

            input_text4 = col_ls[1].text_area('본문3', value = main3, height=100, key = 5)

            input_text5 = col_ls[1].text_area('본문4', value = main4, height=100, key = 6)
            
            input_text6 = col_ls[1].text_area('본문5', value = main5, height=100, key = 7)

            input_text7 = col_ls[1].text_area('맺음말', value = closing, height=100, key = 8)

            to_pass = col_ls[1].text_area('PASS', value = pass_value, height=10, key = 9)
            not_matched = col_ls[1].text_area('본문-맺음말 안 맞음', value = not_matched_value, height=10, key = 10)
        
    except Exception as e:
        st.write("오류 발생:", e)



