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

st.set_page_config(layout="wide")
st.title("엑셀 텍스트 편집기")

uploaded_file = st.file_uploader("엑셀 파일 업로드", type=["xlsx", "xls"])

# apply 버튼 눌림 정보
st.session_state.numbered = False

# 텍스트 입력 및 출력
def number_reload():
    st.session_state.numbered = True
        

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file).fillna("")
    # df_to_download = pd.read_excel(uploaded_file).fillna("")

    if "df_to_download" not in st.session_state:
        st.session_state.df_to_download = pd.read_excel(uploaded_file).fillna("")

    number = st.number_input('Insert a row number', min_value=1, step = 1, max_value = len(df.values), on_change = number_reload)


    
    
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

    # print(opening, main1, main2, main3, main4, main5, closing, pass_value, not_matched_value)
    

    # try:

    key_dict = {2:opening, 3:main1, 4:main2, 5:main3, 6:main4, 7:main5, 8:closing, 9:pass_value, 10:not_matched_value}


    col_buttons = st.columns(3)
    col_ls = st.columns(2)

    # 텍스트 입력 및 출력
    def reload(key):
        if st.session_state.numbered == True:
            st.session_state[key] = key_dict[key]
        else:
            st.session_state.numbered = False
    
    
    
    input_text0 = col_ls[0].text_area('**원문**', value = origin, placeholder='please copy and paste' ,height=1500, key = 1)

    opening_on = col_ls[1].checkbox('**인사말 on:ballot_box_with_check:**', key = 21, value = False)
    if opening != '' or opening_on == True:
        input_text1 = col_ls[1].text_area('**인사말:wave:**', value = opening, placeholder='please copy and paste', height=100, key = 2, on_change=reload, args = [2])
    else:
        input_text1 = ''

    main1_on = col_ls[1].checkbox('**본문1 on:ballot_box_with_check:**', key = 31, value = False)
    if main1 != '' or main1_on == True:
        input_text2 = col_ls[1].text_area('**본문:one:**', value = main1, placeholder='please copy and paste', height=100, key = 3, on_change=reload, args = [3])
    else:
        input_text2 = ''

    main2_on = col_ls[1].checkbox('**본문2 on:ballot_box_with_check:**', key = 41, value = False)
    if main2 != '' or main2_on == True:
        input_text3 = col_ls[1].text_area('**본문:two:**', value = main2, placeholder='please copy and paste', height=100, key = 4, on_change=reload, args = [4])
    else:
        input_text3 = ''


    main3_on = col_ls[1].checkbox('**본문3 on:ballot_box_with_check:**', key = 51, value = False)
    if main3 != '' or main3_on == True:
        input_text4 = col_ls[1].text_area('**본문:three:**', value = main3, placeholder='please copy and paste', height=100, key = 5, on_change=reload, args = [5])
    else:
        input_text4 = ''
    
    main4_on = col_ls[1].checkbox('**본문4 on:ballot_box_with_check:**', key = 61, value = False)
    if main4 != '' or main4_on == True:
        input_text5 = col_ls[1].text_area('**본문:four:**', value = main4, placeholder='please copy and paste', height=100, key = 6, on_change=reload, args = [6])
    else:
        input_text5 = ''
    
    main5_on = col_ls[1].checkbox('**본문5 on:ballot_box_with_check:**', key = 71, value = False)
    if main5 != '' or main5_on == True:
        input_text6 = col_ls[1].text_area('**본문:five:**', value = main5, placeholder='please copy and paste', height=100, key = 7, on_change=reload, args = [7])
    else:
        input_text6 = ''

    closing_on = col_ls[1].checkbox('**맺음말 on:ballot_box_with_check:**', key = 81, value = False)
    if closing != '' or closing_on == True:
        input_text7 = col_ls[1].text_area('**맺음말:end:**', value = closing, placeholder='please copy and paste', height=100, key = 8, on_change=reload, args = [8])
    else:
        input_text7 = ''

    pass_on = col_ls[1].checkbox('**PASS on:ballot_box_with_check:**', key = 91, value = False)
    if pass_value != '' or pass_on == True:
        to_pass = col_ls[1].text_area('**PASS:parking:**', value = pass_value, height=10, key = 9, on_change=reload, args = [9])
    else:
        to_pass = ''
    
    not_matched_on = col_ls[1].checkbox('**not_matched on:ballot_box_with_check:**', key = 101, value = False)
    if not_matched_value != '' or not_matched_on == True:
        not_matched = col_ls[1].text_area('**본문-맺음말 안 맞음:no_entry_sign:**', value = not_matched_value, height=10, key = 10, on_change=reload, args = [10])
    else:
        not_matched = ''

    # col_ls[1].session_state[91] = False
    
    
    df_xlsx = to_excel(st.session_state.df_to_download)
    col_buttons[2].download_button(label='📥 Download Current Result',
                        data=df_xlsx ,
                        file_name= '{input}_수정.xlsx'.format(input = uploaded_file.name))
    
    
    if col_buttons[0].button('Apply'):

        st.session_state.applied = True

        # key_dict = {2:input_text1, 3:input_text2, 4:input_text3, 5:input_text4, 6:input_text5, 7:input_text6, 8:input_text7, 9:to_pass, 10:not_matched}
        
        st.session_state.df_to_download['원문'][number-1] = input_text0
        st.session_state.df_to_download['인사말'][number-1] = input_text1
        st.session_state.df_to_download['본문1'][number-1] = input_text2
        st.session_state.df_to_download['본문2'][number-1] = input_text3
        st.session_state.df_to_download['본문3'][number-1] = input_text4
        st.session_state.df_to_download['본문4'][number-1] = input_text5
        st.session_state.df_to_download['본문5'][number-1] = input_text6
        st.session_state.df_to_download['맺음말'][number-1] = input_text7
        st.session_state.df_to_download['PASS'][number-1] = to_pass
        st.session_state.df_to_download['본문-맺음말 안 맞음'][number-1] = not_matched_value



    if col_buttons[1].button('Get Origin'):
        # 텍스트 입력 및 출력
        input_text0 = col_ls[0].text_area
    
        input_text1 = col_ls[1].text_area('인사말', value = opening, height=100, key = 12)

        input_text2 = col_ls[1].text_area('본문1', value = main1, height=100, key = 13)

        input_text3 = col_ls[1].text_area('본문2', value = main2, height=100, key = 14)

        input_text4 = col_ls[1].text_area('본문3', value = main3, height=100, key = 15)

        input_text5 = col_ls[1].text_area('본문4', value = main4, height=100, key = 16)
        
        input_text6 = col_ls[1].text_area('본문5', value = main5, height=100, key = 17)

        input_text7 = col_ls[1].text_area('맺음말', value = closing, height=100, key = 18)

        to_pass = col_ls[1].text_area('PASS', value = pass_value, height=10, key = 19)
        not_matched = col_ls[1].text_area('본문-맺음말 안 맞음', value = not_matched_value, height=10, key = 110)
        
    # except Exception as e:
    #     st.write("오류 발생:", e)



