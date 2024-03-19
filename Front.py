#streamlit run streamlit.py --server.port=80
import streamlit as st
import pandas as pd
import numpy as np
import time, datetime
from PIL import Image
from streamlit_autorefresh import st_autorefresh




#Шапка сайта, логотип, имя
image_directory = "./Config/frontlogo.PNG"
image = Image.open(image_directory)
st.set_page_config(page_title = "GranatBioTech - DOPC (ﾉ◕ヮ◕)ﾉ*‿︵‿･ﾟ✧", 
                   page_icon = image,
                   layout="wide",
                   initial_sidebar_state="expanded")
st.sidebar.image('./Config/logo.png')
st.sidebar.title("Меню")


#Автообновление виджетов раз в 30 секунд
st_autorefresh(interval=1 * 30 * 1000)

#Чтение данных цилиндров
try:
    dataEVFront = pd.read_excel('./log/EVfront.xlsx')
    def load_dataEVFront():
            dfEVFront = pd.DataFrame(dataEVFront, columns=[options])
            load_datafront = dfEVFront.dropna(how='all')    
            data = load_datafront[load_datafront>0.015].dropna()   
            return data
except:
    pass

#Вычисление средних значений для цилиндров
try:    
    def load_dataMean():
            dfEVFront = pd.DataFrame(dataEVFront, columns=[options])     
            data = dfEVFront[dfEVFront>0.015].dropna()
            dataMean = data.mean()  
            return dataMean
except:
    pass

#Вычисление среднего значения за прошлый день
try:
    def mean_lastday():
        today = datetime.datetime.now() #Получение текущих даты и времени.
        nowYear = today.year    #Время.
        nowMonth= today.month 
        if nowMonth <10:
            nowMonth = '0'+str(nowMonth)
        lastDay = today.day - 1   #Дата.
        if lastDay == 0:
            lastDay = 1
        if lastDay < 10:
            lastDay = '0'+str(lastDay)
        subString = "./log/front_log_"+ str(nowYear) +"-"+ str(nowMonth) + "-" + str(lastDay)+".xlsx"
        dataEVFrontMean = pd.read_excel(subString)
        dfEVFrontMean = pd.DataFrame(dataEVFrontMean, columns=[options])
        data = dfEVFrontMean[dfEVFrontMean>0.015].dropna()      
        frontMean = data[options].mean()
        return frontMean
except:
    pass

#Метрики и таблица 1 шасси - Проценты брака
try:
    dataReject1Ch = pd.read_excel('./Config/Reject.xlsx',sheet_name='Шасси1')
    def load_dataReject1ChMetric():
        dfReject1ch = pd.DataFrame(dataReject1Ch,index=[0,1,2,3,4,5,6,7], columns=['Station','Tested parts','Reject parts','%'])   
        return dfReject1ch
except:
    pass

try:
    dataReject1Ch = pd.read_excel('./Config/Reject.xlsx',sheet_name='Шасси1')
    def load_dataReject1ChTable():
        dfReject1ch = pd.DataFrame(dataReject1Ch,index=[0,1,2,3,4,5,6,7], columns=['Station','Tested parts','Reject parts','%'])   
        dfReject1ch['%'] = dfReject1ch['%'].map('{:,.2f}'.format)
        return dfReject1ch
except:
    pass

#Метрики и таблица 2 шасси - Проценты брака
try:
    dataReject2Ch = pd.read_excel('./Config/Reject.xlsx',sheet_name='Шасси2')
    def load_dataReject2ChMetric():
        dfReject2ch = pd.DataFrame(dataReject2Ch,index=[0,1,2,3,4,5], columns=['Station','Tested parts','Reject parts','%'])   
        return dfReject2ch
except:
    pass

try:
    dataReject2Ch = pd.read_excel('./Config/Reject.xlsx',sheet_name='Шасси2')
    def load_dataReject2ChTable():
        dfReject2ch = pd.DataFrame(dataReject2Ch,index=[0,1,2,3,4,5], columns=['Station','Tested parts','Reject parts','%'])   
        dfReject2ch['%'] = dfReject2ch['%'].map('{:,.2f}'.format)
        return dfReject2ch
except:
    pass
#Метрики статистика по браку
try:
    RejectStat = pd.read_excel('./Config/Reject.xlsx',sheet_name='СтатистикаПоБраку')
    def load_dataRejectStatMetrik():
        dfRejectStat = pd.DataFrame(RejectStat,index=[0,1,2], columns=['name','parts'])   
        return dfRejectStat
except:
    pass


try:
    def history_ch1(MinDay):
        today = datetime.datetime.now()
        nowYear = today.year   
        nowMonth= today.month 
        if nowMonth <10:
            nowMonth = '0'+str(nowMonth)
        minOneDay = today.day - MinDay
        if minOneDay == 0:
            minOneDay = 1
        if minOneDay < 10:
            minOneDay = '0'+str(minOneDay)
        subString = "./log/front_log_"+ str(nowYear) +"-"+ str(nowMonth) + "-" + str(minOneDay)+".xlsx"
        HistoryFrontData = pd.read_excel(subString)
        DFHistoryFrontData = pd.DataFrame(HistoryFrontData, columns=[optionsHistory])
        data = DFHistoryFrontData[DFHistoryFrontData>0.015].dropna()
        DfHistory = data.dropna(how='all') 
        return DfHistory
except:
    pass

def dateHistory(MinusDay):
    today = datetime.datetime.now()
    nowYear = today.year   
    nowMonth= today.month 
    if nowMonth <10:
        nowMonth = '0'+str(nowMonth)
    minOneDay = today.day - MinusDay
    if minOneDay == 0:
        minOneDay = 1
    if minOneDay < 10:
        minOneDay = '0'+str(minOneDay)
    subString = str(nowYear) +"-"+ str(nowMonth) + "-" + str(minOneDay)+".xlsx"
    return subString




#Выбор пользователя
with st.sidebar.form(key='my_form'):
    username = st.selectbox('Пользователь',
    ('Оператор','Инженер'))
    st.form_submit_button('Применить')
    st.experimental_user = username



if st.experimental_user =='Инженер':
    
    
    actual = st.sidebar.checkbox("Актуальные данные")
    if actual == True:
        options = st.selectbox(
        'Выберите EV!',
        ('EV01_02',	'EV01_03',	'EV01_04',	'EV02_01',	'EV02_02',	'EV03_11',	'EV03_12',	'EV03_13', 'EV03_14',	'EV04_11',	'EV04_15',	'EV06_01',	'EV06_02',	'EV08_01',	'EV10_11',	'EV10_12',	'EV10_13','EV10_14',	'EV12_01',	'EV14_01',	'EV15_03',	'EV15_04',	'EV15_05',	'EV20_01',	'EV20_02',	'EV20_04',	'EV21_02',	'EV23_01',	'EV23_04',	'EV25_04',	'EV25_05',	'EV25_06',	'EV29_03',	'EV29_04',	'EV29_05',	'EV35_11',	'EV35_12',	'EV36_01',	'EV36_02',	'EV36_04',	'EV39_01',	'EV39_02',	'EV41_01'))
        st.subheader('График работы цилиндра')
        with st.empty():
            col1, col2 = st.columns(2,gap='large')
            with col1:
                try:
                    dfEVFront = load_dataEVFront()
                    chart = st.line_chart(dfEVFront)
                except:
                    pass
            with col2:
                try:
                    dfEVFront = load_dataEVFront()
                    st.write(dfEVFront)
                except:
                    pass
            with col1:
                try:
                    actual_mean = load_dataMean()
                    mean = actual_mean[options] 
                    lastday_mean = mean_lastday()
                    st.metric('Среднее значение',"{:.3f}".format(mean),"{:.3f}".format(lastday_mean),delta_color='off',help='Нижняя строка, показатель среднего значения за прошлый день')
                except:
                    pass
        
    HistoryCB1 = st.sidebar.checkbox(dateHistory(1),key="H1")
    if HistoryCB1 == True:
        st.subheader(dateHistory(1))
        optionsHistory = st.selectbox(
        'Выберите EV!',
        ('EV01_02',	'EV01_03',	'EV01_04',	'EV02_01',	'EV02_02',	'EV03_11',	'EV03_12',	'EV03_13', 'EV03_14',	'EV04_11',	'EV04_15',	'EV06_01',	'EV06_02',	'EV08_01',	'EV10_11',	'EV10_12',	'EV10_13','EV10_14',	'EV12_01',	'EV14_01',	'EV15_03',	'EV15_04',	'EV15_05',	'EV20_01',	'EV20_02',	'EV20_04',	'EV21_02',	'EV23_01',	'EV23_04',	'EV25_04',	'EV25_05',	'EV25_06',	'EV29_03',	'EV29_04',	'EV29_05',	'EV35_11',	'EV35_12',	'EV36_01',	'EV36_02',	'EV36_04',	'EV39_01',	'EV39_02',	'EV41_01'),key='history')
        st.subheader('График работы цилиндра')
        col1, col2 = st.columns(2,gap='large')
        with st.empty():
            with col1:
                dfEVFront = history_ch1(1)
                chart = st.line_chart(dfEVFront)
            with col2:
                dfEVFront = history_ch1(1)
                st.write(dfEVFront)
            with col1:
                VarMean = dfEVFront.mean()
                var = VarMean[0]
                st.metric('Среднее значение',"{:.3f}".format(var))

    HistoryCB2 = st.sidebar.checkbox(dateHistory(2),key="H2")
    if HistoryCB2 == True:
        st.subheader(dateHistory(2))
        optionsHistory = st.selectbox(
        'Выберите EV!',
        ('EV01_02',	'EV01_03',	'EV01_04',	'EV02_01',	'EV02_02',	'EV03_11',	'EV03_12',	'EV03_13', 'EV03_14',	'EV04_11',	'EV04_15',	'EV06_01',	'EV06_02',	'EV08_01',	'EV10_11',	'EV10_12',	'EV10_13','EV10_14',	'EV12_01',	'EV14_01',	'EV15_03',	'EV15_04',	'EV15_05',	'EV20_01',	'EV20_02',	'EV20_04',	'EV21_02',	'EV23_01',	'EV23_04',	'EV25_04',	'EV25_05',	'EV25_06',	'EV29_03',	'EV29_04',	'EV29_05',	'EV35_11',	'EV35_12',	'EV36_01',	'EV36_02',	'EV36_04',	'EV39_01',	'EV39_02',	'EV41_01'),key='history1')
        st.subheader('График работы цилиндра')
        col1, col2 = st.columns(2,gap='large')
        with st.empty():
            with col1:
                dfEVFront = history_ch1(2)
                chart = st.line_chart(dfEVFront)
            with col2:
                dfEVFront = history_ch1(2)
                st.write(dfEVFront)
            with col1:
                VarMean = dfEVFront.mean()
                var = VarMean[0]
                st.metric('Среднее значение',"{:.3f}".format(var))

    HistoryCB3 = st.sidebar.checkbox(dateHistory(3),key="H3")
    if HistoryCB3 == True:
        st.subheader(dateHistory(3))
        optionsHistory = st.selectbox(
        'Выберите EV!',
        ('EV01_02',	'EV01_03',	'EV01_04',	'EV02_01',	'EV02_02',	'EV03_11',	'EV03_12',	'EV03_13', 'EV03_14',	'EV04_11',	'EV04_15',	'EV06_01',	'EV06_02',	'EV08_01',	'EV10_11',	'EV10_12',	'EV10_13','EV10_14',	'EV12_01',	'EV14_01',	'EV15_03',	'EV15_04',	'EV15_05',	'EV20_01',	'EV20_02',	'EV20_04',	'EV21_02',	'EV23_01',	'EV23_04',	'EV25_04',	'EV25_05',	'EV25_06',	'EV29_03',	'EV29_04',	'EV29_05',	'EV35_11',	'EV35_12',	'EV36_01',	'EV36_02',	'EV36_04',	'EV39_01',	'EV39_02',	'EV41_01'),key='history2')
        st.subheader('График работы цилиндра')
        col1, col2 = st.columns(2,gap='large')
        with st.empty():
            with col1:
                dfEVFront = history_ch1(3)
                chart = st.line_chart(dfEVFront)
            with col2:
                dfEVFront = history_ch1(3)
                st.write(dfEVFront)
            with col1:
                VarMean = dfEVFront.mean()
                var = VarMean[0]
                st.metric('Среднее значение',"{:.3f}".format(var))

if st.experimental_user =='Оператор':
    
    
    try:
        dfRejectMetric1Ch = load_dataReject1ChMetric()
        dfRejectTable1Ch = load_dataReject1ChTable()
        
    except:        
        pass
    try:
        dfRejectStatMetrik = load_dataRejectStatMetrik()
    except:
        pass
    RejectStatMetrik = st.sidebar.checkbox('Статистика по продукции',key='RejectStat')
    if RejectStatMetrik == True:
        st.header('Статистика по продукции')
        try:
            RS1metrik = dfRejectStatMetrik['parts'].loc[dfRejectStatMetrik.index[0]]
            RS2metrik = dfRejectStatMetrik['parts'].loc[dfRejectStatMetrik.index[1]]
            RS3metrik = dfRejectStatMetrik['parts'].loc[dfRejectStatMetrik.index[2]]

            col1,col2,col3 = st.columns(3)
            col1.metric("Кол-во HUB'ов",str(RS1metrik))
            col2.metric("Кол-во Канюль",str(RS2metrik))
            col3.metric("Выпущенных игл",str(RS3metrik))
        except:
            pass

    CH1RejectsMetrik = st.sidebar.checkbox('Проценты брака Шасси 1 - Метрики')
    if CH1RejectsMetrik == True:
        st.header('Шасси 1 - Метрики')
        try:
            ch1st02metrik = dfRejectMetric1Ch['%'].loc[dfRejectMetric1Ch.index[0]]
            ch1st06metrik = dfRejectMetric1Ch['%'].loc[dfRejectMetric1Ch.index[1]]
            ch1st14metrik = dfRejectMetric1Ch['%'].loc[dfRejectMetric1Ch.index[2]]
            ch1st19metrik = dfRejectMetric1Ch['%'].loc[dfRejectMetric1Ch.index[3]]
            ch1st27metrik = dfRejectMetric1Ch['%'].loc[dfRejectMetric1Ch.index[4]]
            ch1st29metrik = dfRejectMetric1Ch['%'].loc[dfRejectMetric1Ch.index[5]]
            ch1st32metrik = dfRejectMetric1Ch['%'].loc[dfRejectMetric1Ch.index[6]]
            ch1st33metrik = dfRejectMetric1Ch['%'].loc[dfRejectMetric1Ch.index[7]]

            col1, col2, col3, col4, col5,col6,col7,col8 = st.columns(8)
            col1.metric("ST02 - HUB PRESENCE CHECK", str(ch1st02metrik) +'%', "{:.2f}".format(ch1st02metrik- float(2)) ,delta_color='inverse')
            col2.metric("ST06 -POREX CHECK", str(ch1st06metrik) +'%', "{:.2f}".format(ch1st06metrik- float(2)),delta_color='inverse')
            col3.metric("ST14 - PRESENCE CHECK", str(ch1st14metrik) +'%', "{:.2f}".format(ch1st14metrik- float(2)),delta_color='inverse')
            col4.metric("ST19 - NEEDLE PRESENCE CHECK", str(ch1st19metrik) +'%', "{:.2f}".format(ch1st19metrik- float(2)),delta_color='inverse')
            col5.metric("ST27 - GLUE VISION CHECK", str(ch1st27metrik) +'%', "{:.2f}".format(ch1st27metrik- float(2)),delta_color='inverse')
            col6.metric("ST29 - PULL-TEST CHECK", str(ch1st29metrik) +'%', "{:.2f}".format(ch1st29metrik- float(2)),delta_color='inverse')
            col7.metric("ST32 - TIP VISION CHECK", str(ch1st32metrik) +'%', "{:.2f}".format(ch1st32metrik- float(2)),delta_color='inverse')
            col8.metric("ST33 - SLANTING AND HEIGHT VISION CHECK",str(ch1st33metrik) +'%', "{:.2f}".format(ch1st33metrik- float(2)),delta_color='inverse')
        except:
            pass
    CH1RejectsSheet = st.sidebar.checkbox('Проценты брака Шасси 1 - Таблица')
    if CH1RejectsSheet == True:
        try:
            st.header('Шасси 1 - Таблица')   
            st.table(dfRejectTable1Ch)
        except:
            pass
    

    try:
        dfRejectMetric2Ch = load_dataReject2ChMetric()
        dfRejectTable2Ch = load_dataReject2ChTable()
    except:        
        pass
    CH2RejectsMetrik = st.sidebar.checkbox('Проценты брака Шасси 2 - Метрики',key='ch2')
    if CH2RejectsMetrik == True:
        st.header('Шасси 2 - Метрики')
        try:
            ch2st02metrik = dfRejectMetric2Ch['%'].loc[dfRejectMetric2Ch.index[0]]
            ch2st06metrik = dfRejectMetric2Ch['%'].loc[dfRejectMetric2Ch.index[1]]
            ch2st14metrik = dfRejectMetric2Ch['%'].loc[dfRejectMetric2Ch.index[2]]
            ch2st19metrik = dfRejectMetric2Ch['%'].loc[dfRejectMetric2Ch.index[3]]
            ch2st27metrik = dfRejectMetric2Ch['%'].loc[dfRejectMetric2Ch.index[4]]
            ch2st29metrik = dfRejectMetric2Ch['%'].loc[dfRejectMetric2Ch.index[5]]
            

            col1, col2, col3, col4, col5,col6 = st.columns(6)
            col1.metric("ST03 - SUB-ASSEMBLY PRESENCE CHECK", str(ch2st02metrik) +'%', "{:.2f}".format(ch2st02metrik- float(2)) ,delta_color='inverse')
            col2.metric("ST09 - OCCULUSION TEST CHECK", str(ch2st06metrik) +'%', "{:.2f}".format(ch2st06metrik- float(2)),delta_color='inverse')
            col3.metric("ST14 - SLEEVE VISION CHECK", str(ch2st14metrik) +'%', "{:.2f}".format(ch2st14metrik- float(2)),delta_color='inverse')
            col4.metric("ST24 - PROTECTIVE SHEAT PRESENCE CHECK", str(ch2st19metrik) +'%', "{:.2f}".format(ch2st19metrik- float(2)),delta_color='inverse')
            col5.metric("ST26 - SUB ASSEMBLY PRESENCE CHECK", str(ch2st27metrik) +'%', "{:.2f}".format(ch2st27metrik- float(2)),delta_color='inverse')
            col6.metric("ST29 - CAP PRESENCE CHECK", str(ch2st29metrik) +'%', "{:.2f}".format(ch2st29metrik- float(2)),delta_color='inverse')
            
        except:
            pass
    CH2RejectsSheet = st.sidebar.checkbox('Проценты брака Шасси 2 - Таблица')
    if CH2RejectsSheet == True:
        try:
            st.header('Шасси 2 - Таблица')   
            st.table(dfRejectTable2Ch)
        except:
            pass
    