import pandas as pd
import numpy as np
import xlsxwriter
import datetime
import pandas_ta as ta


######MA 20#################

ind_par = 20 #paramert indykatora

data = pd.ExcelFile('./input/EM.xlsx').parse('Arkusz1')


#Oblicz wartości wskaźnika
data['Indicator']=data['Close'].rolling(ind_par).mean()

#######################Przelicz transakcje#######################################

#format = '%Y-%m-%d %H:%M:%S'
#data['Data'] = datetime.datetime.strptime(data['Data'], format).date()

data['Data'] = data['Data'].dt.strftime('%Y-%m-%d')

#Inicjuj kolumny informujące o statuie transakcji
data['Kupuj'] = np.nan
data['Sprzedaj'] = np.nan
data['Trzymaj'] = np.nan

#WYznacz zmienne pozycyjne określające gdzie zaczynają się i kończa transakcje
trans_start = ind_par - 1
trans_end = data.shape[0]
ostatni_wiersz = trans_end - 1

#Wypełnij dane na temat statusu transakcji w pierwszym wierszu w którym można było dokonać transakcji
data['Sprzedaj'][trans_start] = 0
data['Trzymaj'][trans_start] = 0
if data['Trzymaj'][trans_start] == 0 and data['Close'][trans_start] > data['Indicator'][trans_start]:
    data['Kupuj'][trans_start] = 1
else:
    data['Kupuj'][trans_start] = 0


#Wypełnij pozostałe wiersze danymi na temat statusu transakcji
for i in range(ind_par,trans_end):
    prev = i - 1

    if data['Kupuj'][prev] == 1:
        data['Trzymaj'][i] = 1
    elif data['Sprzedaj'][prev] == 1:
        data['Trzymaj'][i] = 0
    else:
        data['Trzymaj'][i] = data['Trzymaj'][prev]

    if data['Trzymaj'][i] == 0 and data['Close'][i] > data['Indicator'][i]:
        data['Kupuj'][i] = 1
    else:
        data['Kupuj'][i] = 0

    if data['Trzymaj'][i] == 1 and data['Close'][i] < data['Indicator'][i]:
        data['Sprzedaj'][i] = 1
    else:
        data['Sprzedaj'][i] = 0

#Jeśli na koniec trzymam mam otwartą pozycję to ją zamykam
if data['Trzymaj'][ostatni_wiersz] == 1:
    data['Sprzedaj'][ostatni_wiersz] == 1

#Zainicjuj kolumny wykorzystywane do BEC i Sharpe'a
data['Miesiąc'] = np.nan
data['Rok'] = np.nan

#Wyciągnij informacje na temat roku i miesiąca

for i in range(ind_par,trans_end):
    data['Rok'][i] = data['Data'][i][0:4]
    data['Miesiąc'][i] = data['Data'][i][5:7]

y = pd.DataFrame(data)
y.to_excel("./transakcje_input/MA_20.xlsx")


######MA 50#################

ind_par = 50 #paramert indykatora

data = pd.ExcelFile('./input/EM.xlsx').parse('Arkusz1')


#Oblicz wartości wskaźnika
data['Indicator']=data['Close'].rolling(ind_par).mean()

#######################Przelicz transakcje#######################################

#format = '%Y-%m-%d %H:%M:%S'
#data['Data'] = datetime.datetime.strptime(data['Data'], format).date()

data['Data'] = data['Data'].dt.strftime('%Y-%m-%d')

#Inicjuj kolumny informujące o statuie transakcji
data['Kupuj'] = np.nan
data['Sprzedaj'] = np.nan
data['Trzymaj'] = np.nan

#WYznacz zmienne pozycyjne określające gdzie zaczynają się i kończa transakcje
trans_start = ind_par - 1
trans_end = data.shape[0]
ostatni_wiersz = trans_end - 1

#Wypełnij dane na temat statusu transakcji w pierwszym wierszu w którym można było dokonać transakcji
data['Sprzedaj'][trans_start] = 0
data['Trzymaj'][trans_start] = 0
if data['Trzymaj'][trans_start] == 0 and data['Close'][trans_start] > data['Indicator'][trans_start]:
    data['Kupuj'][trans_start] = 1
else:
    data['Kupuj'][trans_start] = 0

#Wypełnij pozostałe wiersze danymi na temat statusu transakcji
for i in range(ind_par,trans_end):
    prev = i - 1

    if data['Kupuj'][prev] == 1:
        data['Trzymaj'][i] = 1
    elif data['Sprzedaj'][prev] == 1:
        data['Trzymaj'][i] = 0
    else:
        data['Trzymaj'][i] = data['Trzymaj'][prev]

    if data['Trzymaj'][i] == 0 and data['Close'][i] > data['Indicator'][i]:
        data['Kupuj'][i] = 1
    else:
        data['Kupuj'][i] = 0

    if data['Trzymaj'][i] == 1 and data['Close'][i] < data['Indicator'][i]:
        data['Sprzedaj'][i] = 1
    else:
        data['Sprzedaj'][i] = 0

#Jeśli na koniec trzymam mam otwartą pozycję to ją zamykam
if data['Trzymaj'][ostatni_wiersz] == 1:
    data['Sprzedaj'][ostatni_wiersz] == 1

#Zainicjuj kolumny wykorzystywane do BEC i Sharpe'a
data['Miesiąc'] = np.nan
data['Rok'] = np.nan

#Wyciągnij informacje na temat roku i miesiąca

for i in range(ind_par,trans_end):
    data['Rok'][i] = data['Data'][i][0:4]
    data['Miesiąc'][i] = data['Data'][i][5:7]

y = pd.DataFrame(data)
y.to_excel("./transakcje_input/MA_50.xlsx")



######MA 100#################

ind_par = 100 #paramert indykatora

data = pd.ExcelFile('./input/EM.xlsx').parse('Arkusz1')


#Oblicz wartości wskaźnika
data['Indicator']=data['Close'].rolling(ind_par).mean()

#######################Przelicz transakcje#######################################

#format = '%Y-%m-%d %H:%M:%S'
#data['Data'] = datetime.datetime.strptime(data['Data'], format).date()

data['Data'] = data['Data'].dt.strftime('%Y-%m-%d')

#Inicjuj kolumny informujące o statuie transakcji
data['Kupuj'] = np.nan
data['Sprzedaj'] = np.nan
data['Trzymaj'] = np.nan

#WYznacz zmienne pozycyjne określające gdzie zaczynają się i kończa transakcje
trans_start = ind_par - 1
trans_end = data.shape[0]
ostatni_wiersz = trans_end - 1

#Wypełnij dane na temat statusu transakcji w pierwszym wierszu w którym można było dokonać transakcji
data['Sprzedaj'][trans_start] = 0
data['Trzymaj'][trans_start] = 0
if data['Trzymaj'][trans_start] == 0 and data['Close'][trans_start] > data['Indicator'][trans_start]:
    data['Kupuj'][trans_start] = 1
else:
    data['Kupuj'][trans_start] = 0

#Wypełnij pozostałe wiersze danymi na temat statusu transakcji
for i in range(ind_par,trans_end):
    prev = i - 1

    if data['Kupuj'][prev] == 1:
        data['Trzymaj'][i] = 1
    elif data['Sprzedaj'][prev] == 1:
        data['Trzymaj'][i] = 0
    else:
        data['Trzymaj'][i] = data['Trzymaj'][prev]

    if data['Trzymaj'][i] == 0 and data['Close'][i] > data['Indicator'][i]:
        data['Kupuj'][i] = 1
    else:
        data['Kupuj'][i] = 0

    if data['Trzymaj'][i] == 1 and data['Close'][i] < data['Indicator'][i]:
        data['Sprzedaj'][i] = 1
    else:
        data['Sprzedaj'][i] = 0

#Jeśli na koniec trzymam mam otwartą pozycję to ją zamykam
if data['Trzymaj'][ostatni_wiersz] == 1:
    data['Sprzedaj'][ostatni_wiersz] == 1

#Zainicjuj kolumny wykorzystywane do BEC i Sharpe'a
data['Miesiąc'] = np.nan
data['Rok'] = np.nan

#Wyciągnij informacje na temat roku i miesiąca

for i in range(ind_par,trans_end):
    data['Rok'][i] = data['Data'][i][0:4]
    data['Miesiąc'][i] = data['Data'][i][5:7]

y = pd.DataFrame(data)
y.to_excel("./transakcje_input/MA_100.xlsx")


######MA 150#################

ind_par = 150 #paramert indykatora

data = pd.ExcelFile('./input/EM.xlsx').parse('Arkusz1')


#Oblicz wartości wskaźnika
data['Indicator']=data['Close'].rolling(ind_par).mean()

#######################Przelicz transakcje#######################################

#format = '%Y-%m-%d %H:%M:%S'
#data['Data'] = datetime.datetime.strptime(data['Data'], format).date()

data['Data'] = data['Data'].dt.strftime('%Y-%m-%d')

#Inicjuj kolumny informujące o statuie transakcji
data['Kupuj'] = np.nan
data['Sprzedaj'] = np.nan
data['Trzymaj'] = np.nan

#WYznacz zmienne pozycyjne określające gdzie zaczynają się i kończa transakcje
trans_start = ind_par - 1
trans_end = data.shape[0]
ostatni_wiersz = trans_end - 1

#Wypełnij dane na temat statusu transakcji w pierwszym wierszu w którym można było dokonać transakcji
data['Sprzedaj'][trans_start] = 0
data['Trzymaj'][trans_start] = 0
if data['Trzymaj'][trans_start] == 0 and data['Close'][trans_start] > data['Indicator'][trans_start]:
    data['Kupuj'][trans_start] = 1
else:
    data['Kupuj'][trans_start] = 0

#Wypełnij pozostałe wiersze danymi na temat statusu transakcji
for i in range(ind_par,trans_end):
    prev = i - 1

    if data['Kupuj'][prev] == 1:
        data['Trzymaj'][i] = 1
    elif data['Sprzedaj'][prev] == 1:
        data['Trzymaj'][i] = 0
    else:
        data['Trzymaj'][i] = data['Trzymaj'][prev]

    if data['Trzymaj'][i] == 0 and data['Close'][i] > data['Indicator'][i]:
        data['Kupuj'][i] = 1
    else:
        data['Kupuj'][i] = 0

    if data['Trzymaj'][i] == 1 and data['Close'][i] < data['Indicator'][i]:
        data['Sprzedaj'][i] = 1
    else:
        data['Sprzedaj'][i] = 0

#Jeśli na koniec trzymam mam otwartą pozycję to ją zamykam
if data['Trzymaj'][ostatni_wiersz] == 1:
    data['Sprzedaj'][ostatni_wiersz] == 1

#Zainicjuj kolumny wykorzystywane do BEC i Sharpe'a
data['Miesiąc'] = np.nan
data['Rok'] = np.nan

#Wyciągnij informacje na temat roku i miesiąca

for i in range(ind_par,trans_end):
    data['Rok'][i] = data['Data'][i][0:4]
    data['Miesiąc'][i] = data['Data'][i][5:7]

y = pd.DataFrame(data)
y.to_excel("./transakcje_input/MA_150.xlsx")

######MA 200#################

ind_par = 200 #paramert indykatora

data = pd.ExcelFile('./input/EM.xlsx').parse('Arkusz1')


#Oblicz wartości wskaźnika
data['Indicator']=data['Close'].rolling(ind_par).mean()

#######################Przelicz transakcje#######################################

#format = '%Y-%m-%d %H:%M:%S'
#data['Data'] = datetime.datetime.strptime(data['Data'], format).date()

data['Data'] = data['Data'].dt.strftime('%Y-%m-%d')

#Inicjuj kolumny informujące o statuie transakcji
data['Kupuj'] = np.nan
data['Sprzedaj'] = np.nan
data['Trzymaj'] = np.nan

#WYznacz zmienne pozycyjne określające gdzie zaczynają się i kończa transakcje
trans_start = ind_par - 1
trans_end = data.shape[0]
ostatni_wiersz = trans_end - 1

#Wypełnij dane na temat statusu transakcji w pierwszym wierszu w którym można było dokonać transakcji
data['Sprzedaj'][trans_start] = 0
data['Trzymaj'][trans_start] = 0
if data['Trzymaj'][trans_start] == 0 and data['Close'][trans_start] > data['Indicator'][trans_start]:
    data['Kupuj'][trans_start] = 1
else:
    data['Kupuj'][trans_start] = 0

#Wypełnij pozostałe wiersze danymi na temat statusu transakcji
for i in range(ind_par,trans_end):
    prev = i - 1

    if data['Kupuj'][prev] == 1:
        data['Trzymaj'][i] = 1
    elif data['Sprzedaj'][prev] == 1:
        data['Trzymaj'][i] = 0
    else:
        data['Trzymaj'][i] = data['Trzymaj'][prev]

    if data['Trzymaj'][i] == 0 and data['Close'][i] > data['Indicator'][i]:
        data['Kupuj'][i] = 1
    else:
        data['Kupuj'][i] = 0

    if data['Trzymaj'][i] == 1 and data['Close'][i] < data['Indicator'][i]:
        data['Sprzedaj'][i] = 1
    else:
        data['Sprzedaj'][i] = 0

#Jeśli na koniec trzymam mam otwartą pozycję to ją zamykam
if data['Trzymaj'][ostatni_wiersz] == 1:
    data['Sprzedaj'][ostatni_wiersz] == 1

#Zainicjuj kolumny wykorzystywane do BEC i Sharpe'a
data['Miesiąc'] = np.nan
data['Rok'] = np.nan

#Wyciągnij informacje na temat roku i miesiąca

for i in range(ind_par,trans_end):
    data['Rok'][i] = data['Data'][i][0:4]
    data['Miesiąc'][i] = data['Data'][i][5:7]

y = pd.DataFrame(data)
y.to_excel("./transakcje_input/MA_200.xlsx")





######MACD#################

ind_par = 33 #paramert indykatora

data = pd.ExcelFile('./input/EM.xlsx').parse('Arkusz1')

# Policz wskaźnik MACD
data.ta.macd(close='Close', fast=12, slow=26, signal=9, append=True)
#Usuń zbędne kolumny
del data["MACD_12_26_9"]
del data["MACDs_12_26_9"]
#Zmień nazwę kolumny z wskaźnikiem
data=data.rename(columns={"MACDh_12_26_9": "Indicator"})

#######################Przelicz transakcje#######################################

#format = '%Y-%m-%d %H:%M:%S'
#data['Data'] = datetime.datetime.strptime(data['Data'], format).date()

data['Data'] = data['Data'].dt.strftime('%Y-%m-%d')

#Inicjuj kolumny informujące o statuie transakcji
data['Kupuj'] = np.nan
data['Sprzedaj'] = np.nan
data['Trzymaj'] = np.nan

#Wypełnij dane na temat statusu transakcji w pierwszym wierszu w którym można było dokonać transakcji
data['Sprzedaj'][ind_par] = 0
data['Trzymaj'][ind_par] = 0
if data['Trzymaj'][ind_par] == 0 and data['Indicator'][ind_par] > 0:
    data['Kupuj'][ind_par] = 1
else:
    data['Kupuj'][ind_par] = 0

#WYznacz zmienne pozycyjne określające gdzie zaczynają się i kończa transakcje
trans_start = ind_par+1
trans_end = data.shape[0]
ostatni_wiersz = trans_end - 1

#Wypełnij pozostałe wiersze danymi na temat statusu transakcji
for i in range(trans_start,trans_end):
    prev = i - 1

    if data['Kupuj'][prev] == 1:
        data['Trzymaj'][i] = 1
    elif data['Sprzedaj'][prev] == 1:
        data['Trzymaj'][i] = 0
    else:
        data['Trzymaj'][i] = data['Trzymaj'][prev]

    if data['Trzymaj'][i] == 0 and data['Indicator'][i] > 0:
        data['Kupuj'][i] = 1
    else:
        data['Kupuj'][i] = 0

    if data['Trzymaj'][i] == 1 and data['Indicator'][i] < 0:
        data['Sprzedaj'][i] = 1
    else:
        data['Sprzedaj'][i] = 0

#Jeśli na koniec trzymam mam otwartą pozycję to ją zamykam
if data['Trzymaj'][ostatni_wiersz] == 1:
    data['Sprzedaj'][ostatni_wiersz] == 1


#Zainicjuj kolumny wykorzystywane do BEC i Sharpe'a
data['Miesiąc'] = np.nan
data['Rok'] = np.nan

#Wyciągnij informacje na temat roku i miesiąca

for i in range(ind_par,trans_end):
    data['Rok'][i] = data['Data'][i][0:4]
    data['Miesiąc'][i] = data['Data'][i][5:7]

y = pd.DataFrame(data)
y.to_excel("./transakcje_input/MACD.xlsx")


######RSI#################

ind_par = 14 #paramert indykatora

data = pd.ExcelFile('./input/EM.xlsx').parse('Arkusz1')

# Policz wskaźnik RSI
data.ta.rsi(close='Close', append=True)

#Zmień nazwę kolumny z wskaźnikiem
data=data.rename(columns={"RSI_14": "Indicator"})

#######################Przelicz transakcje#######################################

#format = '%Y-%m-%d %H:%M:%S'
#data['Data'] = datetime.datetime.strptime(data['Data'], format).date()

data['Data'] = data['Data'].dt.strftime('%Y-%m-%d')

#Inicjuj kolumny informujące o statuie transakcji
data['Kupuj'] = np.nan
data['Sprzedaj'] = np.nan
data['Trzymaj'] = np.nan

#Wypełnij dane na temat statusu transakcji w pierwszym wierszu w którym można było dokonać transakcji
data['Sprzedaj'][ind_par] = 0
data['Trzymaj'][ind_par] = 0
if data['Trzymaj'][ind_par] == 0 and data['Indicator'][ind_par] > 0:
    data['Kupuj'][ind_par] = 1
else:
    data['Kupuj'][ind_par] = 0

#WYznacz zmienne pozycyjne określające gdzie zaczynają się i kończa transakcje
trans_start = ind_par+1
trans_end = data.shape[0]
ostatni_wiersz = trans_end - 1

#Wypełnij pozostałe wiersze danymi na temat statusu transakcji
for i in range(trans_start,trans_end):
    prev = i - 1

    if data['Kupuj'][prev] == 1:
        data['Trzymaj'][i] = 1
    elif data['Sprzedaj'][prev] == 1:
        data['Trzymaj'][i] = 0
    else:
        data['Trzymaj'][i] = data['Trzymaj'][prev]

    if data['Trzymaj'][i] == 0 and data['Indicator'][i] > 40:
        data['Kupuj'][i] = 1
    else:
        data['Kupuj'][i] = 0

    if data['Trzymaj'][i] == 1 and data['Indicator'][i] < 40:
        data['Sprzedaj'][i] = 1
    else:
        data['Sprzedaj'][i] = 0

#Jeśli na koniec trzymam mam otwartą pozycję to ją zamykam
if data['Trzymaj'][ostatni_wiersz] == 1:
    data['Sprzedaj'][ostatni_wiersz] == 1

#Zainicjuj kolumny wykorzystywane do BEC i Sharpe'a
data['Miesiąc'] = np.nan
data['Rok'] = np.nan

#Wyciągnij informacje na temat roku i miesiąca

for i in range(ind_par,trans_end):
    data['Rok'][i] = data['Data'][i][0:4]
    data['Miesiąc'][i] = data['Data'][i][5:7]

y = pd.DataFrame(data)
y.to_excel("./transakcje_input/RSI.xlsx")


######ROC#################

ind_par = 50 #paramert indykatora

data = pd.ExcelFile('./input/EM.xlsx').parse('Arkusz1')

# Policz wskaźnik ROC
data.ta.roc(close='Close', length=50, append=True)

#Zmień nazwę kolumny z wskaźnikiem
data=data.rename(columns={"ROC_50": "Indicator"})

#######################Przelicz transakcje#######################################

#format = '%Y-%m-%d %H:%M:%S'
#data['Data'] = datetime.datetime.strptime(data['Data'], format).date()

data['Data'] = data['Data'].dt.strftime('%Y-%m-%d')

#Inicjuj kolumny informujące o statuie transakcji
data['Kupuj'] = np.nan
data['Sprzedaj'] = np.nan
data['Trzymaj'] = np.nan

#Wypełnij dane na temat statusu transakcji w pierwszym wierszu w którym można było dokonać transakcji
data['Sprzedaj'][ind_par] = 0
data['Trzymaj'][ind_par] = 0
if data['Trzymaj'][ind_par] == 0 and data['Indicator'][ind_par] > 0:
    data['Kupuj'][ind_par] = 1
else:
    data['Kupuj'][ind_par] = 0

#WYznacz zmienne pozycyjne określające gdzie zaczynają się i kończa transakcje
trans_start = ind_par+1
trans_end = data.shape[0]
ostatni_wiersz = trans_end - 1

#Wypełnij pozostałe wiersze danymi na temat statusu transakcji
for i in range(trans_start,trans_end):
    prev = i - 1

    if data['Kupuj'][prev] == 1:
        data['Trzymaj'][i] = 1
    elif data['Sprzedaj'][prev] == 1:
        data['Trzymaj'][i] = 0
    else:
        data['Trzymaj'][i] = data['Trzymaj'][prev]

    if data['Trzymaj'][i] == 0 and data['Indicator'][i] > 0:
        data['Kupuj'][i] = 1
    else:
        data['Kupuj'][i] = 0

    if data['Trzymaj'][i] == 1 and data['Indicator'][i] < 0:
        data['Sprzedaj'][i] = 1
    else:
        data['Sprzedaj'][i] = 0

#Jeśli na koniec trzymam mam otwartą pozycję to ją zamykam
if data['Trzymaj'][ostatni_wiersz] == 1:
    data['Sprzedaj'][ostatni_wiersz] == 1

#Zainicjuj kolumny wykorzystywane do BEC i Sharpe'a
data['Miesiąc'] = np.nan
data['Rok'] = np.nan

#Wyciągnij informacje na temat roku i miesiąca

for i in range(ind_par,trans_end):
    data['Rok'][i] = data['Data'][i][0:4]
    data['Miesiąc'][i] = data['Data'][i][5:7]

y = pd.DataFrame(data)
y.to_excel("./transakcje_input/ROC.xlsx")

