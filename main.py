import pandas as pd
import numpy as np
import xlsxwriter
import datetime
from scipy.stats import norm


ind_par = 100 #paramert indykatora
ind_par_1=ind_par-1
#nauka_koniec='2015-03-27'wyznacz końcowy zakres przedziału pierwszej nauki zarządzania majątkiem
nauka_koniec='1990-12-31'
wspolczynnik = 0.1 #parametr do obnizenia zaangazowania w dzwignie

data = pd.ExcelFile('./input/EM.xlsx').parse('Arkusz1')


#Oblicz wartości wskaźnika
data['Indicator']=data['Close'].rolling(ind_par).mean()

#######################Przelicz transakcje#######################################

#format = '%Y-%m-%d %H:%M:%S'
#data['Data'] = datetime.datetime.strptime(data['Data'], format).date()

data['Data'] = data['Data'].dt.strftime('%Y-%m-%d')
print(data['Data'][0])

#Inicjuj kolumny informujące o statuie transakcji
data['Kupuj'] = np.nan
data['Sprzedaj'] = np.nan
data['Trzymaj'] = np.nan

#Wypełnij dane na temat statusu transakcji w pierwszym wierszu w którym można było dokonać transakcji
data['Sprzedaj'][ind_par] = 0
data['Trzymaj'][ind_par] = 0
if data['Trzymaj'][ind_par] == 0 and data['Close'][ind_par] > data['Indicator'][ind_par]:
    data['Kupuj'][ind_par] = 1
else:
    data['Kupuj'][ind_par] = 0

#WYznacz zmienne pozycyjne określające gdzie zaczynają się i kończa transakcje
trans_start = ind_par + 1
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

#######################Przelicz jednotstkowe wyniki transakcje#######################################

#Inicjuj kolumny informujące o cenie kupna i zrealizowanym zysku lub stracie
data['Cena kupna'] = np.nan
data['Zrealizowany zysk'] = np.nan
data['Cena kupna'][ind_par_1] = 0

#Wypełnij informacje na temat ceny kupna
for i in range(ind_par,trans_end):
    prev = i - 1

    if data['Kupuj'][i] == 1:
        data['Cena kupna'][i] = data['Close'][i]
    elif data['Sprzedaj'][prev] == 1:
        data['Cena kupna'][i] = 0
    else:
        data['Cena kupna'][i] = data['Cena kupna'][prev]

#Wypełnij informacje na temat zrealizowanego zysku lub straty
for i in range(ind_par,trans_end):
    if data['Sprzedaj'][i] == 1:
        data['Zrealizowany zysk'][i] = (data['Close'][i]-data['Cena kupna'][i])/data['Cena kupna'][i]
    else:
        data['Zrealizowany zysk'][i] = 0

#######################NAUKA#######################################


#Zaiinicjuj kolumne z optymalnymi dzielnikami, wartością oczekiwaną i mnożnikiem dźwigni finansowej
data['max_strata'] = np.nan
data['f_opt'] = np.nan
data['Mnożnik'] = np.nan
data['Skumulowane odsetki'] = np.nan
data['Vince_zysk'] = np.nan

#Wyznacz pozycję, do której system uczy się po raz pierwszy zarządzania majątkiem
nauka_pos = 0
for i in range(ind_par,trans_end):
    if data['Data'][i] == nauka_koniec:
        nauka_pos = i
        break

#Wyznacz zbiór z wynikami tranksacji do nauki
transakcje_nauka = []
otatnia_transakcja = 0
for i in range(ind_par,nauka_pos):
    if data['Sprzedaj'][i] == 1:
        transakcje_nauka.append(data['Zrealizowany zysk'][i])
        otatnia_transakcja = i

#Policz maksymalna stratę, liczbę transakcji w próbie uczącej, prawdopodobieństwa danego wyniku, wartość oczekiwaną
max_strata = min(transakcje_nauka) * (-1)
liczba_transakcji = len(transakcje_nauka)
prawd = 1 / liczba_transakcji
EX = np.mean(transakcje_nauka)

#Funkcja liczaca optymalny dzielnik
def nauka(liczba_transakcji_f, prawd_f, transakcje_nauka_f, max_strata_f):
    #Wwyznacz wartosci funkcji G wzgledem wartosci dzielnika oraz transakcji
    funkcja_wartosci = []
    for i in range(0, 100):
        f = i / 100
        Gf_values = []
        for j in range(0, liczba_transakcji_f):
            value = prawd * np.log(1 + f * transakcje_nauka_f[j] / max_strata_f)
            Gf_values.append(value)
        funkcja_wartosci.append(Gf_values)

    f_temp = -1
    f_opt_f = 0
    sum_temp = 0
    #Zbadaj ktory f jest optymalny na podstwie wartosci funkcji
    for i in range(0, 100):
        sum_temp = sum(funkcja_wartosci[i])
        if sum_temp > f_temp:
            f_temp = sum_temp
        else:
            f_opt_f = (i - 1) / 100
            break

    return wspolczynnik * f_opt_f

#Wyznacz pierwszy optymalny f
f_opt = nauka(liczba_transakcji, prawd, transakcje_nauka, max_strata)

#Wyznacz poczatek proby testowej i pierwsza komorke ktore jest obliczana nieco inaczej niz reszta
test_init = otatnia_transakcja + 1
test_start = otatnia_transakcja + 2

#######################Zastosowanie DŹWIGNI#######################################

#Jesli wartosc oczekiwana jest dodatnia dla pierwszej komorki proby testowej to przypisz optymalne f
if EX > 0:
    data['f_opt'][test_init] = f_opt

#Zainicjuj licznik, ktory bada ktora transakcja jest piewsza do wylotu
licznik_nowych_transakcji = 0
for i in range(test_start,trans_end):
    prev = i - 1
    #Jesli w poprzednim dniu sprzedalem to aktualizuje licznik, max strate, EX, optymalne f
    if data['Sprzedaj'][prev] == 1:
        licznik_nowych_transakcji = licznik_nowych_transakcji % liczba_transakcji
        transakcje_nauka[licznik_nowych_transakcji] = data['Zrealizowany zysk'][prev]
        licznik_nowych_transakcji = licznik_nowych_transakcji + 1

        max_strata = min(transakcje_nauka) * (-1)
        EX = np.mean(transakcje_nauka)
        f_opt = nauka(liczba_transakcji, prawd, transakcje_nauka, max_strata)

#Jesli wartosc oczekiwana jest dodatnia to przypisuje optymalne f
    if  EX > 0:
        data['f_opt'][i] = f_opt
        data['max_strata'][i]  = max_strata

        #Policz mnożnik dźwigni finansowej
        if data['Kupuj'][prev] == 1:
            data['Mnożnik'][i] = data['f_opt'][i] / max_strata
        elif data['Trzymaj'][i] == 1:
            data['Mnożnik'][i] = data['Mnożnik'][prev]

#Policz skumulowane odsetki i zysk z metody Vince'a
for i in range(test_start,trans_end):
    prev = i - 1
    if  data['Kupuj'][i] == 1:
        data['Skumulowane odsetki'][i] = data['Rf'][i]/260
    elif data['Trzymaj'][i] == 1:
        data['Skumulowane odsetki'][i] = data['Skumulowane odsetki'][prev] + data['Rf'][i]/260

    if data['Sprzedaj'][i] == 1:
        data['Vince_zysk'][i] = data['Mnożnik'][i] * data['Zrealizowany zysk'][i] - (data['Mnożnik'][i]-1) * data['Skumulowane odsetki'][i]/100


        #######################Porównanie metod#######################################

#Inicjuj kolumny informujące o wynikach poszczególnych metod
data['BH'] = np.nan
data['Vince'] = np.nan
#data['Vince-BH'] = np.nan

#Wypełnij wiersze informujące o wyniku metody kupuj i trzymaj, tadingu i Vince'a
for i in range(test_init,trans_end):
    prev = i - 1

    #Badam wyniki tylko gdy można stosować Vince'a
    if data['f_opt'][i] > 0:
        if data['Sprzedaj'][i] == 1:
            #Badam wyniki buy and hold
            data['BH'][i] = np.log(1+data['Zrealizowany zysk'][i]) * 100

            if data['Vince_zysk'][i] <= -1:
                data['Vince'][i] = -10000000
            else:
                data['Vince'][i] = np.log(1+data['Vince_zysk'][i]) * 100


y = pd.DataFrame(data)
y.to_excel("output.xlsx")


"""
        #Badam wyniki tradingu i Vonce'a
        if data['Trzymaj'][i] == 1:
            data['Trading'][i] = data['BH'][i]
            data['Vince'][i] = np.log((data['Close'][i]+data['Mnożnik'][i]*(data['Close'][i]-data['Close'][prev])-\
            (data['Mnożnik'][i]-1)*data['Rf'][i]/260*data['Close'][i])/data['Close'][i]) * 100
        else:
            data['Trading'][i] = np.log(1+data['Rf'][i]/260)*100
            data['Vince'][i] = data['Trading'][i]

        
        data['TR-BH'][i] = data['Trading'][i] - data['BH'][i]
        data['Vince-BH'][i] = data['Vince'][i] - data['BH'][i]
        data['Vince-TR'][i] = data['Vince'][i] - data['Trading'][i]





#Policz jaki procent obserwacji kwalifikuje sie do stosowania dzwigni
procent_EX_dodtni = np.count_nonzero(~np.isnan(data['f_opt']))/(trans_end-test_init) * 100

#Wyznacz zmienne porownujace trading z buy and hold
Srednia_TR_BH = np.nanmean(data['TR-BH'])
Wariancja_TR_BH = np.nanvar(data['TR-BH'])
Obserwacje = np.count_nonzero(~np.isnan(data['f_opt']))
t_student_TR_BH = Srednia_TR_BH/np.sqrt(Wariancja_TR_BH/Obserwacje)
pvalue_TR_BH = 1-norm.cdf(t_student_TR_BH)

#Wyznacz zmienne porownujace Vince'a z buy and hold
Srednia_Vince_BH = np.nanmean(data['Vince-BH'])
Wariancja_Vince_BH = np.nanvar(data['Vince-BH'])
t_student_Vince_BH = Srednia_Vince_BH/np.sqrt(Wariancja_Vince_BH/Obserwacje)
pvalue_Vince_BH = 1-norm.cdf(t_student_Vince_BH)

#Wyznacz zmienne porownujace Vince'a z tradingiem
Srednia_Vince_TR = np.nanmean(data['Vince-TR'])
Wariancja_Vince_TR = np.nanvar(data['Vince-TR'])
t_student_Vince_TR = Srednia_Vince_TR/np.sqrt(Wariancja_Vince_TR/Obserwacje)
pvalue_Vince_TR = 1-norm.cdf(t_student_Vince_TR)

#######################Output#######################################

#Stworz ścieżkę pliku outputowego
directory = './output/'
nazwa = 'MA'
path = directory + nazwa + '_' + str(ind_par) + '.xlsx'

# Stwórz plik wynikowy z arkuszem
workbook = xlsxwriter.Workbook(path)
worksheet = workbook.add_worksheet()

#Poszerz kolumny
worksheet.set_column('A:A', 15)
worksheet.set_column('B:B', 17)
worksheet.set_column('C:C', 23)
worksheet.set_column('D:D', 19)
worksheet.set_column('E:E', 17)
worksheet.set_column('F:F', 13)
worksheet.set_column('G:G', 23)
worksheet.set_column('H:H', 24)

#Wypełnij nagłówki wiersz opisowych
worksheet.write('A1', 'Nazwa')
worksheet.write('B1', 'Parametr')
worksheet.write('C1', 'Współczynnik korygujący')
worksheet.write('D1', 'Początek obserwacji')
worksheet.write('E1', 'Koniec obserwacji')
worksheet.write('F1', 'Koniec nauki')
worksheet.write('G1', 'Liczba transakcji uczących')
worksheet.write('H1', 'Procent dni z EX dodatnim')

#Wepełnij wartości opisowe
worksheet.write('A2', 'MA')
worksheet.write('B2', ind_par)
worksheet.write('C2', wspolczynnik)
worksheet.write('D2', data['Data'][0])
worksheet.write('E2', data['Data'][ostatni_wiersz])
worksheet.write('F2', nauka_koniec)
worksheet.write('G2', liczba_transakcji)
worksheet.write('H2', procent_EX_dodtni)

#Wypełnij statystyki porównujące trading z kupuj i trzymaj
worksheet.write('A4', 'TR-BH')
worksheet.write('A5', 'Średnia')
worksheet.write('B5', 'Wariancja')
worksheet.write('C5', 't-student')
worksheet.write('D5', 'p-value')
worksheet.write('A6', Srednia_TR_BH)
worksheet.write('B6', Wariancja_TR_BH)
worksheet.write('C6', t_student_TR_BH)
worksheet.write('D6', pvalue_TR_BH)

#Wypełnij statystyki porównujące Vince'a z kupuj i trzymaj
worksheet.write('A8', 'Vince-BH')
worksheet.write('A9', 'Średnia')
worksheet.write('B9', 'Wariancja')
worksheet.write('C9', 't-student')
worksheet.write('D9', 'p-value')
worksheet.write('A10', Srednia_Vince_BH)
worksheet.write('B10', Wariancja_Vince_BH)
worksheet.write('C10', t_student_Vince_BH)
worksheet.write('D10', pvalue_Vince_BH)

#Wypełnij statystyki porównujące Vince'a z tradingiem
worksheet.write('A12', 'Vince-TR')
worksheet.write('A13', 'Średnia')
worksheet.write('B13', 'Wariancja')
worksheet.write('C13', 't-student')
worksheet.write('D13', 'p-value')
worksheet.write('A14', Srednia_Vince_TR)
worksheet.write('B14', Wariancja_Vince_TR)
worksheet.write('C14', t_student_Vince_TR)
worksheet.write('D14', pvalue_Vince_TR)


workbook.close()
"""


