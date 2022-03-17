import pandas as pd
import numpy as np
import xlsxwriter
import datetime
from scipy.stats import norm

#Słownik z parametrami dla poszczególnych metod
parameter_value = {'MA_20': 20, 'MA_50': 50, "MA_100": 100, 'MA_150': 150, 'MA_200': 200, 'MACD': 33, 'ROC': 50, 'RSI': 14}

#Lista wartości współczynników korygujących, które będą testowane
for wspolczynnik_zakres in [0.2, 0.5]:
#for wspolczynnik_zakres in [0.2]:

#Pętla po metodach analizy technicznej
    for ta_method in ['MA_20', 'MA_50', 'MA_100', 'MA_150', 'MA_200', 'MACD', 'ROC', 'RSI']:
#    for ta_method in ['MA_20']:

        # paramert indykatora
        ind_par = parameter_value[ta_method]
        ind_par_1=ind_par-1
        ind_par_2=ind_par-2
        #nauka_koniec='2015-03-27'wyznacz końcowy zakres przedziału pierwszej nauki zarządzania majątkiem
        nauka_koniec='1994-12-30'
        wspolczynnik = wspolczynnik_zakres #parametr do obnizenia zaangazowania w dzwignie

        #Załaduj dane z transakcjami
        path_input = './transakcje_input/' + ta_method + '.xlsx'

        data = pd.ExcelFile(path_input).parse('Sheet1')

        #######################Przelicz transakcje#######################################

        #format = '%Y-%m-%d %H:%M:%S'
        #data['Data'] = datetime.datetime.strptime(data['Data'], format).date()

    #    data['Data'] = data['Data'].dt.strftime('%Y-%m-%d')



        #WYznacz zmienne pozycyjne określające gdzie zaczynają się i kończa transakcje
        trans_end = data.shape[0]
        ostatni_wiersz = trans_end - 1


        #######################Przelicz jednotstkowe wyniki transakcje#######################################

        #Inicjuj kolumny informujące o cenie kupna i zrealizowanym zysku lub stracie
        data['Cena kupna'] = np.nan
        data['Zrealizowany zysk'] = np.nan
        data['Cena kupna'][ind_par_2] = 0

        #Wypełnij informacje na temat ceny kupna
        for i in range(ind_par_1,trans_end):
            prev = i - 1

            if data['Kupuj'][i] == 1:
                data['Cena kupna'][i] = data['Close'][i]
            elif data['Sprzedaj'][prev] == 1:
                data['Cena kupna'][i] = 0
            else:
                data['Cena kupna'][i] = data['Cena kupna'][prev]

        #Wypełnij informacje na temat zrealizowanego zysku lub straty
        for i in range(ind_par_1,trans_end):
            if data['Sprzedaj'][i] == 1:
                data['Zrealizowany zysk'][i] = (data['Close'][i]-data['Cena kupna'][i])/data['Cena kupna'][i]
            else:
                data['Zrealizowany zysk'][i] = 0

        #######################NAUKA#######################################


        #Zaiinicjuj kolumne z optymalnymi dzielnikami, wartością oczekiwaną, mnożnikiem dźwigni finansowej i zyskiem Vince'a
        #przed zlogarytmowaniem
        data['max_strata'] = np.nan
        data['f_opt'] = np.nan
        data['Mnożnik'] = np.nan
        data['Skumulowane odsetki'] = np.nan
        data['Vince_zysk'] = np.nan

        #Wyznacz pozycję, do której system uczy się po raz pierwszy zarządzania majątkiem
        nauka_pos = 0
        for i in range(ind_par_1,trans_end):
            if data['Data'][i] == nauka_koniec:
                nauka_pos = i
                break

        #Wyznacz zbiór z wynikami tranksacji do nauki
        transakcje_nauka = []
        otatnia_transakcja = 0
        for i in range(ind_par_1,nauka_pos):
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
        for i in range(test_init,trans_end):
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
        data['Vince-BH'] = np.nan

        #Inicjuj kolumny do obliczenia BEC
        data['Vince-BH_norm'] = np.nan
        data['Mnożnik_BEC'] = np.nan

        #Zainicjuj zmienną badającą, czy doszło do bankructa
        zonk = 'nie'

        #Wypełnij wiersze informujące o wyniku metody kupuj i trzymaj, tadingu i Vince'a
        for i in range(test_init,trans_end):
            prev = i - 1

            #Badam wyniki tylko gdy można stosować Vince'a
            if data['f_opt'][i] > 0:
                if data['Sprzedaj'][i] == 1:
                    #Badam wyniki buy and hold
                    data['BH'][i] = np.log(1+data['Zrealizowany zysk'][i]) * 100
                    #Badam wyniki Vince'a i to czy doszło do bankructwa
                    if data['Vince_zysk'][i] <= -1:
                        data['Vince'][i] = -10000000
                        zonk = 'tak'
                    else:
                        data['Vince'][i] = np.log(1+data['Vince_zysk'][i]) * 100

                    #Wyznacz parametry do wyliczenia BEC
                    data['Vince-BH_norm'][i] = data['Vince_zysk'][i] - data['Zrealizowany zysk'][i]
                    data['Mnożnik_BEC'][i] = data['Mnożnik'][i]

                data['Vince-BH'][i] = data['Vince'][i] - data['BH'][i]

#######Sharp#######################

        #Policz wartości dla Sharpe'a
        data['BH_wealth'] = np.nan
        data['Vince_wealth'] = np.nan
        data['BH_w_prev'] = np.nan
        data['Vince_w_prev'] = np.nan
        data['Rf_m'] = np.nan
        data['BH_m_yield'] = np.nan
        data['Vince_m_yield'] = np.nan

        #Zaiinicjuj wartości początkowe majątku dla policzenia Sharpea
        data['BH_wealth'][test_init] = 100
        data['Vince_wealth'][test_init] = 100
        data['BH_w_prev'][test_init] = 100
        data['Vince_w_prev'][test_init] = 100

        #Wylicz wartości majątku w czasie
        trans_end_1 = trans_end - 1
        for i in range(test_start, trans_end):
            prev = i - 1

            #Policz majątek dla obu metod w wartościach bezwględnych
            if data['Sprzedaj'][i] == 1 and data['f_opt'][i] > 0 :
                data['BH_wealth'][i] = data['BH_wealth'][prev] * (1+data['Zrealizowany zysk'][i])
                data['Vince_wealth'][i] = data['Vince_wealth'][prev] * (1+data['Vince_zysk'][i])
            else:
                data['BH_wealth'][i] = data['BH_wealth'][prev]
                data['Vince_wealth'][i] = data['Vince_wealth'][prev]

        #Policz wartości dla Sharpe'a: miesięczna stopę wolną od ryzyka, stopy zwrotu na koniec miesięca i majątek
        #z końca zeszłego miesiąca

        for i in range(test_start, trans_end_1):
            prev = i - 1
            next = i + 1

            if data['Miesiąc'][i] != data['Miesiąc'][next]:
                data['Rf_m'][i] = np.log((1+data['Rf'][i]/100)**(1/12))
                data['BH_m_yield'][i] = np.log(data['BH_wealth'][i]/data['BH_w_prev'][prev])
                data['Vince_m_yield'][i] = np.log(data['Vince_wealth'][i]/data['Vince_w_prev'][prev])
                data['BH_w_prev'][i] = data['BH_wealth'][i]
                data['Vince_w_prev'][i] = data['Vince_wealth'][i]
            else:
                data['BH_w_prev'][i] = data['BH_w_prev'][prev]
                data['Vince_w_prev'][i] = data['Vince_w_prev'][prev]

        #Policz wartości dla ostatniego dnia
        for i in range(trans_end_1, trans_end):
            prev = i - 1

            data['Rf_m'][i] = np.log((1+data['Rf'][i]/100)**(1/12))
            data['BH_m_yield'][i] = np.log(data['BH_wealth'][i]/data['BH_w_prev'][prev])
            data['Vince_m_yield'][i] = np.log(data['Vince_wealth'][i]/data['Vince_w_prev'][prev])
            data['BH_w_prev'][i] = data['BH_wealth'][i]
            data['Vince_w_prev'][i] = data['Vince_wealth'][i]

        #Policz wartości ze wzoru na Sharpe'a dla buy and hold i Vince'a
        stopa_wolna = np.nanmean(data['Rf_m'])
        BH_srednia_Sharp = np.nanmean(data['BH_m_yield'])
        Vince_srednia_Sharp = np.nanmean(data['Vince_m_yield'])
        BH_SD_Sharp = np.sqrt(np.nanvar(data['BH_m_yield']))
        Vince_SD_Sharp = np.sqrt(np.nanvar(data['Vince_m_yield']))

        #Policz wskaźniki Sharpe'a
        sharp_BH = (BH_srednia_Sharp-stopa_wolna)/BH_SD_Sharp
        sharp_Vince = (Vince_srednia_Sharp-stopa_wolna)/Vince_SD_Sharp

        def podokresy(start, koniec):

            licznik_okes = 0
            ruina_okres = 'nie'
            Vince_BH_okres = 0
            Mnożnik_okres = 0

            for i in range(test_init, trans_end):
                if data['Rok'][i] > start and data['Rok'][i] < koniec:
                    if ~np.isnan(data['Vince_zysk'][i]):
                        if data['Vince'][i] == -10000000:
                            ruina_okres = 'tak'
                        else:
                            licznik_okes += 1
                            Vince_BH_okres += data['Vince-BH_norm'][i]
                            Mnożnik_okres += data['Mnożnik_BEC'][i]




            if licznik_okes > 0:
                BEC_okres = (Vince_BH_okres/licznik_okes)/(Mnożnik_okres/licznik_okes)
            else:
                BEC_okres = 'brak'

            print(Vince_BH_okres)
            print(licznik_okes)
            print(Mnożnik_okres)

            return BEC_okres, ruina_okres

        #BEC_1990, ruina_1990 = podokresy(1985, 1996)
        #BEC_1995, ruina_1995 = podokresy(1995, 2001)
        #BEC_2000, ruina_2000 = podokresy(2000, 2006)
        #BEC_2005, ruina_2005 = podokresy(2005, 2011)
        #BEC_2010, ruina_2010 = podokresy(2010, 2016)
        #BEC_2015, ruina_2015 = podokresy(2015, 2050)




 #########Pozostałe statystyki podsumujące###################

        #Policz jaki procent obserwacji kwalifikuje sie do stosowania dzwigni
        procent_EX_dodatni = np.count_nonzero(~np.isnan(data['f_opt']))/(trans_end-test_init) * 100
        #Policz średni zysk Vince'a
        sredni_Vince = np.nanmean(data['Vince'])

        #Wyznacz zmienne porownujace Vince'a z buy and hold
        Srednia_Vince_BH = np.nanmean(data['Vince-BH'])
        Wariancja_Vince_BH = np.nanvar(data['Vince-BH'])
        Obserwacje = np.count_nonzero(~np.isnan(data['Vince_zysk']))
        t_student_Vince_BH = Srednia_Vince_BH/np.sqrt(Wariancja_Vince_BH/Obserwacje)
        pvalue_Vince_BH = 1-norm.cdf(t_student_Vince_BH)

        #Policz dodatkowe statystki
        BEC = np.nanmean(data['Vince-BH_norm'])/np.nanmean(data['Mnożnik_BEC']) * 100

        #######################Output#######################################

        #Stworz ścieżkę pliku outputowego
        path = './output/' + str(wspolczynnik) + '_' + ta_method + '.xlsx'

        # Stwórz plik wynikowy z arkuszem
        workbook = xlsxwriter.Workbook(path)
        worksheet = workbook.add_worksheet()

        #Poszerz kolumny
        worksheet.set_column('A:A', 15)
        worksheet.set_column('B:B', 17)
        worksheet.set_column('C:C', 23)
        worksheet.set_column('D:D', 19)
        worksheet.set_column('E:E', 19)
        worksheet.set_column('F:F', 17)
        worksheet.set_column('G:G', 13)
        worksheet.set_column('H:H', 23)
        worksheet.set_column('I:I', 24)

        #Wypełnij nagłówki wiersz opisowych
        worksheet.write('A1', 'Nazwa')
        worksheet.write('B1', 'Parametr')
        worksheet.write('C1', 'Współczynnik korygujący')
        worksheet.write('D1', 'Zonk')
        worksheet.write('E1', 'Początek obserwacji')
        worksheet.write('F1', 'Koniec obserwacji')
        worksheet.write('G1', 'Koniec nauki')
        worksheet.write('H1', 'Liczba transakcji uczących')
        worksheet.write('I1', 'Procent dni z EX dodatnim')

        #Wepełnij wartości opisowe
        worksheet.write('A2', ta_method)
        worksheet.write('B2', ind_par)
        worksheet.write('C2', wspolczynnik)
        worksheet.write('D2', zonk)
        worksheet.write('E2', data['Data'][0])
        worksheet.write('F2', data['Data'][ostatni_wiersz])
        worksheet.write('G2', nauka_koniec)
        worksheet.write('H2', liczba_transakcji)
        worksheet.write('I2', procent_EX_dodatni)

        #Wypełnij statystyki porównujące Vince'a z kupuj i trzymaj
        worksheet.write('A4', 'Vince-BH')
        worksheet.write('A5', 'Vince - średnia')
        worksheet.write('B5', 'Średnia')
        worksheet.write('C5', 'Wariancja')
        worksheet.write('D5', 't-student')
        worksheet.write('E5', 'p-value')
        worksheet.write('A6', sredni_Vince)
        worksheet.write('B6', Srednia_Vince_BH)
        worksheet.write('C6', Wariancja_Vince_BH)
        worksheet.write('D6', t_student_Vince_BH)
        worksheet.write('E6', pvalue_Vince_BH)

        #Wypełnij dodatkowe statystyki
        worksheet.write('A8', 'BEC')
        worksheet.write('A9', BEC)
        worksheet.write('B8', 'Sharp_BH')
        worksheet.write('B9', sharp_BH)
        worksheet.write('C8', 'Sharp_Vince')
        worksheet.write('C9', sharp_Vince)

        #Wypełnij informacje o podokresach
        #worksheet.write('A11', 'BEC_1990')
        #worksheet.write('A12', BEC_1990)
        #worksheet.write('B11', 'ruina_1990')
        #worksheet.write('B12', ruina_1990)
        #worksheet.write('C11', 'BEC_1995')
        #worksheet.write('C12', BEC_1995)
        #worksheet.write('D11', 'ruina_1995')
        #worksheet.write('D12', ruina_1995)
        #worksheet.write('E11', 'BEC_2000')
        #worksheet.write('E12', BEC_2000)
        #worksheet.write('F11', 'ruina_2000')
        #worksheet.write('F12', ruina_2000)
        #worksheet.write('G11', 'BEC_2005')
        #worksheet.write('G12', BEC_2005)
       #worksheet.write('H11', 'ruina_2005')
        #worksheet.write('H12', ruina_2005)
        #worksheet.write('I11', 'BEC_2010')
        #worksheet.write('I12', BEC_2010)
        #worksheet.write('J11', 'ruina_2010')
        #worksheet.write('J12', ruina_2010)
        #worksheet.write('A14', 'BEC_2015')
        #worksheet.write('A15', BEC_2015)
        #worksheet.write('B14', 'ruina_2015')
        #worksheet.write('B15', ruina_2015)

        workbook.close()

        #Usun zbędne kolumny
        #del data['Vince-BH_norm']
        #del data['Mnożnik_BEC']
        del data['Cena kupna']
        del data['Skumulowane odsetki']
        del data['Unnamed: 0']
        del data['Indicator']
        del data['max_strata']
        #del data['f_opt']
        del data['Mnożnik']
        del data['BH']
        del data['Vince']
        #del data['Vince-BH']

        y = pd.DataFrame(data)
        y.to_excel("output.xlsx")


