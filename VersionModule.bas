Attribute VB_Name = "VersionModule"
' FORREST SOFTWARE
' Copyright (c) 2016 Mateusz Forrest Milewski
'
' Permission is hereby granted, free of charge,
' to any person obtaining a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation the rights to
' use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
' and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

' ======================================================
' wersja 2.92 proste dodanie wylaczenia alertu podczas usuwania arkuszy side;owych
' dodanie nowych pivotow na stale + ribbon button na uruchomienie jak i "go to"
' ======================================================



' ======================================================
' wersja 2.91 - scaffold dla all 2016-04-27
''
' pod backup
' ======================================================

' ======================================================
' wersja 2.90 - scaffold dla all 2016-04-27
'
' rozszerzenie rep all
'
''
' ' .Cells(x, 1) = "PLT"
' ' .Cells(x, 2) = "PROJ"
' ' .Cells(x, 3) = "BG"
' ' .Cells(x, 4) = "MY"
' ' .Cells(x, 5) = "FAZA"
' ' .Cells(x, 6) = "MRD"
' ' .Cells(x, 7) = "MRDd"
' ' .Cells(x, 8) = "COORD"
' ' .Cells(x, 9) = "RESP"
' ' .Cells(x, 10) = "FUP"
' ' .Cells(x, 11) = "PN"
' ' .Cells(x, 12) = "DEL CONF"
' ' .Cells(x, 13) = "Comments"
' ' .Cells(x, 14) = "DUNS"
' ' .Cells(x, 15) = "SUPP NM"
' ' .Cells(x, 16) = "Total QTY"
' ' .Cells(x, 17) = "Ordered Date"
' ' .Cells(x, 18) = "PPAP Status"
' ' .Cells(x, 19) = "Ordered Qty"
' ' .Cells(x, 20) = "Conf Qty"
' ' .Cells(x, 21) = "Fst Pickup Date"
' ======================================================

' ======================================================
' wersja 2.89 - scaffold dla all 2016-04-21
'
''
' 1. rozszerzony modul usuwania
' 2. jak madrzej sobie radzic z danymi wizarda?
' 3. narazie na dzien 21 04 2016 jeszcze nie udostepniam
' tej wersji z powodu tylko i wylacznie nowej grupy del
' musi byc cos wiecej...
' zmiany lagodzace zle zaciagane dane.
' ======================================================

' ======================================================
' wersja 2.88 - scaffold dla all 2016-04-05
'
''
' udostepnienie - ostatnie szlify
' w tej wersji przesuwanie
' z repow do repow fupow i side'ow
' dziala tylko i wylacznie per arkusz side
' takie zalozenie nieco burzy harmonie dzialania
' ale zakladam ze ta funkcjonalnosc nie bedzie juz tak
' popularna jak raport all dla pivotow
' ======================================================




' ======================================================
' wersja 2.87 - scaffold dla all 2016-04-04
'
''
' pierwsze proby z githubem
' dodam teraz tekst by sprawdzic jak branche dzialaja
' ======================================================

' ======================================================
' wersja 2.86 - scaffold dla all 2016-04-01
'
''
' kopia zapasowe
' plus wstepne sledztwo co moze byc nie tak z static run
' poniewaz odkad wsadzilem logike run all - tam cos sie posypalo - to verify!
' ======================================================



' ======================================================
' wersja 2.85 - scaffold dla all 2016-03-31
'
''
' filtry  - nowe guziki

' dodawanie bez usuwania starych
' usuwanie z juz wybranego zbioru
' zostawianie z juz wybranego zbioru

' ======================================================

' ======================================================
' wersja 2.84 - scaffold dla all 2016-03-30
'
''
' del conf pivot i duzo slicerow
' oraz timeline - rewelacja!

' ======================================================

' ======================================================
' wersja 2.83 - scaffold dla all 2016-03-30
'
''
' jesli chodzi o ta wersje to wstepnie udalo sie podmienic linkowanie z nazw arkuszy na pesel (unique id)
' jeszcze nie bylo kompleksowych testow
' ale wyglada na to ze bedzie dzialac
' dzis to jest 3-30 rano kolo godziny 6:30 pod prysznicem wpadlem na pomysl aby raport all byl raportem 3 stronicowym
' poza tym dodatkowo bede chcial wykorzystac obiekt typu dictionary aby pobrac wszystkie dane na raz
' czyli obiekt zostanie utworzony przed iteracja po plikach nie bedzie czyszczony podczas i dopiero podsumowanie strzele na koniec
' opcja wstawiania za kazdym razem bedzie opierac sie na nowym arkuszu

' ======================================================

' ======================================================
' wersja 2.82 - scaffold dla all 2016-03-29
' narazie za 16 marca jest to kopia wersji 2.78
' dodalem ino nowa klase RecordHandlerAllApproach
' ktora ma przechowywac podstawowe info dla kazdego recordu (linii danych z wizarda)
' bez jakiej kolwiek ingerencji w przeliczanie
'
'
'
' rowniez prawdopodobnie bedzie to osobny arkusz jak i osobna logika otwierania plikow oraz zaciagania danych
' bez przeliczania co jest okiem a co jest nokiem - dodatkowo skupie sie tylko i wylacznie na del confie
' tak aby mozna bylo szybko i klarownie otrzymac raport
' comment 2.81
'
'
'
'
'
''
' 2.82 jeszcze dodatkowo
' dostosowanie numeru pesel z wizardow aby byly pelnoprawnym linkiem miedzy repami a arkuszami side

'
' jednak zeby za duzego balaganu nie bylo pozostawiam bez jeku
'  With wh
    ' .wyczysc_cfg_sheet_i_jej_tmp_list_na_name_i_phase

' troche trzeba tam liczenia ale jest trudno
' - nie chce za bardzo teraz inwigilowac w ta logike
' - moze kiedys to usune bezbolesnie

' ======================================================


' ======================================================
' wersja 2.81 - scaffold dla all 2016-0-16
' narazie za 16 marca jest to kopia wersji 2.78
' dodalem ino nowa klase RecordHandlerAllApproach
' ktora ma przechowywac podstawowe info dla kazdego recordu (linii danych z wizarda)
' bez jakiej kolwiek ingerencji w przeliczanie
'
'
'
' rowniez prawdopodobnie bedzie to osobny arkusz jak i osobna logika otwierania plikow oraz zaciagania danych
' bez przeliczania co jest okiem a co jest nokiem - dodatkowo skupie sie tylko i wylacznie na del confie
' tak aby mozna bylo szybko i klarownie otrzymac raport
' ======================================================



' ======================================================
' wersja 2.78 dzien po dniu liczby pi
' wlasciwie to samo co 77 jednak to jest taki zapis awaryjny
' szczegolnie gdy usunalem jedna klase dosyc wazna dodalem
' interfejs oraz dodatkowy guzik rozlewania w prawo
' wiec... tak na wypadek
' ======================================================
' wersja 2.77 2016-03-08
'
' to be implemented...
' 1. rozlanie info po prawej z perpektywy del conf ino
' narazie nic nie ma ale bedzie technicznie problem
' z rozpoznawaniem ile danych zostalo juz rozlanych
' bedzie trzeba aby kazde rozlanie dzialalo tak ssamo
' co powoduje ze musze i tak jeszcze raz zweryfikowac logike rozlewania dla fupow pojedynczo
'
'
' 2. dalej jest problem z linkowaniem arkuszy sideowych przez bardzo duzej spojnosci nazewnictwa projektow
' rozwiazaniem moze sie okazac nadanie kazdemu projektowi unikalnego identyfikatora niezaleznego od jego parametrow
' (prawdopodobnie id ten bedzie zapisany gdzie w komentarzu, w pierwszych kolumnach dla kazdej iteracji/projektu
'
' 2.1 przy okazji tego nowego rozwiazanie nie bede sie rozczulal nad wczesniejszym rozwiazaniem i jego srogim szafowaniem
' zostawie to wszystko tylko odlacze oden logika wizania miedzy arkuszami - trudno - bedzie duzo zbednego kodu zmieniajacego nazwy arkuszy
' ======================================================


' wersja 2.76 2016-02-23
' ======================================================
' fix na malym bugu zwiazanym z del confem
' lekka zmiana logiki na subie
' special treatment for delivery confirmation
' jest bug jesli chodzi o powyzsza zmiane
' pattern porownanie ma zbyt mala tolerancje
' i traktuje POT ITDC oraz ITDC jako to samo :)


' wersja 2.75 2016-02-23
' ======================================================
' rozszerzenie interfejsu dynamic del conf
' DynDelConfModule - duzo zmian
' na formularzu przyjacielskim zmiany z checkboxa na combobox

' wersja 2.74 2016-02-18
' ======================================================
' tylko od 3.9x kompatybilny


' wersja 2.73 2016-02-16
' ======================================================
' dopasowanie nowych wersji dla wizard w 3 - dla testow
' adjust na przesun dane z side do rep fup


' wersja 2.72 2016-02-16
' ======================================================
' dopasowanie nowych wersji dla wizard w 3 - dla testow
' adjust na przesun dane z side do rep
' adjust na przesun dane z rep do rep fup

' wersja 2.71 2016-02-15
' ======================================================
' dopasowanie nowych wersji dla wizard w 3,9x
' dodanie blank fupa co mozna bylo sprawdzic czy podlicza rowno rozlewanie


' wersja 2.65 2016-02-11
' ======================================================
' adjust w stosunku do wersji 2.64
' powazna zmiana obiekt klasy Dictionary zostal zmieniony na kolekcje ktorzy przechowuje elementy typu
' MyDictionary
'
' brakuje ucase na nazwie NM w klasie MyDictionary - to fix


' wersja 2.64 2016-02-10
' ======================================================
' udostepniony ze niby dziala jednak okazuje sie
' Poni?sze punkty, które wymagaj? poprawy:
' - data CW06 nie zawsze si? podswietli?a
' - w arkuszach NOK np. u Kasi po wyporze responsibility jest do wyboru FMA, FMA/  (w TD jest FMA/KSK)…


' ======================================================
' wersja 2.63/4 2016-02-09 TESTY!
' plus zwiecha!


' ======================================================
' wersja 2.62 2016-02-09

' overwrite only changed content OOCC! od wersji 2.6 std
' dodatkowe kolumny dla poj. FUPow


' ======================================================
' wersja 2.61 2016-02-08

' overwrite only changed content OOCC! od wersji 2.6 std
' nowa funkcjonalnosc z rep do rep fup


' ======================================================
' wersja 2.55 2016-02-05
' 2.54 jako tako dziala jednak przesun dane jest bardzo czasochlonne i wymaga zbyt duzej ilosci czasu
' gdy nasze podejscie sklania sie za kazdym razem do podmiany wszystkich danych
'
' zatem moja nowa koncepcja jest taka aby zmieniac tylko to co zostalo zmienione
' overwrite only changed content OOCC!
'
' ------------------
' niewinnie zaczne spogladac w strone rozszerzenia funkcjonalnosc na pojedynczego fupa
' ------------------
'
'
' ======================================================

' ======================================================
' wersja 2.54 2016-02-01
' przesun dane do rep - podzielic logike na pojedyncze arkusze za duzo czasu zajmuje
' nowe dane details
'
'
' ======================================================

' ======================================================
' wersja 2.53 2016-01-29
' przygotowanie fundamentu pod dynamiczna granice dla nowego jednak arkusza
' rozdzielic mrd pola - NOK
' w koncu zrobic transpozycje komentarza - done
' dodac pola po prawej stronie podzial na fupy - NOK
' plus reset do default - teraz w wersji od 2.53

' dodatkowe 3 paramsy zepsuly go trough selection - poprawione gwiazdkami w cond


' mrd1 date nie pokazuje sie w komentarzu - sprawdzic na ile to jest kupowe - chyba da rade teraz
' zle byl wpisany prefix
'
'
' ======================================================
' ======================================================
' wersja 2.43 2016-01-26
' przygotowanie fundamentu pod dynamiczna granice dla nowego jednak arkusza
' rozdzielic mrd pola
' w koncu zrobic transpozycje komentarza
' dodac pola po prawej stronie podzial na fupy
' plus reset do default

' dodatkowe 3 paramsy zepsuly go trough selection


' mrd1 date nie pokazuje sie w komentarzu - sprawdzic na ile to jest kupowe
'
'
' ======================================================

' ======================================================
' wersja 2.42 2016-01-25
' dodatkowe trzy pola gotowe (details)
' w global module param odpoweidzialny za testy lokalne
' w config nowa tabelka wraz z ustawianiem ok now dla dyn granicy dla del conf
' co z pus date? dalej nie pokazuje na niebiesko?
'
' przygotowanie fundamentu pod dynamiczna granice dla nowego jednak arkusza
'
' ======================================================


' ======================================================
' wersja 2.41 2016-01-22
' dodatkowe trzy pola gotowe
' w global module param odpoweidzialny za testy lokalne
'
' przygotowanie fundamentu pod dynamiczna granice dla nowego jednak arkusza
'
' ======================================================

' ======================================================
' wersja 2.4 od 2015-12-03
' kolumna pickup date nie swieci sie na niebiesko - to be
' dynamiczna granica del conf w nowym arkuszu - to be
' sideowe arkusze pelniejsze komentarze - to be done
' gotowe nowe pola details
' account comment w fazie
' mrd w polu w mrd - gotowe
'
'
' ======================================================

' ======================================================
' wersja 2.3 od 2015-12-03
' transpozycja komentarza
' kolumna pickup date nie swieci sie na niebiesko
'
'
' ======================================================



' ======================================================
' wersja 2.2 2015-11-26
' podwojne klikniecie na danych
' nie zaciaga niektorych danych
'
'
' ======================================================
' ======================================================
' wersja 2.1 2015-11-24
' nowa generacja - priorytet drugi do zrealizowania
' przesun dane do rep i rep fup - aktualizacje z wzgledu na zmiany w side arkuszach
' czerwono na NOKach
' niebieski na aktualny tydzien
' duns i pn wybor pojedynczy
'
' ======================================================

' ======================================================
' wersja 2.0
' now generacja
' dodana kolumna order date
' ======================================================


' wersja 1.6
' ON STOCK adjustted as true
' dalej rozwoj rep fup
' tym razem juz wszsytko bedzie widoczne an arksuzu
' i komentarze tez :D


' wersja 1.5
' kolekcja poraz drugi uzyta dla nowego arkusza rep fup
' narazie kolekcja gotowa plus zrzut tylko do fma resp kolumny

' wersja 1,4 2015-11-05 (Collector)
' Dim fh As FilterHandling - przed praca usuwam wszelkie chowane dane!

' wersja 1.3 2015-11-02 (Collector)
' pierwsze niewinne proby raportu dla fupow

' wersja 1.2 2015-10-31 (Collector)
' dopasowanie gotowych nazw

' wersja 0.8 2015-10-26 (pierwszy raz Collector)
' - poprawiona logika reassigningu nazwy arkuszu - nie bylo loopa - stad bledy wynikaly

' wersja 0.7 2015-10-26
' wersja przejsciowa dla reasigningu

' wersja 0.6 2015-10-26
' - ostateczne zerwanie z logika nie otwierania (zbyt wolne)
' - nowe arkusze - raport dla kazdego T/D

' wersja 0.5 2015-10-22
' - dodatkowa wersja lecimy tutaj - wraz
' - dodatkowy arkusz wszystkich danych w innym formacie ale tylko dla rozwinietej wersji lecimy tutaj


' wersja 0.4 2015-10-21
' - double click nareszczie dziala poprawnie
' - wciaz stara wersja lecimy tutaj

' wersja 0.3 od CW 43 2015 2015-10-20
' - rozszerzenie komentarza + double click dla generowania listy z komentarza
' - linkowanie do pliku na x z pierwszej kolumny
' - dodatkowe kolekcje zapamietujace co bylo nokiem (do rozwazenia)

' wersja 0.2 od CW 43 2015
' - proba rozszerzenia mozliwosci ExecuteExcel4Macro dla duzego zasiegu (proba nie powiodla sie,
'  zostawilem jednak obsoletowe sub'y, co bym widzial ze nie poszlo tak jak nalezy
' - weryfikacja poprawnosci zaciagania wszystkich Wizardow


' wersja 0.1 do CW 42 2015
' - bardzo proste sciaganie danych z arkuszy details
' - chyba jest problem ze sciagnieciem wszystkich plikow
' - ExecuteExcel4Macro wykorzystuje tylko dla pojedynczych komorek
