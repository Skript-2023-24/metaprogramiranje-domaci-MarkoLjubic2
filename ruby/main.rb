require './domaci.rb'

file_path1 = './Book1.xlsx'
file_path2 = './Book2.xlsx'
sheet_name1 = 'Sheet1'
sheet_name2 = 'Sheet2'

table1 = ExcelTable.new(file_path1, sheet_name1)
table2 = ExcelTable.new(file_path1, sheet_name2)
table3 = ExcelTable.new(file_path2, sheet_name1)

#Biblioteka može da vrati dvodimenzioni niz sa vrednostima tabele
p table1.array2D

#Moguće je pristupati redu preko t.row(1), i pristup njegovim elementima po sintaksi niza.
p table1.row(0)
p table1.row(1)[0]

#Mora biti implementiran Enumerable modul(each funkcija), gde se vraćaju sve ćelije unutar tabele, sa leva na desno.
table1.each { |row| p row }

#Biblioteka treba da vodi računa o merge-ovanim poljima(samo prvi gornji red uzima vrednost, ostali su nil)
p table1.array2D

#Biblioteka vraća celu kolonu kada se napravi upit t[“Prva Kolona”]
p table1['Ime']
#Biblioteka omogućava pristup vrednostima unutar kolone po sledećoj sintaksi t[“Prva Kolona”][1] za pristup drugom elementu te kolone
p table1['Ime'][1]

#c.	Biblioteka omogućava podešavanje vrednosti unutar ćelije po sledećoj sintaksi t[“Prva Kolona”][1]= 2556
table1['Ime',1]= 'Vladimir'
p table1['Ime'][1]

#Biblioteka omogućava direktni pristup kolonama, preko istoimenih metoda.
p table1.Godina

#Subtotal/Average  neke kolone se može sračunati preko sledećih sintaksi t.prvaKolona.sum i t.prvaKolona.avg
p table1.Godina.sum
p table1.Godina.avg

#Iz svake kolone može da se izvuče pojedinačni red na osnovu vrednosti jedne od ćelija(t.indeks.rn2310)
p table1.Indeks.rn1022

#Kolona mora da podržava funkcije kao što su map, select,reduce. Naprimer: t.prvaKolona.map { |cell| cell+=1 }
p table1.Godina.reduce { |year| year + 1 }
p table1.Godina.map { |year| year + 1 }
p table1.Grupa.select { |group| group > 302 }

#Biblioteka prepoznaje ukoliko postoji na bilo koji način ključna reč total ili subtotal unutar sheet-a, i ignoriše taj red
p table3.array2D

#Moguce je sabiranje dve tabele, sve dok su im headeri isti. Npr t1+t2, gde svaka predstavlja, tabelu unutar jednog od worksheet-ova. Rezultat će vratiti novu tabelu gde su redovi(bez headera) t2 dodati unutar t1. (SQL UNION operacija)
table1+table2
p table1.array2D

table1+table3
p table1.array2D

#Moguce je oduzimanje dve tabele, sve dok su im headeri isti. Npr t1-t2, gde svaka predstavlja reprezentaciju jednog od worksheet-ova. Rezultat će vratiti novu tabelu gde su svi redovi iz t2 uklonjeni iz t1, ukoliko su identicni.
table1-table2
p table1.array2D

table1-table3
p table1.array2D

#Biblioteka prepoznaje prazne redove, koji mogu biti ubačeni izgleda radi
p table3.array2D


