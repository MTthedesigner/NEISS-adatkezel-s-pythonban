https://www.cpsc.gov/cgibin/NEISSQuery/home.aspx
# NEISS sportserülések elemzése (2016–2025)

## alkalmazott libaryk:
-contextlib ---> a fuggvenyt lehet vele blokkent hasznalni neiisadatokmegnyit megnyitja az excelt <br />
-pathlib ---> a mappautvonalak es fileok kezelesere (excel fileok + eredmeny mappa)<br />
-io ---> ha az excel "hibas" (maskepp van az elrendezes, stb) a javitott verziot tarolja a memoriaban<br />
-zipfile ---> metaadat javitasahoz<br />
-xml.etree ---> hibas datummezok eltavolitasa<br />

-matplotlib.pyplot ---> diagram<br />
-pandas ---> adatkezeleshez , cvs-be mentes<br />
-openpyxl ---> excel lapok es sorok beolvasasa<br />
-matplotlib.ticker ---> diagram<br />

## program mukodese:
sportagak kigyujtve egy halmazba ami alapjan keresi a program a nevet es ami neve legyen

egysegesites: NEISS adatokat egyseges szoveges formara alakitja

neiisadatokmegnyit: Megnyitja az excelt hibas datum-metaadatnal memoriabeli masolatot keszit + javit

beolvasas: Beolvassa az adott ev termekkodjainak hivatalos megnevezeset

sportrend: Sportagat rendel a NEISS termekkod hivatalos nevehez

serulesbetolt: afuggveny beolvassa a megadott excel fajlokat kivalasztja a sporthoz kotheto seruleseket es egy pandas tablazatba visszaadja

diagment: a beolvasott serulesi adatokbol elkesziti es kepkent elmenti a sportagankenti ferfi no bontasu diagramot

# alkamlazott irodalmak:
https://www.geeksforgeeks.org/python/reading-excel-file-using-python/
https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html
https://docs.python.org/3/library/contextlib.html
https://matplotlib.org/stable/tutorials/pyplot.html
https://matplotlib.org/stable/api/ticker_api.html
https://openpyxl.readthedocs.io/en/stable/
https://docs.python.org/3/library/pathlib.html

# A kapott eredmeny:
Besorolt NEISS-esetek: 616 691, 
Országos becslés a Weight mező alapján: 21 180 647
<img width="2080" height="2083" alt="sportag_es_nem_szerint" src="https://github.com/user-attachments/assets/70cfdeaa-c9bb-4f64-98f7-aa7628ade983" />
<img width="594" height="957" alt="image" src="https://github.com/user-attachments/assets/57ba9aaa-12b1-4272-acf6-e556391aa3f5" />

