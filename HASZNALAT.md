# Bábolna Sped – Alvállalkozói jutalékelszámolás v2.2.0

## Streamlit webes használat
A telepítés lépései a STREAMLIT_CLOUD.md fájlban találhatók. A belépési fájl app.py.
A feltöltött adatok munkamenetenként elkülönülnek. Nincs havi lezárás vagy lezárási archívum.

## Kimutatás készítése
Töltsd fel a teljes, több hónapot tartalmazó eredeti Fuvarlistát. Válassz hónapot és kapcsolattartót.
Nincs havi lezárás. A több hónapot érintő dátumokhoz rendelj hónapot az Ellenőrzések részen.
Mentés után sikerüzenet jelenik meg, a rendezett fuvar kikerül az ellenőrizendők közül.

A pozitív árrés jutaléka tételenként 10%, egész forintra, a feleket nullától távolodva kerekítve.
Negatív vagy nulla árrés esetén 0 Ft. A negatív sorok pirossal megmaradnak.
Azonos fuvardíj- és alvállalkozói pénznemnél az exportált Árrés extra számlasorokkal értéket használjuk.
Vegyes HUF/EUR pénznemnél: a fuvardíj forintértéke mínusz az alvállalkozói díj forintértéke.
Az EUR-összegek az elszámolási hónap pontos MNB havi átlagával kerülnek forintosításra.
Vegyes pénznemnél a kész exportált árrés nem mérvadó, és külön extra számlasorok nem adódnak hozzá.
A részletező Árrés extrákkal mezőjében ezeknél a számított HUF-árrés látható.
Teljes pontosságú árfolyammal számolunk, de a felületen, PDF-ben és Excelben két tizedes látszik.
A még folyamatban lévő hónap teljes havi átlagárfolyama nem ismert. Hibás vagy nem számítható tételeknél a kimutatás hiányosként jelölt.

A PDF első oldala kártyás összesítő, utána fekvő A4-es részletező. Az Excel első lapja Összesítő.
A dátumok ÉÉÉÉ.HH.NN formában jelennek meg. A díjaktól minden oszlop középre igazított, a jutalék után Ft áll.
A fájlnév: Jutalek_elszámolási-hónap_fuvarszervező_létrehozási-dátum.xlsx, illetve .pdf.

## Fuvar átsorolása
A külön Fuvar átsorolása fülön add meg a teljes fuvarszámot, majd válaszd ki az új hónapot és mentsd.
Az átsorolás minden kimutatásban és az elemzésben érvényes. A fuvar egyetlen hónapban jelenik meg.
Az eredeti kiegyenlítési dátum megmarad, de az új elszámolási hónap árfolyamával számolunk.
A kézi átsorolás visszavonható. A korábbi kifizetési előzményt az átsorolás nem törli.

## Korábbi kifizetések ellenőrzése
A Korábbi kimutatás (opcionális) fülön válaszd ki a fuvarszervezőt és tölts fel korábban kifizetett XLSX-kimutatást.
A régi kézi táblázat és az alkalmazás Excelje is használható. Add meg a kifizetés dátumát,
és külön erősítsd meg, hogy a pozitív jutalékok ténylegesen ki lettek fizetve.
A feltöltött kimutatás önmagában nem banki bizonylat. A letöltés nem jelöl automatikusan kifizetést.
A nyilvántartott fuvar újra fizetendő jutaléka 0 Ft, kifizetve jelzéssel. Ugyanaz a fuvar nem importálható kétszer.
A védelem fuvarszám alapján, másik hónapra vagy fuvarszervezőre átsorolás után is érvényes.
A nulla jutalékos sorokat nem rögzítjük kifizetésként. Hibás vagy negatív kifizetést tartalmazó fájlt javításig nem lehet importálni.
A képletek mentett eredményét olvassuk; ha az hiányzik, előbb Excelben vagy Sheetsben újraszámolva mentsd a kimutatást.
Ha a korábbi összeg vagy fuvarszervező eltér a jelenlegi számítástól, figyelmeztetés jelenik meg.
Automatikus különbözeti ráfizetés nincs. A téves kifizetési jelölés külön, megerősítéssel visszavonható és naplózva marad.
A Kifizetési ellenőrzés Excel-lap mutatja a tényleges történeti összeget és a forrást.

## Jutalékelemzés
Válassz egy vagy több fuvarszervezőt és egy vagy több hónapot.
A diagramokon, tooltipben és táblázatokban az összegek ezres tagolással és Ft jelöléssel láthatók.
Például: 1 234 567 Ft; a havi átlag például 479 974,10 Ft.
Külön szerepel a havi jutalék, a kijelölt hónapok összege és átlaga, valamint a személyenkénti összeg és átlag.
A normál elemzésben nincsenek kifizetési állapotok vagy fizetendő oszlopok: a teljes számított jutalékot mutatja.
Az átlag nevezője a kijelölt hónapok száma, az üres kiválasztott hónapokat is beleértve.
Több fuvarszervezőnél a fő átlag a közös havi összeg átlaga.

## Webes munkamenet és mentés
A felhasználók adatai külön munkamenetben vannak, nem közös cloud-adatbázisban.
Új munkamenet vagy szerver-újraindítás után töltsd fel újra a fuvarlistát.
A bal oldali JSON-mentés az átsorolásokat, dátumdöntéseket és árfolyamokat tartalmazza.
Visszatöltés: Használat és beállítások → Mentett beállítások visszatöltése.
A mentés nem tartalmaz fuvarlistát vagy kifizetési állapotot. Korábbi kifizetés ellenőrzéséhez
az opcionális fülön töltsd fel újra a kifizetett kimutatást.
Jelszó a Streamlit Secretsben: app_password. A tényleges secrets.toml fájlt ne tedd GitHubra.

## Drive-archívum és korábbi hónapok összehasonlítása
A v2.2.0 külön Drive-archívum fület ad. A beállításokat a
DRIVE_ARCHIVUM_BEALLITAS.md részletesen tartalmazza. A funkció nélkül
a korábbi munkafolyamat változatlanul használható.

A csatlakoztatás után töltsd fel a teljes forráslistát, nyisd meg a Drive-archívumot,
és válaszd ki az ellenőrzött hónapokat. A mentés minden érintett fuvarszervezőre
vonatkozik. Teljességi megerősítés után külön, dátumozott állapotot menthetsz.
A korábbi fájlokat az app nem írja felül. A mentés nem lezárás és nem kifizetés.

Következő feltöltéskor a megnyitott archívum a kiválasztott korábbi állapottal
vet össze: új / hiányzó fuvarok, változott mezők, hónap- és fuvarszervezőváltás,
árfolyamkülönbség, számított jutalék régi és új értéke. Havi és személyenkénti
bontás is látható. Csak a mindkét állapotban ellenőrzött hónapokat hasonlítjuk.
A hiányzó fuvar lehet szűkebb export következménye, ezért nem minősítjük törlésnek.
A különbözetből nincs automatikus ráfizetés vagy visszavonás.

Az archívum a teljes eredeti Excelt és a beállításokat is őrzi. Korábbi eredeti
forráslistát első alapként feltölthetsz és elmenthetsz. Az archivált forrás
letölthető, az átsorolások és árfolyamok külön megerősítéssel átvehetők az új
munkamenetbe. Kifizetési jelölés nem mentődik és nem állítódik vissza.
