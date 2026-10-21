# Bábolna Sped – Alvállalkozói jutalékelszámolás v2.4.0

## Napi használat
1. Töltsd fel a teljes eredeti XLSX-fuvarlistát, a korábbi három hónap adataival és minden fuvarszervezővel.
2. Automatikusan visszatöltődnek a hónapdöntések, és megjelenik az előző három naptári hónap ellenőrzése.
3. Válassz elszámolási hónapot és fuvarszervezőt. A még eldöntetlen, több hónapot érintő fuvarokhoz rendelj hónapot.
4. Kattints az Excel letöltése vagy PDF letöltése gombra. A kész állapot automatikusan a Drive-ra mentődik és visszaellenőrződik, majd elindul a letöltés.

Nincs havi lezárás, kézi archívummentés vagy korábbi kifizetési kimutatás importja.
A mentés a letöltés kezdeményezésekor készül; a böngészőben történő tényleges helyi fájlmentés nem igazolható a szerverről.
Drive-hiba vagy hiányos számítás esetén nem indul letöltés kész mentés nélkül.
A mentés eredményét és az esetleges hibát a letöltési gombok alatt látod.

## Automatikus mentés
A kész mentés a kiválasztott fuvarszervező adott havi kimutatását jelenti.
Egy másik fuvarszervező vagy másik hónap állapotát nem teszi automatikusan kész állapottá.
A dátumozott ZIP tartalmazza a teljes eredeti source.xlsx fájlt, a beállításokat és számításokat tartalmazó snapshot.json fájlt,
valamint a kész report.xlsx és report.pdf fájlt. Ezek a Drive mappájában a ZIP letöltése és kicsomagolása után elérhetők.
A kimutatások tartalma ellenőrzőösszeggel is védett. A korábbi mentéseket az alkalmazás nem írja felül és nem törli.
Azonos változat Excel- és PDF-letöltése ugyanabban a munkamenetben közös mentést használ.
A feltöltés, előnézet és elemzés önmagában nem készít új kész állapotot.
A letöltés nem jelent kifizetést. A kimutatás mindig a teljes számított jutalékot mutatja.

## Hónapdöntések és átsorolások
A több hónapot érintő fuvar hónapdöntése a fő fülön adható meg.
A Fuvar átsorolása fülön teljes fuvarszám alapján bármelyik megfelelő hónap kiválasztható.
A döntés azonnal, külön JSON-bejegyzésként, visszaellenőrzéssel a Drive-ra mentődik.
Ez nem kész pénzügyi állapot, csak a besorolás megőrzése.
Az átsorolás visszavonása szintén külön bejegyzés; régi mentésből nem térhet vissza a visszavont beállítás.
A fuvarszámhoz kötött hónapdöntés új forrás feltöltése és új belépés után is megmarad.
Dátum- és összegváltozás után is ugyanaz a hónap érvényes, amíg kifejezetten át nem sorolod.
Az eredeti forrásdátum változását ettől függetlenül jelzi a visszaellenőrzés.
A kézi átsorolás elsőbbséget élvez a többdátumos hónapdöntéssel szemben.
Sikertelen Drive-mentésnél nincs sikeres helyi döntés; a művelet ugyanazzal az azonosítóval újrapróbálható.

## Visszaellenőrzés
Mindig az aktuális naptári hónapot megelőző három hónapot vizsgálja.
2026 októberében például júliust, augusztust és szeptembert.
Hónaponként és fuvarszervezőnként a legutóbb letöltött kész kimutatáshoz hasonlít.
A korábbi verzióban kézzel mentett teljes havi archívumok is felhasználhatók.
Egy összevont jelzés mutatja az eltérő és a nem teljesen ellenőrizhető hónapokat; részletek lenyithatók.
Jelzi az új vagy hiányzó fuvart, a módosult forrásmezőt, dátumot, hónapot, fuvarszervezőt, árfolyamot és jutalékot.
A hiányzó fuvar szűkebb exportot is jelenthet, ezért nem automatikusan törlés.
Ha nincs kész mentés, vagy egy fuvarszervezőhöz még nem készült kimutatás, nem igazol eltérésmentességet.
A különbözet nem automatikus ráfizetés vagy visszavonás.
Az Automatikus ellenőrzés frissítése gombbal változatlan forrás mellett is újraolvashatók a Drive-on lévő mentések és döntések.

## Számítás és export
A pozitív árrés 10%-a a jutalék, tételenként egész forintra, a feleket nullától távolodva kerekítve.
Negatív vagy nulla árrés jutaléka 0 Ft. A negatív sorok pirosak; darabszámuk és részleteik csak akkor látszanak, ha vannak.
Azonos fuvardíj- és alvállalkozói pénznemnél az exportált extra árrés az alap.
Vegyes HUF/EUR pénznemnél a forintosított fuvardíj mínusz a forintosított alvállalkozói díj; a kész árrést nem vesszük át, külön extra számlasorokat nem adunk hozzá.
A forintosításhoz az elszámolási hónap pontos MNB havi átlagát használjuk, a kijelzés két tizedes.
Az árfolyam kézi rögzítése csak automatikus lekérdezési hiba esetére szolgál, forrással és indoklással.
A folyamatban lévő hónap teljes havi átlaga még nem ismert. Hiányos kimutatás nem tölthető le kész állapotként.

Az Excel első lapja az összesítő, a PDF első oldala kártyás összesítő; utánuk a részletező következik.
A dátumok pontokkal tagoltak; a díjaktól az oszlopok középre igazítottak; a jutalék után Ft áll.
A letöltött fájl neve tartalmazza az elszámolási hónapot, fuvarszervezőt és a létrehozási dátumot.

## Jutalékelemzés
Egy vagy több fuvarszervező és hónap választható.
Megjelenik a havi érték, a kiválasztott hónapok összege és átlaga, valamint a személyenkénti bontás.
Az összegek ezres tagolással és Ft jelöléssel látszanak.
Az átlag nevezője a kiválasztott hónapok száma, az üres kiválasztott hónapokkal együtt.
Több fuvarszervezőnél a fő átlag a közös havi összeg átlaga.

## Belépés és adatmegőrzés
A forrás feltöltése munkamenetenként történik; új belépés után töltsd fel újra a fuvarlistát.
A hónapdöntések és kész kimutatások a közös Drive-mappában maradnak.
Jelszó és szolgáltatásfiók-kulcs kizárólag a Streamlit Secretsben legyen; GitHubra ne kerüljön titok vagy üzleti Excel.
