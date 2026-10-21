# Drive-archívum beállítása – v2.2.0

Az új funkció külön menti az eredeti forrás-Excelt, a pontos havi árfolyamokat,
a hónapdöntéseket, az átsorolásokat és a kijelölt hónapok számított jutalékait.
A mentés nem lezárás és nem kifizetés. Az app korábbi állapotot nem ír felül.

## 1. Frissítsd az alkalmazás fájljait

A csomag `Babolna_Sped_Jutalek` mappájának tartalma kerüljön a meglévő
GitHub-tároló gyökerébe. Új fájlok: `drive_archive.py`, `archive_ui.py`,
`DRIVE_ARCHIVUM_BEALLITAS.md`, `test_drive_archive.py`.
Módosult az `app.py`, `requirements.txt`, `HASZNALAT.md`, `STREAMLIT_CLOUD.md`
és a `.streamlit/secrets.toml.example` minta.
A `core.py` számítási szabályai nem változtak.
Valódi titkot, privát kulcsot vagy forrás-Excelt ne tölts fel a GitHubra.

## 2. Válaszd ki a céges megosztott meghajtót

A Google Drive **Megosztott meghajtók** részében válassz meglévő céges
meghajtót, vagy hozz létre egyet például `Bábolna Sped elszámolások` néven.
Abban készíts `Jutalék archívum` mappát.

Ez NEM egy Saját meghajtó alatti megosztott mappa. Az alkalmazás
szolgáltatásfiókkal dolgozik; ennek nincs saját Drive-tárhelye, ezért az írás
céges megosztott meghajtóra történik. Business Standard támogatja ezeket.
Ha ezt a lehetőséget nem látod, a Workspace-rendszergazda segítsége szükséges.
Az alkalmazás mappaszinten ellenőrzi a megfelelő típust és az írási jogosultságot.

A mappa URL-je például `https://drive.google.com/drive/folders/MAPPA_AZONOSITO`.
A `folders/` utáni azonosítót jegyezd fel. Nem a meghajtó azonosítója kell.

## 3. Állítsd be a Google Cloud-hozzáférést

1. A https://console.cloud.google.com/ oldalon hozz létre vagy válassz projektet.
2. Engedélyezd a **Google Drive API** szolgáltatást.
3. Az **IAM & Admin → Service Accounts** oldalon készíts szolgáltatásfiókot.
   Az alkalmazáshoz nem kell tulajdonosi Google Cloud-projektszerepet adni.
4. Készíts hozzá JSON-kulcsot. A kulcs titok, nem felhasználói adatfájl;
   csak a Streamlit Secretsben használd, és ne küldd el üzenetben.
5. A JSON `client_email` címét add hozzá a célmappához / megosztott meghajtóhoz
   olyan hozzáféréssel, amely fájlok listázását, olvasását és hozzáadását engedi
   (Contributor / Közreműködő vagy a Workspace által megengedett megfelelő szerep).
   Nyilvános linkmegosztást nem kell és nem javasolt bekapcsolni.

Ha a céges szabály tilt szolgáltatásfiók-kulcsot vagy ilyen megosztást,
ne lazítsd fel emiatt az egész szervezet beállításait. A rendszergazdával
egyeztetett hozzáférés kell; OAuth-alapú alternatíva külön fejlesztés lenne.

## 4. Streamlit Secrets

Az alkalmazás **Settings → Secrets** mezőjében a már működő
`app_password` sort tartsd meg. Az alábbi blokkokat add hozzá;
a mintaértékeket a saját adataidra cseréld:

```toml
# A meglévő belépési jelszó maradjon a fájl tetején, a táblákon kívül.
app_password = "A_MEGLEVO_SAJAT_JELSZAVAD"

[drive_archive]
# A céges megosztott meghajtón belüli célmappa azonosítója.
folder_id = "MAPPA_AZONOSITO"

[gcp_service_account]
# Az értékeket a Google Cloud JSON-kulcsából másold.
type = "service_account"
project_id = "SAJAT_PROJEKT"
private_key_id = "A_JSON_PRIVATE_KEY_ID_ERTEKE"
# A három idézőjel közötti tartalom a JSON private_key értéke,
# tényleges sortörésekkel; a BEGIN és END sor is maradjon benne.
private_key = """-----BEGIN PRIVATE KEY-----
IDE_KERUL_A_SAJAT_TITKOS_KULCS
-----END PRIVATE KEY-----
"""
client_email = "SZOLGALTATASFIÓK@SAJAT_PROJEKT.iam.gserviceaccount.com"
client_id = "A_JSON_CLIENT_ID_ERTEKE"
token_uri = "https://oauth2.googleapis.com/token"
```

Mentsd a Secrets tartalmát, majd indítsd újra az appot, ha azt kéri.
Új függőség a `google-auth`; ezt a friss `requirements.txt` telepíti.

## 5. Első ellenőrzött állapot

1. Töltsd fel a teljes eredeti fuvarlistát a szokásos fő fülön.
2. Rendezd a több hónapos fuvarok hónapját; javítsd a hibás sorokat.
3. A **Drive-archívum** fülön jelöld be a **Drive-archívum megnyitása** mezőt.
4. Válaszd ki az **Ellenőrzött hónapok** listáját. A mentés ezek minden
   fuvarszervezőjére vonatkozik, nem csak a fő fülön kiválasztott személyre.
5. Ellenőrizd, hogy az export a hónapok minden fuvarját tartalmazza.
   Ezt a program egy szűrt exportból nem tudja biztosan eldönteni.
6. Erősítsd meg a teljességet, majd **Ellenőrzött állapot mentése**.
7. Csak a Drive-ból visszaolvasott és egyeztetett mentés kap sikeres jelzést.

Korábbi eredeti forrás-Excelt is feltölthetsz és menthetsz első állapotként.
Kézzel készített összesítő vagy korábbi kifizetési kimutatás nem teljes
forráslista; azt továbbra is a **Korábbi kimutatás (opcionális)** fülön kezeld.
A mentés ideje a tényleges mentés ideje; nem állítja, hogy a forrást már
akkoriban archiválta a rendszer.

## 6. Új feltöltés összehasonlítása

A fő fülön feltöltött új forrást a **Drive-archívum** fülön az alapértelmezetten
legújabb, illetve kiválasztott mentéssel hasonlítjuk össze. A megnyitott
archívum az új forrás vagy a beállítások változásakor újra számol.

Új munkamenetben az **Archivált beállítások átvétele** gombbal visszaállíthatod
a régi átsorolásokat, dátumdöntéseket és pontos árfolyamokat. Előtte megerősítés
szükséges, mert az azonos kulcsú helyi értékeket felülírja. Kifizetési jelölést
nem állít vissza. Az eredeti Excel az archívumból le is tölthető.

A lista mutatja az új / hiányzó fuvarokat, a változott mezők régi és új értékét,
a hónap- és kapcsolattartóváltást, az árfolyamot és az esetleges számítási
verzióeltérést. A havi, fuvarszervezőnkénti bontásban a számított teljes
jutalék különbsége látható, kifizetési levonás nélkül.

Csak mindkét állapotban ellenőrzött hónapok hasonlíthatók össze. Az alapban
nem szereplő hónapokra nincs korábbi viszonyítás. A hiányzó sor nem bizonyít
törlést: ellenőrizd az export időszakát és szűrőit. A negatív eltérés nem
automatikus visszavonás, a pozitív eltérés nem automatikus ráfizetés.

Az új állapot csak a külön mentésgombbal lesz későbbi összehasonlítási alap.
A korábbi állapotok továbbra is választhatók. Több felhasználó mentése
egymás mellett megmarad; lista frissítésével kérd le mások új mentéseit.

## Megőrzés és korlátok

- A Drive megőrzi a mentéseket az app újraindulásától függetlenül.
- Az app nem ír felül és nem töröl korábbi archívumot. A megfelelő Drive-
  jogosultsággal rendelkező emberek továbbra is módosíthatnak / törölhetnek
  fájlokat; ez nem jogilag változtathatatlan napló.
- Az archiválás nem indít kifizetést, és nem ment kifizetési nyilvántartást.
- Az app csak az általa készített, megjelölt archívumokat olvassa automatikusan
  a beállított mappából; nem kutatja át a teljes vállalati Drive-ot.
- Elérési vagy visszaolvasási hibánál ne tekintsd mentettnek az állapotot.
  Frissítsd az archívumlistát. Azonos állapot ismétlése meglévő azonosítóval
  visszaellenőrizhető anélkül, hogy új fájlt töltene fel.

Hivatalos háttér:
https://developers.google.com/workspace/drive/api/guides/about-shareddrives
https://developers.google.com/workspace/drive/api/guides/enable-shareddrives
https://developers.google.com/workspace/drive/api/guides/manage-uploads
https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/secrets-management
