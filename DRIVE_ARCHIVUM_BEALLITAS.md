# Automatikus Drive-mentés beállítása – v2.4.1

Excel vagy PDF letöltésekor a kiválasztott fuvarszervező adott havi kész kimutatása automatikusan mentődik.
A ZIP tartalma az eredeti forrás, beállítások, számítások, kész Excel és PDF.
A hónapdöntések és átsorolások külön JSON-bejegyzésként azonnal mentődnek, a visszavonásokkal együtt.
Azonos forrás és kimutatás Excel/PDF-letöltése közös mentést használ. A mentés nem kifizetés vagy havi lezárás.

## 1. Frissítsd az alkalmazás fájljait
A csomag Babolna_Sped_Jutalek mappájának tartalmát másold a GitHub-tároló gyökerébe.
Új a persistence.py; frissítendő az app.py, core.py, settlement.py, ui_features.py,
web_session.py, exports.py, drive_archive.py, archive_ui.py, auto_check.py, runtime_cache.py és requirements.txt is.
A meglévő Secrets és mappa változatlanul használható. Nincs szükség külön archívumfülre vagy kézi mentésre.
Titkot vagy üzleti Excelt ne tölts fel a GitHubra.

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

## 5. Első kész állapot és automatikus ellenőrzés
1. Töltsd fel a teljes eredeti fuvarlistát a fő fülön.
2. Rendezd a több hónapos fuvarokat; a döntések azonnal, visszaellenőrizve mentődnek a Drive-ra.
3. Válassz hónapot és fuvarszervezőt, majd töltsd le az Excelt vagy PDF-et.
4. A letöltés előtt elkészül a kész állapot mentése. A letöltés csak visszaellenőrzött mentés után indul.

Nincs külön archívumfül, kézi jóváhagyás vagy mentésgomb.
Minden letöltött kimutatás az adott személy adott havi összehasonlítási alapját frissíti.
Azonos változat két formátumának letöltése közös mentést használ egy munkamenetben.
A ZIP-ben a teljes forrás és beállítások mellett mindkét kész kimenet megmarad.
A hónapdöntések JSON-fájljai önmagukban nem kész pénzügyi állapotok.

Új feltöltésnél a fő fül automatikusan ellenőrzi az előző három naptári hónapot,
hónaponként és fuvarszervezőnként a legutóbbi kész állapot alapján.
A korábban kézzel mentett teljes havi archívumok is használhatók.
A hiányzó alapot, hibás számítást, új / hiányzó fuvart, dátum-, hónap-,
fuvarszervező-, árfolyam- és jutalékváltozást jelzi. Hiányzó sor esetén a szűrőket is ellenőrizni kell.
A különbözet nem automatikus kifizetési javaslat. Korábbi kifizetési kimutatást már nem importálunk.

Új belépés után a hónapdöntések automatikusan visszaállnak; a visszavonásuk is megmarad.
Az Automatikus ellenőrzés frissítése gombbal ugyanahhoz a forráshoz is újraolvashatók a mentések.
A korábbi forrás és kész kimenetek a Drive-on lévő ZIP-ekből visszakereshetők.

## Megőrzés és korlátok

- A Drive megőrzi a mentéseket az app újraindulásától függetlenül.
- Az app nem ír felül és nem töröl korábbi archívumot. A megfelelő Drive-
  jogosultsággal rendelkező emberek továbbra is módosíthatnak / törölhetnek
  fájlokat; ez nem jogilag változtathatatlan napló.
- Az archiválás nem indít kifizetést, és nem ment kifizetési nyilvántartást.
- Az app csak az általa készített, megjelölt archívumokat olvassa automatikusan
  a beállított mappából; nem kutatja át a teljes vállalati Drive-ot.
- Átmeneti 500/502/503/504 hibánál és kéréskorlátozásnál az olvasásokat
  rövid várakozással, legfeljebb háromszor próbálja meg az app. Ezeknél nem
  jelenik meg félrevezető, jogosultságmódosítást kérő üzenet.
- Ha a feltöltés válasza bizonytalan, az állapotazonosító alapján keresi a
  már létrejött mentést, majd a tartalmát is visszaellenőrzi. A feltöltést nem
  ismétli meg automatikusan, így a hibakezelés nem készít másodpéldányt.
- Ha a visszaellenőrzés az újrapróbálások után is sikertelen, nem jelez
  igazolt mentést. Megkülönbözteti a már létrejött fájlt a még bizonytalan
  feltöltéstől. Ilyenkor próbáld újra ugyanazt a letöltést.
  Azonos állapot ismétlése meglévő azonosítóval
  visszaellenőrizhető anélkül, hogy új fájlt töltene fel.

Hivatalos háttér:
https://developers.google.com/workspace/drive/api/guides/about-shareddrives
https://developers.google.com/workspace/drive/api/guides/enable-shareddrives
https://developers.google.com/workspace/drive/api/guides/manage-uploads
https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/secrets-management
