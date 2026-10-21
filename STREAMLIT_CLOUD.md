# Streamlit Community Cloud – telepítés

1. A csomag Babolna_Sped_Jutalek mappájának tartalmát tedd egy GitHub-repozitórium gyökerébe.
   Az app.py, requirements.txt, core.py, settlement.py, ui_features.py, web_session.py, exports.py, drive_archive.py, archive_ui.py, auto_check.py, persistence.py, runtime_cache.py,
   assets mappa és .streamlit/config.toml maradjon együtt. A .streamlit mappa rejtett lehet a fájlkezelőben.
2. Nyisd meg a https://share.streamlit.io/ oldalt és hozz létre új alkalmazást.
3. Válaszd ki a repozitóriumot, az ágat és belépési fájlként az app.py-t.
4. Az Advanced settings alatt válassz Python 3.12-t. A Secrets mezőben állíthatsz jelszót:
   `app_password = "a-sajat-jelszavad"`. Ehhez mintát ad a .streamlit/secrets.toml.example.
5. Indítsd el a telepítést (Deploy). A függőségeket a requirements.txt alapján telepíti a rendszer.

## Frissítés v2.4.1-re
A csomag összes alkalmazásfájlját frissítsd, különösen a persistence.py, runtime_cache.py és requirements.txt fájlt.
A kattintásra előállított, mentéssel összekötött letöltéshez Streamlit 1.59 vagy újabb szükséges.
A meglévő Streamlit Secrets és Drive-mappa változatlanul használható.

## Webes használat
- Excel vagy PDF letöltésekor automatikus, visszaellenőrzött kész mentés kerül a Drive-ra.
- A hónapdöntések és átsorolások külön, azonnal mentődnek, visszavonásaikkal együtt.
- Az előző három naptári hónap visszaellenőrzése automatikus.
- Nincs Drive-archívum vagy Korábbi kimutatás főfül, kézi állapotmentés és kifizetési nyilvántartás.
- Új belépés után töltsd fel a forrást; a mentett hónapdöntések automatikusan visszatöltődnek.
- Drive-kapcsolat nélkül az előnézet elérhető, de kész kimutatás nem tölthető le mentés nélkül.
- Részletes használat: HASZNALAT.md. Drive-beállítás: DRIVE_ARCHIVUM_BEALLITAS.md.

Hivatalos dokumentáció:
https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/file-organization
https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/app-dependencies

## A telepítési csomag tartalma
A ZIP kizárólag a Streamlit Cloud alkalmazásfájljait, szükséges képeket és betűkészleteket,
a konfigurációt, Secrets-mintát és használati/telepítési útmutatókat tartalmazza.
Nincs Windows-indító BAT, helyi virtuális környezet, adatbázis, tesztfájl vagy requirements-dev.txt.
Ha korábbi feltöltésből az INDÍTÁS.bat, requirements-dev.txt vagy test_*.py fájlok még a GitHub-tárolóban vannak,
azokat töröld; egy új csomag feltöltése a régi fájlokat önmagában nem távolítja el.

## Futási optimalizálás
A forrás beolvasása tartalmi változásig újrahasznosítható ugyanabban a munkamenetben.
A számítás és elemzés a forrás, beállítások vagy kiválasztás változásakor számol újra.
A hónapbesorolások lekérdezése kötegelt. A Drive-tartalom csak változott módosítási időnél olvasódik újra;
az archívumlista frissítése megmarad. Az új mentésekből hónaponként és személyenként a legutóbbi szükséges változatot olvassa,
a régi archívumformátumok kompatibilitása mellett. A gyorsítótárak munkamenetenként különülnek el.
A letöltési állapot lekérdezése csak folyamatban lévő letöltésnél ismétlődik.
