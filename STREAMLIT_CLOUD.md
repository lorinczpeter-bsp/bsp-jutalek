# Streamlit Community Cloud – telepítés

1. A csomag Babolna_Sped_Jutalek mappájának tartalmát tedd egy GitHub-repozitórium gyökerébe.
   Az app.py, requirements.txt, core.py, settlement.py, ui_features.py, web_session.py, exports.py,
   assets mappa és .streamlit/config.toml maradjon együtt. A .streamlit mappa rejtett lehet a fájlkezelőben.
2. Nyisd meg a https://share.streamlit.io/ oldalt és hozz létre új alkalmazást.
3. Válaszd ki a repozitóriumot, az ágat és belépési fájlként az app.py-t.
4. Az Advanced settings alatt válassz Python 3.12-t. A Secrets mezőben állíthatsz jelszót:
   `app_password = "a-sajat-jelszavad"`. Ehhez mintát ad a .streamlit/secrets.toml.example.
5. Indítsd el a telepítést (Deploy). A függőségeket a requirements.txt alapján telepíti a rendszer.

## Webes használat

- Nincs havi lezárás és nincs lezárási archívum.
- A fuvarlisták, átsorolások és az opcionális korábbi kimutatás munkamenetenként elkülönülnek.
- A normál kimutatás és az elemzés kifizetési állapot nélkül használható.
- Csak a Korábbi kimutatás (opcionális) fülre feltöltött, kifizetettként megerősített
  régi Excel aktiválja a dupla kifizetés ellenőrzését az adott munkamenetben.
- Új munkamenetben és szerver-újraindítás után töltsd fel újra a fuvarlistát.
- Az átsorolások és árfolyamok a bal oldali Beállítások mentése (JSON) gombbal továbbvihetők.
  Visszatöltés: Használat és beállítások → Mentett beállítások visszatöltése.
- A JSON-mentés nem tartalmaz fuvarlistát, kifizetési állapotot vagy jelszót.
- Régi kifizetés ellenőrzéséhez új munkamenetben töltsd fel újra a korábbi kimutatást.
- Nincs megosztott, automatikusan tartós cloud-adatbázis. Ha később több felhasználó között
  közös, tartós nyilvántartás szükséges, külön adattárolót kell csatlakoztatni.

A csomag nem tartalmaz éles céges adatot és jelszót. A telepítéshez szükséges fájlok készen állnak;
a repozitórium létrehozása és a Streamlit Cloudhoz kapcsolása még szükséges.

Hivatalos dokumentáció:
https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/file-organization
https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/app-dependencies
