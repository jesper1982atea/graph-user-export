# Graph User Export

Kör appen helt utan installation på en annan macOS-dator.

## Alternativ 1: Zippad statisk build (rekommenderas)
1. Skapa paket (på din utvecklingsmaskin):
   - `npm run package`
   - Det skapas en fil `graph-user-export-dist.zip` i projektroten.
2. Flytta zip-filen till valfri Mac och packa upp.
3. Öppna `index.html` i den uppackade mappen i valfri modern webbläsare.
   - Vi använder HashRouter så appen fungerar via `file://` URL.
   - Tokenklistra och alla funktioner körs helt i webbläsaren.

Tips: Om Safari blockerar vissa `file://` begäranden, använd Chrome/Edge.

## Publicera en release på GitHub (automatisk zip)
Workflow finns under `.github/workflows/release.yml`. När du pushar en tag som börjar med `v` skapas en release med zip-bilaga.

Snabbt sätt att skapa och pusha en version (ökar patch-versionen):

1. Säkerställ att allt är committat på `main`.
2. Kör:
  - `npm run release:tag`
3. GitHub Actions bygger, zipp:ar och publicerar releasen för taggen.

Vill du styra version själv (ex. v1.1.0):
- `npm version 1.1.0 -m "chore(release): %s"`
- `git push --follow-tags`

Resultat: Gå till GitHub Releases och ladda ner `graph-user-export-dist-<tag>.zip`.

## Alternativ 2: Starta en lokal server (om du får problem med `file://`)
- På måldatorn, om du har Python eller Node:
  - Python 3: `python3 -m http.server 8080` i dist-mappen, öppna `http://localhost:8080/`
  - Node (npx): `npx serve -s` i dist-mappen

## Bygga igen (om källkod ändras)
- Krävde Node/npm på den datorn. Kör:
  - `npm install`
  - `npm run build`

## Behörigheter för Graph
- För att se alla data:
  - Möten: Calendars.Read
  - Chattar: Chat.Read
  - Användare/sök: User.Read.All eller Directory.Read.All (för filtrerade sökningar/utökade fält)

