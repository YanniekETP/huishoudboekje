---
name: et-webapp-builder
description: Bouw en onderhoud Google Apps Script-webapps met Google Sheets, GitHub, clasp en geautomatiseerde controles.
---
# ET WebApp Builder

Gebruik deze werkwijze wanneer een nieuwe Google Apps Script-webapp wordt gebouwd of een bestaande wordt aangepast.

## Uitgangspunten
- Hergebruik de bestanden in `starter/template/`; kopieer nooit persoonlijke data, spreadsheet-ID's, deployment-ID's of repository-secrets.
- Werk vanuit één GitHub-repository per app, met `Code.gs`, `index.html`, `tests/` en een deployment-workflow.
- Verander bestaande productie-apps alleen na een expliciet verzoek.
- Vraag autorisatie uitsluitend wanneer de gebruiker echt een Google- of GitHub-toestemming moet geven.
- Een commit is **geen** bewijs dat de webapp gepubliceerd is.

## Nieuwe app
1. Inventariseer het doel, de bron (bijv. Sheet), gebruikers en toegangsrechten. Vraag alleen ontbrekende essentiële gegevens.
2. Maak een nieuwe repo vanuit `starter/template/`. Gebruik een nieuw Apps Script-project en nieuwe deployment. Hergebruik nooit bestaande IDs.
3. Configureer GitHub Actions secrets: `CLASPRC_JSON`, `SCRIPT_ID`, `DEPLOYMENT_ID`. Sla credentials nooit op in broncode of logs.
4. Maak een minimale versie: `doGet`, één veilige data-API, UI met loading/success/error/timeout, en healthcheck.
5. Draai lokale syntaxis- en unit-tests. Controleer dataformaten, lege resultaten, formules en foutafhandeling.
6. Publiceer via CI; controleer de nieuwste workflow-run, deployment-ID en live URL.
7. Voer een rooktest uit op de live app (laden, data lezen, foutmelding, gebruikersrechten). Claim geen succes zonder verificatie.

## Wijziging aan bestaande app
1. Lees eerst de actuele GitHub-bestanden en relevante Sheets-ranges.
2. Leg de huidige werking vast met tests, inclusief een test voor de gevraagde wijziging.
3. Wijzig alleen de noodzakelijke code; bescherm spreadsheetformules en bestaande data.
4. Valideer alle inline JavaScript uit HTML met een parser/syntaxischeck. Controleer ook element-ID's.
5. Commit; wacht op geslaagde CI; verifieer de live versie.
6. Bij fout: lees Actions-logs, herstel gericht en test opnieuw. Vraag geen screenshot als logs toegankelijk zijn.

## Verplichte kwaliteitsregels
- Geen oneindige laadschermen: foutafhandeling en timeout.
- Geen hardcoded financiële waarden als brondata beschikbaar zijn.
- Data niet stilzwijgend muteren; read-only testen waar mogelijk.
- Behoud van originele rekenregels; UI-sortering mag bronvolgorde niet veranderen.
- Productie en test gescheiden; geen automatische overschrijving van andere apps.
- Rapporteer duidelijk: gebouwd, getest, gepubliceerd en live geverifieerd zijn afzonderlijke statussen.

## Hergebruik
Start met `starter/template/README.md`. Vervang alle `REPLACE_ME`-waarden. Gebruik de workflow in `starter/template/deploy.yml` als uitgangspunt; pas de Apps Script-bestandsnaam aan aan de daadwerkelijke clasp-projectstructuur.
