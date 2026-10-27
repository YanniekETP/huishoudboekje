# ET WebApp Starter — Google Apps Script + Sheets

Kopieer deze map naar een **nieuwe** repository. Deze starter bevat geen echte spreadsheetgegevens of toegangssleutels.

## Eenmalig per app
1. Maak een nieuwe Google Sheet (indien nodig) en een nieuw Apps Script-project.
2. Leg `SCRIPT_ID` en `DEPLOYMENT_ID` vast als GitHub Actions secrets. Maak de eerste webapp-deployment eenmalig in Apps Script.
3. Voeg `CLASPRC_JSON` toe als GitHub secret, afkomstig van de geautoriseerde clasp-login. Deel deze waarde niet in chat.
4. Kopieer `deploy.yml` naar `.github/workflows/deploy.yml`.
5. Kopieer `Code.gs`, `index.html`, `tests/smoke.test.cjs` naar de repo.
6. Vervang `REPLACE_ME_SPREADSHEET_ID` en pas de datalaag aan.

## Veiligheid
- Gebruik een aparte deployment per app. Een andere deployment-ID kan de verkeerde productie-app overschrijven.
- GitHub Actions controleert syntax en tests vóór publicatie.
- Een groene workflow betekent nog niet dat de live UI functioneel is; doe een live rooktest.
- Autorisaties en toegangsrechten zijn eenmalige stappen per nieuw Google-project en kunnen niet altijd volledig worden geautomatiseerd.

## Uitbreiden
Voeg specifieke datafuncties toe in `Code.gs`, UI in `index.html` en regressietests in `tests/`. Vermijd het direct laden van HTML uit de GitHub-main-branch in productie: zo voorkom je een mismatch tussen frontend en backend.
