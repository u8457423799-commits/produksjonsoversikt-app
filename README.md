# Produksjonsoversikt

Lukket PWA for Android, iPhone og PC. Ansatte logger inn med Microsoft-jobbkonto,
og appen leser `Produksjonsoversikt/produksjon2.json` fra SharePoint gjennom
Microsoft Graph.

## Sikkerhet

- Ingen produksjonsdata eller passord ligger i kildekoden.
- Appen bruker delegert `Files.Read`.
- Den innloggede brukeren må ha lesetilgang til SharePoint-filen.
- Service worker mellomlagrer ikke Graph- eller SharePoint-data.
- Det skal ikke opprettes en client secret for denne appen.

## Lokal utvikling

```sh
npm install
npm run dev
```

## Bygg

```sh
npm install
npm run build
```

Innholdet i `dist` kan publiseres som en statisk webapp.

Repositoryet inneholder også en GitHub Actions-flyt som bygger og publiserer
appen automatisk med GitHub Pages når kode sendes til `main`.

## Microsoft Entra

Legg den endelige HTTPS-adressen inn under:

`App registrations → Produksjonsoversikt → Authentication → Add a platform → Single-page application`

Redirect URI må være identisk med adressen appen åpnes på, inkludert avsluttende `/`.

Anbefalt repositorynavn er `produksjonsoversikt-app`. Med GitHub-brukeren som
eier dagens dashboard blir forventet redirect URI:

`https://u8457423799-commits.github.io/produksjonsoversikt-app/`
