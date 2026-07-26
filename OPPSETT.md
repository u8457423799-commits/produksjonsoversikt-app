# Oppsett av Produksjonsoversikt

Denne appen skal publiseres parallelt med dagens dashboard. Ikke endre eller
slett `Dashbord2` før den nye appen er ferdig testet.

## 1. Opprett et nytt GitHub-repository

1. Opprett et nytt offentlig repository med navnet `produksjonsoversikt-app`.
2. Pakk ut ZIP-pakken.
3. Last opp alt innholdet fra mappen `produksjonsoversikt-app`.
4. Kontroller at `.github/workflows/deploy-pages.yml` er med.

Repositoryet er offentlig fordi GitHub Pages skal levere appskallet gratis.
Pakken inneholder ingen produksjonsdata, passord eller hemmelige nøkler.

## 2. Slå på GitHub Pages

1. Åpne det nye repositoryet.
2. Velg `Settings → Pages`.
3. Under `Build and deployment`, velg `GitHub Actions` som kilde.
4. Åpne fanen `Actions` og vent til `Publiser produksjonsapp` er grønn.

Forventet adresse:

`https://u8457423799-commits.github.io/produksjonsoversikt-app/`

## 3. Legg inn redirect URI i Microsoft Entra

1. Åpne `Microsoft Entra ID → App registrations → Produksjonsoversikt`.
2. Velg `Authentication`.
3. Trykk `Add a platform`.
4. Velg `Single-page application`.
5. Legg inn:

   `https://u8457423799-commits.github.io/produksjonsoversikt-app/`

6. Trykk `Configure` eller `Save`.

Ikke opprett Client secret. Ikke slå på Implicit grant.

## 4. Kontroller SharePoint-tilgangen

Filen som leses er:

`Shared Documents/Produksjonsoversikt/produksjon2.json`

Hver ansatt som skal bruke appen, må ha lesetilgang til filen eller mappen.
Bruk en Microsoft 365- eller sikkerhetsgruppe dersom mange skal ha tilgang.
Ikke bruk en anonym «Alle med koblingen»-lenke.

## 5. Test

1. Åpne appadressen.
2. Trykk `Logg inn med Microsoft`.
3. Logg inn med jobbkonto.
4. Godta `User.Read` og `Files.Read` dersom Microsoft spør.
5. Kontroller at produksjonsdata og tidspunkt for siste oppdatering vises.
6. Test med en annen ansatt som har lesetilgang.
7. Test også med en bruker som ikke har tilgang; data skal ikke vises.

## 6. Installer

Android:

- Åpne appen i Chrome.
- Velg `Installer app` eller `Legg til på startskjermen`.

iPhone:

- Åpne appen i Safari.
- Trykk `Del`.
- Velg `Legg til på Hjem-skjermen`.

## 7. Overgang etter godkjent test

Når den nye appen fungerer for alle aktuelle brukere:

1. Stopp GitHub-oppdateringen i Power Automate, men behold SharePoint-oppdateringen.
2. Gjør det gamle `Dashbord2`-repositoryet privat.
3. Bekreft at den nye appen fortsatt oppdateres hvert femte minutt.
