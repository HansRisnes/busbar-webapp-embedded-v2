# BUSBAR-WEBAPP-EMBEDDED-V2

## Legge endringer inn i GitHub-repoet

1. **Sjekk status**
   ```bash
   git status -sb
   ```
   Da ser du hvilke filer som er endret.

2. **Legg til filene i commiten**
   ```bash
   git add app.js styles.css index.html
   ```
   Juster listen over filer slik at den matcher endringene du har gjort.

3. **Lag en commit**
   ```bash
   git commit -m "Beskriv endringen"
   ```
   Bruk en meningsfull commit-melding som forklarer hva som er gjort.

4. **Skyv endringene til GitHub**
   ```bash
   git push origin <branch-navn>
   ```
   Erstatt `<branch-navn>` med navnet på grenen du jobber på.

5. **Opprett en Pull Request**
   Når endringene ligger på GitHub, opprett en Pull Request slik at de kan gjennomgås og merges inn i hovedgrenen.

Disse stegene sikrer at montasjeverktøyet og øvrige oppdateringer blir en del av prosjektet i GitHub-repoet.

### Må jeg kopiere kommandoene ord for ord?

Nei – se på kommandoene over som maler. Du kan enten kopiere dem rett inn i terminalen eller skrive dem manuelt og tilpasse dem etter behov. For eksempel kan du bruke `git add .` hvis du vil ta med alle endrede filer, eller du kan bare oppgi filnavnene som faktisk er endret.

Hvis du foretrekker grafiske verktøy (f.eks. GitHub Desktop eller Git-integrasjonen i VS Code), kan du gjøre de samme operasjonene via knappene der: sjekk status, velg filer, skriv commit-melding, og trykk «push». Resultatet blir det samme som å kjøre kommandoene i terminalen.

### «Select repository … isn't associated with your workspace»

Når du prøver å trykke «Apply» på et forslag og får meldingen:

> Select repository – Choose a local repo to apply code changes to. These changes were made in HansRisnes/busbar-webapp-embedded-v2, which isn't associated with your workspace. They may not apply cleanly.

betyr det at verktøyet ikke finner en lokal klone av repoet `HansRisnes/busbar-webapp-embedded-v2` i den arbeidsflaten du har åpen. Løsningen er å åpne/klone riktig repo i miljøet før du trykker «Apply»:

1. **Kontroller hvilke repoer som er knyttet til arbeidsflaten.** I et terminalmiljø kan du kjøre `git remote -v` for å se hvilke GitHub-url-er som er tilknyttet. Mangler `HansRisnes/busbar-webapp-embedded-v2`, må du legge den til.
2. **Klone eller bytt til riktig repo.**
   * Lokalt: `git clone git@github.com:HansRisnes/busbar-webapp-embedded-v2.git`
   * GitHub Codespaces / devcontainer / lignende: åpne codespacen eller mappen som peker på dette repoet.
3. **Prøv «Apply» på nytt.** Når arbeidsflaten peker på samme repo som forslaget er laget for, får du ikke feilmeldingen og endringen kan legges inn automatisk.

Alternativt kan du kopiere diffen manuelt og lime den inn i filene i den repoen du jobber i. Da slipper du å bruke «Apply», men sørg for å lagre og committe endringene i riktig prosjekt.

### Hvordan åpne riktig repo lokalt

Hvis koden din ligger i `C:\\dev\\busbar-webapp-embedded-v2`, må IDE-en peke på akkurat denne mappen for at Git-integrasjonen
skal forstå at det er det samme som GitHub-repoet `HansRisnes/busbar-webapp-embedded-v2`.

1. **Åpne mappen i IDE-en din.**
   * I VS Code: `File` → `Open Folder…` og naviger til `C:\\dev\\busbar-webapp-embedded-v2` (du kan også åpne terminalen og
     kjøre `code C:\\dev\\busbar-webapp-embedded-v2`).
   * I JetBrains/IntelliJ-produkter: `File` → `Open…` og velg samme mappe.
2. **Bekreft at riktig Git-remote er satt.** Åpne en terminal i mappen og kjør:
   ```bash
   git remote -v
   ```
   Du skal se `https://github.com/HansRisnes/busbar-webapp-embedded-v2.git` (eller SSH-varianten) både for `fetch` og `push`.
   Hvis adressen mangler kan du legge den til med:
   ```bash
   git remote add origin https://github.com/HansRisnes/busbar-webapp-embedded-v2.git
   ```
   eller oppdatere eksisterende remote:
   ```bash
   git remote set-url origin https://github.com/HansRisnes/busbar-webapp-embedded-v2.git
   ```
3. **Last inn endringene på nytt.** Når IDE-en peker på riktig mappe med korrekt remote, skal «Apply»-funksjonen og andre Git-
   handlinger kjenne igjen repoet og fungere uten advarselen om at det ikke er tilknyttet arbeidsflaten.

### Bruke «Copy git apply» eller «Copy patch» fra nettleseren

Når du jobber i nettleserutgaven av Codex og får valgene «Opprett utkast til PR», «Copy git apply» og «Copy patch», kan du
bruke de to sistnevnte for å hente endringene ned lokalt:

1. **Copy git apply** kopierer en ferdig `git apply`-kommando til utklippstavlen. På lokal maskin kan du lime den inn i en
   terminal som står i prosjektmappen (`C:\\dev\\busbar-webapp-embedded-v2`) og kjøre den direkte. Kommandoen ser typisk ut
   som:
   ```bash
   git apply <<'PATCH'
   ...patch-innhold...
   PATCH
   ```
   Kjør kommandoen i terminalen for å legge inn endringene i de lokale filene. Deretter kan du teste, committe og pushe som
   vanlig.
2. **Copy patch** kopierer kun selve diffen. Lim den inn i en fil (for eksempel `endring.patch`) og kjør så:
   ```bash
   git apply endring.patch
   ```
   Alternativt kan du bruke `pbpaste | git apply` på macOS eller `Get-Clipboard | git apply` i PowerShell for å slippe å lage
   en midlertidig fil.
3. **Opprett utkast til PR** lager en gren på GitHub. Hvis du velger dette kan du senere hente endringene lokalt med
   ```bash
   git fetch origin <navn-på-gren>
   git checkout <navn-på-gren>
   ```
   og så teste eller justere før du merger.

Uansett metode bør du til slutt bekrefte at endringene ser riktige ut (`git status`, kjøre appen lokalt osv.), og deretter
committe og pushe dem til GitHub.
