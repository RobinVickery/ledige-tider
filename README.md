# Ledige tider

En knap i Outlook, når du skriver en mail: **Ledige tider**. Den åbner et vindue
med ugen, som du kender den fra Outlook. De grønne felter er dine ledige tider
inden for det tidsrum, du har valgt (fx kl. 9–16), og som er mindst så lange som
mødet (fx 30 min).

- **Vælg selv:** tryk på de grønne, du vil foreslå. De kommer i listen til højre.
- **Alle med:** alle ledige tider er med fra start. Fjern dem, du ikke vil sende,
  med krydset eller ved at trykke på feltet.
- Pilene (eller ← →) skifter uge. Valgte tider huskes på tværs af ugerne.
- **Indsæt i mail** lægger listen ind i mailen, dér hvor markøren står:

  - Mandag 19. oktober kl. 10.30–12.00 og 13.30–14.30
  - Tirsdag 20. oktober kl. 12.00–16.00

Foreløbige aftaler tæller som optaget, men vises stiplede. Aflyste møder og
"ledig"/"arbejder et andet sted" tæller som ledigt.

## Filerne

| Fil | Hvad |
|---|---|
| `kalender.html/.js/.css` | Vinduet med ugen. Uden `clientId` kører det i demo med opdigtede møder. |
| `commands.html/.js` | Knappen: åbner vinduet og sætter listen ind i mailen. |
| `config.js` | `clientId` fra Azure. |
| `manifest.skabelon.xml` | Beskriver knappen for Outlook. `tools/lav-manifest.mjs` laver `manifest.xml`. |
| `vendor/msal-browser.min.js` | Microsofts login-bibliotek (MSAL 2.38.3). |
| `tools/server.mjs` | Lokal visning: `node tools/server.mjs` → http://localhost:8091/kalender.html |
| `tools/lav-ikoner.mjs` | Tegner ikonerne i `ikoner/`. |

## Sådan kommer den i Outlook (én gang)

1. **Læg filerne på en https-adresse.** Fx GitHub Pages:
   `https://<bruger>.github.io/ledige-tider/`. Alle filer ligger åbent, men der er
   ingen hemmeligheder i dem (clientId er ikke hemmeligt).

2. **Registrér appen i Azure** (portal.azure.com → *App registrations* → *New registration*):
   - Navn: `Ledige tider`
   - Kontotyper: *Accounts in this organizational directory only*
   - Redirect URI: platform **Single-page application (SPA)**, adresse
     `https://<adressen>/kalender.html`
   - Under *API permissions*: `Microsoft Graph → Delegated → Calendars.Read`
     (`User.Read` står der i forvejen). Tryk *Grant admin consent*, hvis du er administrator.
   - Kopiér *Application (client) ID* ind i `config.js` som `clientId`, og
     *Directory (tenant) ID* som `tenant`.

3. **Lav manifestet:** `node tools/lav-manifest.mjs https://<adressen>`

4. **Installér det i Outlook:** Åbn https://aka.ms/olksideload i Outlook på nettet →
   *Mine tilføjelsesprogrammer* → *Tilføj et brugerdefineret tilføjelsesprogram* →
   *Tilføj fra fil* → vælg `manifest.xml`. Knappen kommer så frem i den nye Outlook,
   Outlook på nettet og den klassiske Outlook.

   Til hele firmaet: Microsoft 365 admin → *Integrerede apps* → *Upload brugerdefinerede apps*.

Første gang du trykker på knappen, beder Microsoft dig logge ind og godkende, at
appen må læse din kalender. Derefter husker den det.
