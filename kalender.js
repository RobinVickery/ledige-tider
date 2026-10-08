'use strict';
// Ledige tider - vinduet, der åbnes fra knappen i Outlook.
// Læser ugen i kalenderen (Microsoft Graph), finder hullerne og lader
// brugeren vælge dem, der skal med i mailen som punktliste.

const K = window.LEDIGE_KONFIG || {};
const SCOPES = ['Calendars.Read'];
const DAGE = ['søndag', 'mandag', 'tirsdag', 'onsdag', 'torsdag', 'fredag', 'lørdag'];
const KORT = ['Søn', 'Man', 'Tir', 'Ons', 'Tor', 'Fre', 'Lør'];
const MDR = ['januar', 'februar', 'marts', 'april', 'maj', 'juni', 'juli', 'august', 'september', 'oktober', 'november', 'december'];
const PX_TIME = 56;

const $ = (id) => document.getElementById(id);

// Gemmes pr. enhed - aldrig noget, siden ikke kan klare sig uden
const gemt = (() => { try { return JSON.parse(localStorage.getItem('ledige-tider') || '{}'); } catch (_) { return {}; } })();
function gem() {
  try { localStorage.setItem('ledige-tider', JSON.stringify({ varighed: s.varighed, fra: s.fra, til: s.til, tilstand: s.tilstand })); } catch (_) {}
}

const s = {
  uge: mandag(new Date()),
  varighed: gemt.varighed || 30,
  fra: gemt.fra ?? 9,
  til: gemt.til ?? 16,
  tilstand: gemt.tilstand || 'til', // 'til' = vælg selv, 'fra' = alle med, fjern dem du ikke vil
  valgte: new Map(),     // nøgle (start ISO) -> { start, slut }
  fravalgte: new Set(),  // i "alle med": huller, brugeren har fjernet
  set: new Set(),        // uger, der er vist i "alle med"
  cache: new Map(),
  iOffice: false,
  demo: !K.clientId,
};

// ---------- datoer ----------
function mandag(d) { const x = new Date(d); x.setHours(0, 0, 0, 0); x.setDate(x.getDate() - ((x.getDay() + 6) % 7)); return x; }
function plusDage(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x; }
function ugenr(d) {
  const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  t.setUTCDate(t.getUTCDate() + 4 - (t.getUTCDay() || 7));
  return Math.ceil(((t - Date.UTC(t.getUTCFullYear(), 0, 1)) / 864e5 + 1) / 7);
}
const kl = (d) => d.getHours() + '.' + String(d.getMinutes()).padStart(2, '0');
const stort = (t) => t[0].toUpperCase() + t.slice(1);
const datoLang = (d) => stort(DAGE[d.getDay()]) + ' ' + d.getDate() + '. ' + MDR[d.getMonth()];
const samme = (a, b) => a.toDateString() === b.toDateString();
const op15 = (d) => { const x = new Date(d); const m = x.getMinutes(); if (m % 15 || x.getSeconds() || x.getMilliseconds()) x.setMinutes(m + 15 - (m % 15), 0, 0); return x; };
const ned15 = (d) => { const x = new Date(d); x.setMinutes(x.getMinutes() - (x.getMinutes() % 15), 0, 0); return x; };
const esc = (t) => String(t).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

// ---------- kalenderen ----------
let msalApp = null;

async function klarLogin() {
  if (s.demo) return;
  msalApp = new msal.PublicClientApplication({
    auth: {
      clientId: K.clientId,
      authority: 'https://login.microsoftonline.com/' + (K.tenant || 'organizations'),
      redirectUri: location.origin + location.pathname,
    },
    cache: { cacheLocation: 'localStorage' },
  });
  if (msalApp.initialize) await msalApp.initialize();
  const svar = await msalApp.handleRedirectPromise();
  const konto = svar?.account || msalApp.getAllAccounts()[0];
  if (konto) msalApp.setActiveAccount(konto);
}

async function hentToken() {
  const konto = msalApp.getActiveAccount();
  if (!konto) { await msalApp.loginRedirect({ scopes: SCOPES }); return new Promise(() => {}); }
  try {
    return (await msalApp.acquireTokenSilent({ scopes: SCOPES, account: konto })).accessToken;
  } catch (_) {
    await msalApp.acquireTokenRedirect({ scopes: SCOPES, account: konto });
    return new Promise(() => {});
  }
}

// Graph svarer "2026-10-09T06:30:00.0000000" i UTC, når vi beder om det
const fraUtc = (t) => new Date(t.replace(/\.\d+$/, '') + 'Z');
const lokalDato = (t) => { const [y, m, d] = t.slice(0, 10).split('-').map(Number); return new Date(y, m - 1, d); };

async function hentFraGraph(a, b) {
  const token = await hentToken();
  let url = 'https://graph.microsoft.com/v1.0/me/calendarView'
    + '?startDateTime=' + a.toISOString() + '&endDateTime=' + b.toISOString()
    + '&$select=subject,start,end,showAs,isAllDay,isCancelled&$top=100&$orderby=start/dateTime';
  const ud = [];
  for (let side = 0; url && side < 20; side++) {
    const r = await fetch(url, { headers: { Authorization: 'Bearer ' + token, Prefer: 'outlook.timezone="UTC"' } });
    if (!r.ok) throw new Error('Kalenderen svarede ' + r.status);
    const j = await r.json();
    ud.push(...j.value);
    url = j['@odata.nextLink'];
  }
  return ud.filter((e) => !e.isCancelled).map((e) => ({
    titel: e.subject || '(uden titel)',
    start: e.isAllDay ? lokalDato(e.start.dateTime) : fraUtc(e.start.dateTime),
    slut: e.isAllDay ? lokalDato(e.end.dateTime) : fraUtc(e.end.dateTime),
    status: e.showAs,
    heldag: e.isAllDay,
  }));
}

// Opdigtede møder, så siden kan prøves uden login
function demoUge(uge) {
  const m = (dag, a, b, titel, status = 'busy') => {
    const t = (hhmm) => { const d = plusDage(uge, dag); const [h, mi] = hhmm.split(':'); d.setHours(+h, +mi, 0, 0); return d; };
    return { titel, start: t(a), slut: t(b), status, heldag: false };
  };
  const skift = ugenr(uge) % 2;
  return [
    m(0, '07:15', '09:30', 'Aflevering'), m(0, '09:30', '10:30', 'Opfølgning, projekt'), m(0, '12:30', '13:30', 'Tandlæge'), m(0, '14:30', '16:00', 'Illustrationsmøde'),
    m(1, '08:30', '10:00', 'Ledelsesrum'), m(1, '10:00', '12:00', 'Værksted'),
    m(2, '08:30', '09:00', 'Check-in'), m(2, '09:30', '10:30', 'Partnermøde'), m(2, '12:30', '13:00', 'Kort møde'), m(2, '13:00', '15:00', 'Workshop-planlægning'),
    skift ? m(3, '13:00', '16:30', 'Seminar', 'tentative') : m(3, '09:00', '11:00', 'Styregruppe'),
    m(4, '08:00', skift ? '15:30' : '11:00', 'Heldagsarrangement'), m(4, '13:30', '14:00', 'Touchdown'), m(4, '14:00', '14:30', 'Stand-up'),
  ];
}

async function hentUge(uge) {
  const n = uge.toISOString();
  if (!s.cache.has(n)) s.cache.set(n, s.demo ? demoUge(uge) : await hentFraGraph(uge, plusDage(uge, 7)));
  return s.cache.get(n);
}

// ---------- hullerne ----------
const blokerer = (e) => e.status === 'busy' || e.status === 'oof' || e.status === 'tentative';

function huller(dag, ev) {
  const v0 = new Date(dag); v0.setHours(s.fra, 0, 0, 0);
  const v1 = new Date(dag); v1.setHours(s.til, 0, 0, 0);
  const nu = new Date();
  let t = nu > v0 ? op15(nu) : v0;
  const blok = ev.filter((e) => blokerer(e) && e.slut > v0 && e.start < v1).sort((a, b) => a.start - b.start);
  const ud = [];
  for (const e of blok) {
    if (e.start > t) ud.push([t, e.start]);
    if (e.slut > t) t = e.slut;
  }
  if (t < v1) ud.push([t, v1]);
  return ud
    .map(([a, b]) => [op15(a), ned15(b)])
    .filter(([a, b]) => (b - a) / 60000 >= s.varighed)
    .map(([a, b]) => ({ start: a, slut: b, n: a.toISOString() }));
}

// Overlappende møder side om side, som i Outlook
function baner(ev) {
  const ud = []; let klynge = [], baneSlut = [], klyngeSlut = 0;
  const luk = () => { klynge.forEach((x) => (x.antal = baneSlut.length)); klynge = []; baneSlut = []; };
  for (const e of ev) {
    if (klynge.length && e.start >= klyngeSlut) luk();
    let i = baneSlut.findIndex((t) => t <= e.start);
    if (i < 0) { i = baneSlut.length; baneSlut.push(0); }
    baneSlut[i] = e.slut;
    const x = { ...e, bane: i };
    klynge.push(x); ud.push(x);
    klyngeSlut = Math.max(klyngeSlut, +e.slut);
  }
  luk();
  return ud;
}

// ---------- tegning ----------
let ugensHuller = [];

async function tegn() {
  const uge = s.uge;
  const fre = plusDage(uge, 4);
  $('ugetitel').textContent = 'Uge ' + ugenr(uge) + ' · ' + uge.getDate() + '.' + (samme(uge, fre) ? '' : (uge.getMonth() !== fre.getMonth() ? ' ' + MDR[uge.getMonth()] : '')) + '–' + fre.getDate() + '. ' + MDR[fre.getMonth()] + ' ' + fre.getFullYear();
  $('kal').innerHTML = '<p class="tom" style="padding:20px">Henter kalenderen …</p>';

  let ev;
  try { ev = await hentUge(uge); } catch (e) {
    $('kal').innerHTML = '<p class="tom" style="padding:20px">Kunne ikke hente kalenderen: ' + esc(e.message) + '</p>';
    return;
  }
  if (uge !== s.uge) return; // brugeren har skiftet uge imens

  const hoejde = (s.til - s.fra) * PX_TIME;
  const y = (d) => ((d.getHours() + d.getMinutes() / 60) - s.fra) * PX_TIME;
  const idag = new Date();
  ugensHuller = [];

  let h = '<div class="uge" style="grid-template-columns:34px repeat(5,minmax(0,1fr))"><div></div>';
  for (let i = 0; i < 5; i++) {
    const d = plusDage(uge, i);
    const heldag = ev.filter((e) => e.heldag && e.start <= d && e.slut > d);
    h += '<div class="hoved' + (samme(d, idag) ? ' idag' : '') + '">' + KORT[d.getDay()] + '<b>' + d.getDate() + '. ' + MDR[d.getMonth()].slice(0, 3) + '</b>'
      + heldag.map((e) => '<div class="heldag" title="' + esc(e.titel) + '">' + esc(e.titel) + '</div>').join('') + '</div>';
  }
  h += '<div class="tid" style="height:' + hoejde + 'px">';
  for (let t = s.fra; t <= s.til; t++) h += '<span style="top:' + Math.min(Math.max((t - s.fra) * PX_TIME, 6), hoejde - 6) + 'px">' + t + '</span>';
  h += '</div>';

  for (let i = 0; i < 5; i++) {
    const d = plusDage(uge, i);
    const slutDag = new Date(d); slutDag.setHours(s.til, 0, 0, 0);
    h += '<div class="dagkol' + (slutDag < idag ? ' fortid' : '') + '" style="height:' + hoejde + 'px">';
    for (let t = s.fra + 1; t < s.til; t++) h += '<div class="linje" style="top:' + (t - s.fra) * PX_TIME + 'px"></div>';

    const v0 = new Date(d); v0.setHours(s.fra, 0, 0, 0);
    const dagens = ev.filter((e) => !e.heldag && samme(e.start, d) && e.status !== 'free').sort((a, b) => a.start - b.start);
    const foer = dagens.filter((e) => e.slut <= v0).length;
    const efter = dagens.filter((e) => e.start >= slutDag).length;
    if (foer) h += '<div class="uden-for" style="top:1px">↑ ' + foer + ' før</div>';
    for (const e of baner(dagens.filter((e) => e.slut > v0 && e.start < slutDag))) {
      const top = Math.max(y(e.start), 0), bund = Math.min(y(e.slut), hoejde);
      const bredde = 100 / e.antal;
      const kl2 = kl(e.start) + '–' + kl(e.slut);
      h += '<div class="b ' + (e.status === 'tentative' ? 'forel' : 'optaget') + '" title="' + esc(kl2 + ' ' + e.titel + (e.status === 'tentative' ? ' (foreløbig)' : '')) + '" '
        + 'style="top:' + (top + 1) + 'px;height:' + Math.max(bund - top - 2, 12) + 'px;left:calc(' + e.bane * bredde + '% + 2px);width:calc(' + bredde + '% - 4px)">'
        + esc(e.titel) + (bund - top > 34 ? '<small>' + kl2 + '</small>' : '') + '</div>';
    }
    for (const hul of huller(d, ev)) {
      ugensHuller.push(hul);
      const top = y(hul.start), bund = y(hul.slut);
      h += '<button class="b ledig" data-n="' + hul.n + '" title="' + kl(hul.start) + '–' + kl(hul.slut) + '" style="top:' + (top + 1) + 'px;height:' + (bund - top - 2) + 'px">'
        + kl(hul.start) + '–' + kl(hul.slut) + '</button>';
    }
    if (efter) h += '<div class="uden-for" style="bottom:1px">↓ ' + efter + ' efter</div>';
    if (samme(d, idag) && idag.getHours() >= s.fra && idag.getHours() < s.til) h += '<div class="nulinje" style="top:' + y(idag) + 'px"></div>';
    h += '</div>';
  }
  h += '</div>';
  $('kal').innerHTML = h;

  // "Alle med": første gang en uge vises, kommer alle dens huller med
  if (s.tilstand === 'fra' && !s.set.has(uge.toISOString())) {
    s.set.add(uge.toISOString());
    for (const hul of ugensHuller) if (!s.fravalgte.has(hul.n)) s.valgte.set(hul.n, hul);
  }
  marker();
  tegnListe();
}

function marker() {
  document.querySelectorAll('.b.ledig').forEach((b) => b.classList.toggle('valgt', s.valgte.has(b.dataset.n)));
}

function grupper() {
  const alle = [...s.valgte.values()].sort((a, b) => a.start - b.start);
  const dage = [];
  for (const t of alle) {
    const sidste = dage[dage.length - 1];
    if (sidste && samme(sidste.dag, t.start)) sidste.tider.push(t); else dage.push({ dag: t.start, tider: [t] });
  }
  return dage;
}
const og = (dele) => dele.length < 2 ? dele.join('') : dele.slice(0, -1).join(', ') + ' og ' + dele[dele.length - 1];

function tegnListe() {
  const dage = grupper();
  $('tom').hidden = dage.length > 0;
  $('tom').textContent = s.tilstand === 'fra' ? 'Ingen tider tilbage. Tryk på de grønne felter for at tage dem med igen.' : 'Tryk på de grønne felter for at tage tiderne med.';
  $('valgte').innerHTML = dage.map((g) => '<li><span class="dag" data-uge="' + mandag(g.dag).toISOString() + '" title="Vis ugen">' + datoLang(g.dag) + '</span> kl. '
    + og(g.tider.map((t) => '<span class="chip" data-n="' + t.start.toISOString() + '">' + kl(t.start) + '–' + kl(t.slut)
      + '<button data-x="' + t.start.toISOString() + '" aria-label="Fjern" title="Fjern">×</button></span>')) + '</li>').join('');
  $('indsaet').textContent = s.iOffice ? 'Indsæt i mail' : 'Kopiér listen';
}

function linjer() {
  return grupper().map((g) => datoLang(g.dag) + ' kl. ' + og(g.tider.map((t) => kl(t.start) + '–' + kl(t.slut))));
}

// ---------- handlinger ----------
function skift(n) {
  const hul = ugensHuller.find((x) => x.n === n) || s.valgte.get(n);
  if (!hul) return;
  if (s.valgte.has(n)) { s.valgte.delete(n); s.fravalgte.add(n); }
  else { s.valgte.set(n, hul); s.fravalgte.delete(n); }
  marker(); tegnListe();
}

function saetTilstand(t) {
  s.tilstand = t; gem();
  $('tilstand-til').classList.toggle('on', t === 'til');
  $('tilstand-fra').classList.toggle('on', t === 'fra');
  s.valgte.clear(); s.fravalgte.clear(); s.set.clear();
  if (t === 'fra') { s.set.add(s.uge.toISOString()); for (const hul of ugensHuller) s.valgte.set(hul.n, hul); }
  marker(); tegnListe();
}

function gaaTil(uge) { s.uge = uge; tegn(); }

function lys(n, til) {
  document.querySelectorAll('[data-n="' + n + '"]').forEach((e) => e.classList.toggle('lys', til));
}

function besked(t) { const b = $('besked'); b.textContent = t; b.hidden = !t; }

async function indsaet() {
  const l = linjer();
  if (!l.length) { besked('Vælg mindst én tid først.'); setTimeout(() => besked(''), 2500); return; }
  const tekst = l.map((x) => '• ' + x).join('\n');
  const html = '<ul>' + l.map((x) => '<li>' + esc(x) + '</li>').join('') + '</ul>';
  if (s.iOffice) {
    Office.context.ui.messageParent(JSON.stringify({ type: 'indsaet', html, tekst }));
    return;
  }
  try { await navigator.clipboard.writeText(tekst); } catch (_) {
    const ta = document.createElement('textarea'); ta.value = tekst; document.body.appendChild(ta); ta.select();
    try { document.execCommand('copy'); } catch (_) {}
    ta.remove();
  }
  $('indsaet').textContent = 'Kopieret ✓';
  setTimeout(tegnListe, 1500);
}

function fyldTimer() {
  const valg = (fra, til, v) => { let h = ''; for (let t = fra; t <= til; t++) h += '<option value="' + t + '"' + (t === v ? ' selected' : '') + '>' + t + '</option>'; return h; };
  $('fra').innerHTML = valg(6, 20, s.fra);
  $('til').innerHTML = valg(7, 22, s.til);
  $('varighed').value = String(s.varighed);
}

function genberegn() {
  // Nye huller betyder nye nøgler - "alle med" starter forfra for de viste uger
  if (s.tilstand === 'fra') { s.valgte.clear(); s.fravalgte.clear(); s.set.clear(); }
  gem(); tegn();
}

function kobl() {
  $('forrige').onclick = () => gaaTil(plusDage(s.uge, -7));
  $('naeste').onclick = () => gaaTil(plusDage(s.uge, 7));
  $('idag').onclick = () => gaaTil(mandag(new Date()));
  $('tilstand-til').onclick = () => saetTilstand('til');
  $('tilstand-fra').onclick = () => saetTilstand('fra');
  $('varighed').onchange = (e) => { s.varighed = +e.target.value; genberegn(); };
  $('fra').onchange = (e) => { s.fra = Math.min(+e.target.value, s.til - 1); fyldTimer(); genberegn(); };
  $('til').onchange = (e) => { s.til = Math.max(+e.target.value, s.fra + 1); fyldTimer(); genberegn(); };
  $('indsaet').onclick = indsaet;
  $('ryd').onclick = () => { for (const n of s.valgte.keys()) s.fravalgte.add(n); s.valgte.clear(); marker(); tegnListe(); };

  $('kal').addEventListener('click', (e) => { const b = e.target.closest('.b.ledig'); if (b) skift(b.dataset.n); });
  $('valgte').addEventListener('click', (e) => {
    const x = e.target.closest('[data-x]'); if (x) { skift(x.dataset.x); return; }
    const d = e.target.closest('[data-uge]'); if (d) gaaTil(new Date(d.dataset.uge));
  });
  // Hold musen over en tid - den lyser begge steder
  let sidst = null;
  document.addEventListener('mouseover', (e) => {
    const n = e.target.closest('[data-n]')?.dataset.n || null;
    if (n === sidst) return;
    if (sidst) lys(sidst, false);
    if (n) lys(n, true);
    sidst = n;
  });
  document.addEventListener('keydown', (e) => {
    if (e.target.tagName === 'SELECT') return;
    if (e.key === 'ArrowRight') $('naeste').click();
    if (e.key === 'ArrowLeft') $('forrige').click();
  });
}

(async function start() {
  fyldTimer();
  kobl();
  saetTilstandKnapper();
  try {
    await new Promise((ok) => (window.Office ? Office.onReady((i) => { s.iOffice = !!i.host; ok(); }) : ok()));
  } catch (_) {}
  if (s.demo) besked('Demo: møderne er opdigtede. Sæt clientId i config.js for at se din egen kalender.');
  try { await klarLogin(); } catch (e) { besked('Login fejlede: ' + e.message); }
  tegn();
})();

function saetTilstandKnapper() {
  $('tilstand-til').classList.toggle('on', s.tilstand === 'til');
  $('tilstand-fra').classList.toggle('on', s.tilstand === 'fra');
}
