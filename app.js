/* Abrechnung – Reisezeit, Arbeitszeit und Belege erfassen.
   Laeuft offline, Daten bleiben im Geraet (IndexedDB). Export als JSON,
   das ein Skript am Rechner in die Excel-Vorlagen einsetzt.              */

'use strict';

// ------------------------------------------------------------- Konstanten --
// Spiegelt das Blatt "Grunddaten" der Reiserechnungs-Vorlage.
const SATZ = {
  gebuehrTag: 30,
  gebuehr12: 2.5,
  essenSelbstbehalt: 17,
  gebuehrNacht: 17,
  mindestStunden: 3,
  kmGeldAktiv: false,
  kmGeld: 0.5,
};

const ZWECKE = ['Kundenbesuche', 'Kundenbesuch', 'Homeoffice', 'Office',
  'Messe', 'Workshop', 'Vertriebstagung', 'Anreise', 'Rückfahrt', 'Exkursion'];

const TAGTYPEN = [
  { id: 'reise', label: 'Außendienst' },
  { id: 'homeoffice', label: 'Homeoffice' },
  { id: 'urlaub', label: 'Urlaub' },
  { id: 'feiertag', label: 'Feiertag' },
];

const WOCHENTAG = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];

// ------------------------------------------------------------------ Hilfen --
const $ = (s) => document.querySelector(s);
const $$ = (s) => Array.from(document.querySelectorAll(s));

function iso(d) {
  const z = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${z(d.getMonth() + 1)}-${z(d.getDate())}`;
}
function heuteIso() { return iso(new Date()); }
function monatVon(isoTag) { return isoTag.slice(0, 7); }

/** Minuten auf die naechste Viertelstunde runden. */
function viertel(datum) {
  const m = datum.getHours() * 60 + datum.getMinutes();
  const g = Math.round(m / 15) * 15;
  return hhmm(Math.min(g, 24 * 60));
}
function hhmm(minuten) {
  const h = Math.floor(minuten / 60), m = minuten % 60;
  return `${h}:${String(m).padStart(2, '0')}`;
}
function zuMinuten(s) {
  if (!s) return null;
  const [h, m] = s.split(':').map(Number);
  return h * 60 + m;
}
function stdMin(stunden) {
  const m = Math.round(stunden * 60);
  return `${Math.floor(m / 60)}:${String(m % 60).padStart(2, '0')}`;
}
function eur(n) {
  return n.toLocaleString('de-AT', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

/* Uhrzeit aus einem Textfeld pruefen und aufs Viertelstunden-Raster bringen.
   Erlaubt sind 0:00 bis 24:00 sowie 0:01, der Marker fuer den Folgetag einer
   Mehrtagesreise. Rueckgabe null heisst leer, false heisst ungueltig.      */
function zeitPruefen(roh) {
  const s = String(roh || '').trim().replace('.', ':');
  if (!s) return null;
  const m = s.match(/^(\d{1,2}):?(\d{2})$/);
  if (!m) return false;
  let h = Number(m[1]), mi = Number(m[2]);
  if (h > 24 || mi > 59 || (h === 24 && mi > 0)) return false;
  if (h === 0 && mi === 1) return '0:01';
  mi = Math.round(mi / 15) * 15;
  if (mi === 60) { h += 1; mi = 0; }
  if (h > 24) return false;
  return `${h}:${String(mi).padStart(2, '0')}`;
}

/** Taggeld-Zwoelftel eines Tages, exakt wie Spalte T der Vorlage rechnet. */
function zwoelftel(abfahrt, ankunft) {
  if (!abfahrt || !ankunft) return 0;
  const [ch, cm] = abfahrt.split(':').map(Number);
  const [eh, em] = ankunft.split(':').map(Number);
  const W = eh - ch;            // volle Stundendifferenz
  const X = em - cm;            // Minutendifferenz
  const Z = X > 0 ? 1 : 0;      // angefangene Stunde zaehlt voll
  if (W > 11) return 12;
  if (W + Z <= SATZ.mindestStunden) return 0;
  return W + Z;
}

// ------------------------------------------------------------- Datenbank ---
const DB = (() => {
  let db = null;
  function offen() {
    if (db) return Promise.resolve(db);
    return new Promise((ok, fehler) => {
      const r = indexedDB.open('abrechnung', 1);
      r.onupgradeneeded = () => {
        const d = r.result;
        if (!d.objectStoreNames.contains('tage')) d.createObjectStore('tage', { keyPath: 'datum' });
        if (!d.objectStoreNames.contains('belege')) d.createObjectStore('belege', { keyPath: 'id' });
        if (!d.objectStoreNames.contains('einst')) d.createObjectStore('einst', { keyPath: 'k' });
      };
      r.onsuccess = () => { db = r.result; ok(db); };
      r.onerror = () => fehler(r.error);
    });
  }
  function tx(store, modus, fn) {
    return offen().then((d) => new Promise((ok, fehler) => {
      const t = d.transaction(store, modus);
      const req = fn(t.objectStore(store));
      t.oncomplete = () => ok(req && 'result' in req ? req.result : undefined);
      t.onerror = () => fehler(t.error);
    }));
  }
  return {
    hole: (s, k) => tx(s, 'readonly', (o) => o.get(k)),
    alle: (s) => tx(s, 'readonly', (o) => o.getAll()),
    setz: (s, v) => tx(s, 'readwrite', (o) => o.put(v)),
    weg: (s, k) => tx(s, 'readwrite', (o) => o.delete(k)),
  };
})();

// ----------------------------------------------------------------- Status ---
const S = {
  tag: null,          // Tagesobjekt von heute
  monat: heuteIso().slice(0, 7),
  tage: [],           // alle Tage des gewaehlten Monats
  belege: [],         // Belege des gewaehlten Monats
  einst: { kmStart: {}, kmEnde: {}, orte: [] },
};

function leererTag(datum) {
  return {
    datum, typ: 'reise',
    abfahrt: null, ankunft: null,
    segmente: [], laeuft: null,
    orte: [], zweck: null, text: '',
    arbeitsessen: 0, naechtigung: false,
    grenze: null, abgeschlossen: false,
  };
}

function arbeitStunden(t) {
  let ms = (t.segmente || []).reduce((a, s) => a + (s.ende - s.start), 0);
  if (t.laeuft) ms += Date.now() - t.laeuft;
  return Math.round((ms / 3600000) * 4) / 4;   // Viertelstunden
}

async function tagSpeichern() {
  S.tag.arbeitszeit = arbeitStunden(S.tag);
  await DB.setz('tage', S.tag);
  const i = S.tage.findIndex((x) => x.datum === S.tag.datum);
  if (i >= 0) S.tage[i] = S.tag; else S.tage.push(S.tag);
  zeichneHeute(); zeichneMonat();
}

// ------------------------------------------------- Mehrtagesreise-Uebergang --
/* Wenn an einem Tag "Haus verlassen" gedrueckt wurde, aber kein "Heim
   gekommen" folgte, schliesst die Vorlage den Tag mit 24:00 ab und der
   Folgetag beginnt mit 0:01. Das ist die bisherige Konvention von Hand. */
async function uebergangPruefen() {
  const alle = await DB.alle('tage');
  const h = heuteIso();
  for (const t of alle) {
    if (t.datum >= h) continue;
    if (t.abfahrt && !t.ankunft) {
      t.ankunft = '24:00';
      t.naechtigung = true;
      if (t.laeuft) { t.segmente.push({ start: t.laeuft, ende: t.laeuft }); t.laeuft = null; }
      await DB.setz('tage', t);

      const naechster = iso(new Date(new Date(t.datum + 'T12:00:00').getTime() + 864e5));
      let n = await DB.hole('tage', naechster);
      if (!n) n = leererTag(naechster);
      if (!n.abfahrt) { n.abfahrt = '0:01'; await DB.setz('tage', n); }
    }
  }
}

// ------------------------------------------------------------------- Laden --
async function laden() {
  const e = await DB.hole('einst', 'haupt');
  if (e) S.einst = Object.assign(S.einst, e.wert);
  orteNormalisieren();
  await uebergangPruefen();
  S.tag = (await DB.hole('tage', heuteIso())) || leererTag(heuteIso());
  await monatLaden(S.monat);
  zeichneAlles();
}

async function monatLaden(m) {
  S.monat = m;
  const alle = await DB.alle('tage');
  S.tage = alle.filter((t) => monatVon(t.datum) === m).sort((a, b) => a.datum.localeCompare(b.datum));
  const b = await DB.alle('belege');
  S.belege = b.filter((x) => monatVon(x.datum) === m).sort((a, b2) => a.id - b2.id);
}

async function einstSpeichern() {
  await DB.setz('einst', { k: 'haupt', wert: S.einst });
}

// ------------------------------------------------------------------ Toast --
let toastTimer;
function toast(txt) {
  const t = $('#toast');
  t.textContent = txt; t.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { t.hidden = true; }, 2600);
}

/** Rueckfrage vor dem Loeschen. */
function bestaetigt(frage) {
  return window.confirm(frage);
}

// ----------------------------------------------------------------- Dialog --
function dialog({ titel, hint, inhalt, ok = 'Speichern', abbrechen = 'Abbrechen',
                  extra, beimOeffnen, beiOk }) {
  const hg = $('#dialog-hg'), d = $('#dialog');
  d.innerHTML = `<h3>${titel}</h3>${hint ? `<p class="dlg-hint">${hint}</p>` : ''}
    <div id="dlg-inhalt">${inhalt}</div>
    <div class="dlg-knoepfe">
      <button id="dlg-ab" class="flach">${abbrechen}</button>
      <button id="dlg-ok" class="primaer">${ok}</button>
    </div>
    ${extra ? `<button id="dlg-extra" class="flach gefahr">${extra.label}</button>` : ''}`;
  hg.hidden = false; d.hidden = false;
  const zu = () => { hg.hidden = true; d.hidden = true; };
  $('#dlg-ab').onclick = zu;
  hg.onclick = zu;
  $('#dlg-ok').onclick = async () => {
    const r = beiOk ? await beiOk(d) : true;
    if (r !== false) zu();
  };
  if (extra) {
    $('#dlg-extra').onclick = async () => {
      const r = await extra.tun(d);
      if (r !== false) zu();
    };
  }
  if (beimOeffnen) beimOeffnen(d);
}

/** Chip-Gruppe: einfach- oder mehrfachauswahl. */
function chips(id, werte, aktive = [], mehrfach = false) {
  const h = werte.map((w) => {
    const v = typeof w === 'string' ? w : w.id;
    const l = typeof w === 'string' ? w : w.label;
    return `<button type="button" data-v="${encodeURIComponent(v)}" class="${aktive.includes(v) ? 'an' : ''}">${l}</button>`;
  }).join('');
  return `<div class="chips" id="${id}" data-mehrfach="${mehrfach ? 1 : 0}">${h}</div>`;
}
function chipsBinden(wurzel) {
  wurzel.querySelectorAll('.chips').forEach((g) => {
    g.onclick = (ev) => {
      const b = ev.target.closest('button[data-v]');
      if (!b) return;
      if (g.dataset.mehrfach === '1') b.classList.toggle('an');
      else { g.querySelectorAll('button').forEach((x) => x.classList.remove('an')); b.classList.add('an'); }
      g.dispatchEvent(new CustomEvent('wahl', { bubbles: true }));
    };
  });
}
function chipsWert(id) {
  const b = $(`#${id} button.an`);
  return b ? decodeURIComponent(b.dataset.v) : null;
}
function chipsWerte(id) {
  return $$(`#${id} button.an`).map((b) => decodeURIComponent(b.dataset.v));
}

// ------------------------------------------------------------------ Orte --
/* Bewusst ohne GPS und ohne Geocoding-Dienst: die App soll keine Koordinaten
   nach aussen geben. Stattdessen lernt sie die Stammorte mit und bietet die
   haeufigsten als Ein-Tipp-Auswahl an.                                     */

/** Ortsliste aus alten Versionen (reine Strings) mitnehmen. */
function orteNormalisieren() {
  S.einst.orte = (S.einst.orte || []).map(
    (o) => (typeof o === 'string' ? { name: o, anzahl: 1 } : o)
  ).filter((o) => o && o.name);
}

function orteHaeufig(n = 8) {
  return S.einst.orte
    .slice()
    .sort((a, b) => (b.anzahl || 0) - (a.anzahl || 0) || a.name.localeCompare(b.name))
    .slice(0, n);
}

async function ortMerken(name) {
  const t = S.einst.orte.find((o) => o.name.toLowerCase() === name.toLowerCase());
  if (t) { t.anzahl = (t.anzahl || 0) + 1; }
  else { S.einst.orte.push({ name, anzahl: 1 }); }
  await einstSpeichern();
}

async function ortHinzufuegen(name) {
  const n = String(name || '').trim();
  if (!n) return;
  S.tag.orte.push({ name: n, zeit: viertel(new Date()) });
  await ortMerken(n);
  await tagSpeichern();
  toast(`${n} erfasst`);
}

function ortDialog() {
  const bekannt = S.einst.orte.slice().sort((a, b) => a.name.localeCompare(b.name));
  dialog({
    titel: 'Ort hinzufügen',
    hint: 'Tippen reicht beim ersten Mal. Danach steht der Ort in der Schnellwahl.',
    inhalt: `<label class="feld">Ortsname</label>
      <input type="text" id="ort-name" list="ort-liste" placeholder="z.B. Pöchlarn"
             autocomplete="off" autocapitalize="words" enterkeyhint="done">
      <datalist id="ort-liste">${bekannt.map((o) => `<option value="${o.name}">`).join('')}</datalist>`,
    beimOeffnen: (d) => {
      setTimeout(() => $('#ort-name').focus(), 60);
      $('#ort-name').addEventListener('keydown', (e) => {
        if (e.key === 'Enter') { e.preventDefault(); $('#dlg-ok').click(); }
      });
    },
    beiOk: async () => {
      const n = $('#ort-name').value.trim();
      if (!n) return false;
      await ortHinzufuegen(n);
    },
  });
}

// ------------------------------------------------------------ Grenzdialog --
function grenzDialog() {
  const g = S.tag.grenze || {};
  dialog({
    titel: 'Grenzübertritt',
    hint: 'Zeitpunkt und Land, so wie es die Reiserechnung in den Spalten AK bis AO erwartet.',
    inhalt: `
      <label class="feld">Hinweg, Uhrzeit</label>
      <input type="time" id="gr-hin" step="900" value="${g.hin ? g.hin.padStart(5, '0') : ''}">
      <label class="feld">Retour, Uhrzeit</label>
      <input type="time" id="gr-ret" step="900" value="${g.retour ? g.retour.padStart(5, '0') : ''}">
      <label class="feld">Land</label>
      ${chips('gr-land', ['D', 'A', 'CH', 'I', 'HU', 'SK', 'CZ', 'SI', 'NL'], g.land ? [g.land] : [])}`,
    beimOeffnen: (d) => chipsBinden(d),
    beiOk: async () => {
      const hin = zeitPruefen($('#gr-hin').value);
      const ret = zeitPruefen($('#gr-ret').value);
      if (hin === false || ret === false) { toast('Uhrzeit bitte als hh:mm'); return false; }
      const land = chipsWert('gr-land');
      S.tag.grenze = (hin || ret || land) ? { hin, retour: ret, land } : null;
      await tagSpeichern();
    },
  });
}

// --------------------------------------------------------- Tagesabschluss --
function abschlussDialog() {
  const t = S.tag;
  const orte = t.orte.map((o) => o.name);
  const einzig = orte.filter((o, i) => orte.indexOf(o) === i);

  dialog({
    titel: 'Tag abschließen',
    hint: 'Das Ergebnis landet in Spalte J der Reiserechnung.',
    ok: 'Abschließen',
    inhalt: `
      <label class="feld">Art des Tages</label>
      ${chips('ab-typ', TAGTYPEN, [t.typ])}
      <label class="feld">Orte, Reihenfolge wie erfasst</label>
      ${einzig.length ? chips('ab-orte', einzig, einzig, true)
        : '<p class="dlg-hint" style="margin:0">Keine Orte erfasst. Trag sie unten direkt in den Text ein.</p>'}
      <label class="feld">Zweck <span class="hint">mehrere möglich</span></label>
      ${chips('ab-zweck', ZWECKE, t.zwecke || (t.zweck ? [t.zweck] : []), true)}
      <label class="feld">Text für die Excel</label>
      <textarea id="ab-text">${t.text || ''}</textarea>`,
    beimOeffnen: (d) => {
      chipsBinden(d);
      const bauen = () => {
        const o = chipsWerte('ab-orte');
        const z = chipsWerte('ab-zweck');
        const typ = chipsWert('ab-typ');
        let s;
        if (typ === 'urlaub') s = 'Urlaub';
        else if (typ === 'feiertag') s = 'Feiertag';
        else if (!o.length) s = z.length ? z.join(', ') : (typ === 'homeoffice' ? 'Homeoffice' : '');
        else s = o.join(', ') + (z.length ? ` (${z.join(', ')})` : '');
        $('#ab-text').value = s;
      };
      if (!t.text) bauen();
      d.addEventListener('wahl', bauen);
    },
    beiOk: async () => {
      t.typ = chipsWert('ab-typ') || 'reise';
      t.zwecke = chipsWerte('ab-zweck');
      delete t.zweck;
      t.text = $('#ab-text').value.trim();
      if (t.laeuft) { t.segmente.push({ start: t.laeuft, ende: Date.now() }); t.laeuft = null; }
      t.abgeschlossen = true;
      await tagSpeichern();
      toast('Tag abgeschlossen');
    },
  });
}

// ----------------------------------------------------------- Tag bearbeiten --
function tagBearbeiten(datum) {
  const t = S.tage.find((x) => x.datum === datum) || leererTag(datum);
  const d = new Date(datum + 'T12:00:00');
  dialog({
    titel: `${WOCHENTAG[d.getDay()]}, ${d.getDate()}.${d.getMonth() + 1}.`,
    inhalt: `
      <label class="feld">Haus verlassen, hh:mm</label>
      <input type="text" id="be-ab" inputmode="numeric" placeholder="7:00" value="${t.abfahrt || ''}">
      <label class="feld">Heim gekommen, hh:mm <span class="hint">24:00 heißt auswärts übernachtet</span></label>
      <input type="text" id="be-an" inputmode="numeric" placeholder="18:00" value="${t.ankunft || ''}">
      <label class="feld">Arbeitszeit in Stunden</label>
      <input type="number" id="be-az" step="0.25" inputmode="decimal" value="${t.arbeitszeit ?? ''}">
      <label class="feld">Text für Spalte J</label>
      <textarea id="be-text">${t.text || ''}</textarea>
      <label class="feld">Arbeitsessen</label>
      <input type="number" id="be-essen" step="1" min="0" inputmode="numeric" value="${t.arbeitsessen || 0}">
      <label class="feld">Nächtigung auswärts</label>
      ${chips('be-nacht', [{ id: 'ja', label: 'Ja' }, { id: 'nein', label: 'Nein' }], [t.naechtigung ? 'ja' : 'nein'])}`,
    beimOeffnen: (el) => chipsBinden(el),
    extra: {
      label: 'Diesen Tag komplett leeren',
      tun: async () => {
        if (!bestaetigt(`${d.getDate()}.${d.getMonth() + 1}. wirklich leeren?`)) return false;
        await DB.weg('tage', datum);
        if (datum === heuteIso()) S.tag = leererTag(datum);
        await monatLaden(S.monat);
        zeichneAlles();
        toast('Tag geleert');
      },
    },
    beiOk: async () => {
      const ab = zeitPruefen($('#be-ab').value);
      const an = zeitPruefen($('#be-an').value);
      if (ab === false || an === false) { toast('Uhrzeit bitte als hh:mm, z.B. 7:15'); return false; }
      if ((ab === null) !== (an === null)) {
        toast('Abfahrt und Heimkehr müssen beide gesetzt sein oder beide leer');
        return false;
      }
      t.abfahrt = ab; t.ankunft = an;
      const az = $('#be-az').value;
      t.arbeitszeit = az === '' ? null : Number(az);
      t.segmente = t.segmente || []; t.laeuft = null;
      t.text = $('#be-text').value.trim();
      t.arbeitsessen = Number($('#be-essen').value || 0);
      t.naechtigung = chipsWert('be-nacht') === 'ja';
      t.handbearbeitet = true;
      await DB.setz('tage', t);
      if (t.datum === heuteIso()) S.tag = t;
      await monatLaden(S.monat);
      zeichneAlles();
      toast('Tag gespeichert');
    },
  });
}

// ------------------------------------------------------------------ Belege --
/** Foto verkleinern, damit der Monatsexport nicht explodiert. */
function fotoVerkleinern(datei, maxKante = 1600, guete = 0.75) {
  return new Promise((ok, fehler) => {
    const leser = new FileReader();
    leser.onload = () => {
      const bild = new Image();
      bild.onload = () => {
        const f = Math.min(1, maxKante / Math.max(bild.width, bild.height));
        const c = document.createElement('canvas');
        c.width = Math.round(bild.width * f);
        c.height = Math.round(bild.height * f);
        c.getContext('2d').drawImage(bild, 0, 0, c.width, c.height);
        ok(c.toDataURL('image/jpeg', guete));
      };
      bild.onerror = () => fehler(new Error('Bild nicht lesbar'));
      bild.src = leser.result;
    };
    leser.onerror = () => fehler(leser.error);
    leser.readAsDataURL(datei);
  });
}

let fotoPuffer = null;

async function belegSpeichern() {
  const text = $('#beleg-text').value.trim();
  const betrag = $('#beleg-betrag').value;
  if (!text && !betrag) { toast('Beschreibung oder Betrag fehlt'); return; }
  const b = {
    id: Date.now(),
    datum: heuteIso(),
    text, betrag: betrag === '' ? null : Number(betrag),
    foto: fotoPuffer, fotoTyp: fotoPuffer ? 'image/jpeg' : null,
  };
  await DB.setz('belege', b);
  if (monatVon(b.datum) === S.monat) S.belege.push(b);
  fotoPuffer = null;
  $('#beleg-text').value = ''; $('#beleg-betrag').value = '';
  $('#foto-vorschau').hidden = true; $('#foto-text').hidden = false;
  $('#foto-text').textContent = 'Beleg fotografieren';
  zeichneBelege();
  toast('Beleg gespeichert');
}

async function belegLoeschen(id) {
  const b = S.belege.find((x) => x.id === id);
  if (!bestaetigt(`Beleg "${b?.text || 'ohne Beschreibung'}" löschen?`)) return;
  await DB.weg('belege', id);
  S.belege = S.belege.filter((x) => x.id !== id);
  zeichneBelege();
  toast('Beleg gelöscht');
}

// ------------------------------------------------------------------ Export --
function monatsDaten() {
  const tage = S.tage
    .filter((t) => t.abfahrt || t.ankunft || t.text || t.arbeitszeit || t.arbeitsessen)
    .map((t) => ({
      datum: t.datum,
      abfahrt: t.abfahrt, ankunft: t.ankunft,
      arbeitszeit: t.arbeitszeit ?? arbeitStunden(t) ?? null,
      text: t.text || '',
      arbeitsessen: t.arbeitsessen || 0,
      naechtigung: !!t.naechtigung,
      grenze: t.grenze || null,
    }));
  return {
    version: 1,
    erstellt: new Date().toISOString(),
    monat: S.monat,
    kmStart: S.einst.kmStart[S.monat] ?? null,
    kmEnde: S.einst.kmEnde[S.monat] ?? null,
    tage,
    belege: S.belege.map((b) => ({
      datum: b.datum, text: b.text, betrag: b.betrag,
      foto: b.foto || null, fotoTyp: b.fotoTyp || null,
    })),
  };
}

async function dateiRaus(name, text) {
  const blob = new Blob([text], { type: 'application/json' });
  const datei = new File([blob], name, { type: 'application/json' });

  // iOS: Teilen-Menue ist der zuverlaessige Weg zum PC (Mail, iCloud, Dateien).
  if (navigator.canShare && navigator.canShare({ files: [datei] })) {
    try { await navigator.share({ files: [datei], title: name }); return; }
    catch (e) { if (e.name === 'AbortError') return; }
  }
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
  toast('Datei liegt in den Downloads');
}

async function exportieren() {
  const d = monatsDaten();
  if (!d.tage.length) { toast('Für diesen Monat ist nichts erfasst'); return; }
  if (d.kmStart == null) { toast('KM-Stand am 1. fehlt'); return; }
  await dateiRaus(`abrechnung-${S.monat}.json`, JSON.stringify(d, null, 2));
}

async function backup() {
  const tage = await DB.alle('tage');
  const belege = await DB.alle('belege');
  const einst = await DB.hole('einst', 'haupt');
  const d = { version: 1, typ: 'backup', erstellt: new Date().toISOString(), tage, belege, einst: einst?.wert };
  await dateiRaus(`abrechnung-backup-${heuteIso()}.json`, JSON.stringify(d));
}

// ------------------------------------------------------------ Darstellung --
function zeichneAlles() { zeichneHeute(); zeichneMonat(); zeichneBelege(); }

function zeichneHeute() {
  const t = S.tag;
  const d = new Date(t.datum + 'T12:00:00');
  $('#kopf-datum').textContent =
    `${WOCHENTAG[d.getDay()]}, ${d.getDate()}. ${d.toLocaleString('de-AT', { month: 'long' })}`;

  const unterwegs = !!t.abfahrt && !t.ankunft;
  const p = $('#kopf-status');
  p.className = 'pill' + (unterwegs ? ' unterwegs' : (t.laeuft ? ' arbeit' : ''));
  p.textContent = unterwegs ? 'unterwegs' : (t.laeuft ? 'Arbeit läuft' : 'zuhause');

  $('#an-abfahrt').textContent = t.abfahrt || '–';
  $('#an-ankunft').textContent = t.ankunft || '–';
  $('#btn-los').hidden = !!t.abfahrt;
  $('#btn-heim').hidden = !t.abfahrt || !!t.ankunft;

  const z = zwoelftel(t.abfahrt, t.ankunft);
  $('#reise-fuss').textContent = t.abfahrt && t.ankunft
    ? `${z}/12 Taggeld, also ${eur(z * SATZ.gebuehr12)} Euro brutto.`
    : (t.abfahrt ? 'Läuft. Taggeld steht beim Heimkommen fest.'
      : 'Solange du nicht wegfährst, gibt es kein Taggeld.');

  $('#an-arbeit').textContent = stdMin(arbeitStunden(t));
  $('#an-arbeit-seit').textContent = t.laeuft
    ? new Date(t.laeuft).toLocaleTimeString('de-AT', { hour: '2-digit', minute: '2-digit' }) : '–';
  $('#btn-arbeit').textContent = t.laeuft ? 'Arbeit beenden' : 'Arbeit starten';
  $('#btn-arbeit').className = 'gross' + (t.laeuft ? ' erfolg' : '');

  $('#orte-zahl').textContent = t.orte.length ? `${t.orte.length} erfasst` : '';
  $('#orte-liste').innerHTML = t.orte.map((o, i) => `
    <li><div class="haupt"><b>${o.name}</b><small>${o.zeit}</small></div>
    <button class="weg" data-ort="${i}">×</button></li>`).join('');

  const heute = t.orte.map((o) => o.name.toLowerCase());
  $('#orte-schnell').innerHTML = orteHaeufig()
    .filter((o) => !heute.includes(o.name.toLowerCase()))
    .map((o) => `<button type="button" data-schnell="${encodeURIComponent(o.name)}">${o.name}</button>`)
    .join('');

  $('#an-essen').textContent = t.arbeitsessen || 0;
  $('#chk-nacht').checked = !!t.naechtigung;
  const g = t.grenze;
  $('#grenze-fuss').textContent = g
    ? `${g.land || '?'}${g.hin ? ', hin ' + g.hin : ''}${g.retour ? ', retour ' + g.retour : ''}`
    : '';
  $('#abschluss-fuss').textContent = t.abgeschlossen && t.text ? `Eingetragen: ${t.text}` : '';
}

function rechneMonat() {
  const zw = S.tage.reduce((a, t) => a + zwoelftel(t.abfahrt, t.ankunft), 0);
  const essen = S.tage.reduce((a, t) => a + (t.arbeitsessen || 0), 0);
  const naechte = S.tage.filter((t) => t.naechtigung).length;
  const arbeit = S.tage.reduce((a, t) => a + (t.arbeitszeit ?? arbeitStunden(t) ?? 0), 0);
  const tgBrutto = zw * SATZ.gebuehr12;
  const kuerzung = essen * SATZ.essenSelbstbehalt;
  const nacht = naechte * SATZ.gebuehrNacht;
  const kmS = S.einst.kmStart[S.monat], kmE = S.einst.kmEnde[S.monat];
  const km = (kmS != null && kmE != null) ? kmE - kmS : null;
  return {
    zw, essen, naechte, arbeit, tgBrutto, kuerzung, nacht, km,
    summe: tgBrutto - kuerzung + nacht,
  };
}

function zeichneMonat() {
  const [j, m] = S.monat.split('-').map(Number);
  const d = new Date(j, m - 1, 1);
  $('#monat-titel').textContent = d.toLocaleString('de-AT', { month: 'long', year: 'numeric' });
  $('#km-start').value = S.einst.kmStart[S.monat] ?? '';
  $('#km-ende').value = S.einst.kmEnde[S.monat] ?? '';

  const r = rechneMonat();
  $('#km-summe').textContent = r.km != null ? `${r.km.toLocaleString('de-AT')} km im Monat` : '';

  const zeilen = [
    ['Taggeld-Zwölftel', `${r.zw}/12`],
    ['Taggeld brutto', `${eur(r.tgBrutto)} €`],
    ['Abzug Arbeitsessen', r.essen ? `− ${eur(r.kuerzung)} €` : '0,00 €'],
    ['Nächtigungen', `${r.naechte} × ${SATZ.gebuehrNacht} = ${eur(r.nacht)} €`],
    ['Kilometer', r.km != null ? `${r.km.toLocaleString('de-AT')} km` : '–'],
    ['Arbeitszeit gesamt', stdMin(r.arbeit)],
  ];
  $('#vorschau-body').innerHTML =
    zeilen.map(([a, b]) => `<tr><td>${a}</td><td>${b}</td></tr>`).join('')
    + `<tr class="summe"><td>Summe brutto</td><td>${eur(r.summe)} €</td></tr>`;

  const anzahl = new Date(j, m, 0).getDate();
  let h = '';
  for (let tag = 1; tag <= anzahl; tag++) {
    const dat = `${S.monat}-${String(tag).padStart(2, '0')}`;
    const t = S.tage.find((x) => x.datum === dat);
    const wd = new Date(j, m - 1, tag).getDay();
    const we = wd === 0 || wd === 6;
    const voll = t && t.text && (t.arbeitszeit || t.abfahrt);
    const teil = t && (t.text || t.abfahrt || t.arbeitszeit);
    h += `<li class="${we ? 'wochenende' : ''}" data-tag="${dat}">
      <span class="tag-nr">${WOCHENTAG[wd]} ${tag}.</span>
      <span class="punkt ${voll ? 'voll' : (teil ? 'teil' : '')}"></span>
      <div class="haupt"><b>${t?.text || '<span style="color:var(--grau);font-weight:400">leer</span>'}</b>
      <small>${t?.abfahrt ? `${t.abfahrt}–${t.ankunft || '…'}` : ''}${t?.arbeitszeit ? ` · ${t.arbeitszeit} h` : ''}</small></div>
    </li>`;
  }
  $('#monat-liste').innerHTML = h;
}

function zeichneBelege() {
  const s = S.belege.reduce((a, b) => a + (b.betrag || 0), 0);
  $('#beleg-summe').textContent = S.belege.length ? `${S.belege.length} · ${eur(s)} €` : '';
  $('#beleg-liste').innerHTML = S.belege.map((b) => `
    <li>
      ${b.foto ? `<img class="beleg-bild" src="${b.foto}" alt="">` : '<span class="beleg-bild"></span>'}
      <div class="haupt"><b>${b.text || 'ohne Beschreibung'}</b><small>${b.datum}</small></div>
      <span class="betrag">${b.betrag != null ? eur(b.betrag) + ' €' : ''}</span>
      <button class="weg" data-beleg="${b.id}">×</button>
    </li>`).join('');
  if (S.belege.length > 15) {
    toast('Mehr als 15 Belege. Das Skript legt automatisch ein zweites Blatt an.');
  }
}

// ----------------------------------------------------------- Verdrahtung ---
function verdrahten() {
  $$('#tabs button').forEach((b) => {
    b.onclick = () => {
      $$('#tabs button').forEach((x) => x.classList.remove('aktiv'));
      $$('.tab').forEach((x) => x.classList.remove('aktiv'));
      b.classList.add('aktiv');
      $(`#tab-${b.dataset.tab}`).classList.add('aktiv');
      window.scrollTo(0, 0);
    };
  });

  $('#btn-los').onclick = async () => {
    S.tag.abfahrt = viertel(new Date());
    if (!S.tag.laeuft && !S.tag.segmente.length) S.tag.laeuft = Date.now();
    await tagSpeichern();
    toast(`Abfahrt ${S.tag.abfahrt} erfasst`);
  };
  $('#btn-heim').onclick = async () => {
    S.tag.ankunft = viertel(new Date());
    if (S.tag.laeuft) { S.tag.segmente.push({ start: S.tag.laeuft, ende: Date.now() }); S.tag.laeuft = null; }
    await tagSpeichern();
    const z = zwoelftel(S.tag.abfahrt, S.tag.ankunft);
    toast(`Zuhause um ${S.tag.ankunft}. ${z}/12 Taggeld.`);
  };
  $('#btn-arbeit').onclick = async () => {
    if (S.tag.laeuft) { S.tag.segmente.push({ start: S.tag.laeuft, ende: Date.now() }); S.tag.laeuft = null; }
    else S.tag.laeuft = Date.now();
    await tagSpeichern();
  };

  $('#btn-ort').onclick = ortDialog;
  $('#orte-schnell').onclick = async (e) => {
    const b = e.target.closest('[data-schnell]');
    if (b) await ortHinzufuegen(decodeURIComponent(b.dataset.schnell));
  };
  $('#orte-liste').onclick = async (e) => {
    const b = e.target.closest('[data-ort]');
    if (!b) return;
    S.tag.orte.splice(Number(b.dataset.ort), 1);
    await tagSpeichern();
  };

  $$('[data-essen]').forEach((b) => {
    b.onclick = async () => {
      S.tag.arbeitsessen = Math.max(0, (S.tag.arbeitsessen || 0) + Number(b.dataset.essen));
      await tagSpeichern();
    };
  });
  $('#chk-nacht').onchange = async (e) => {
    S.tag.naechtigung = e.target.checked;
    await tagSpeichern();
  };
  $('#btn-grenze').onclick = grenzDialog;
  $('#btn-abschluss').onclick = abschlussDialog;
  $('#btn-korrigieren').onclick = () => tagBearbeiten(S.tag.datum);

  $('#monat-zurueck').onclick = () => monatWechseln(-1);
  $('#monat-vor').onclick = () => monatWechseln(1);
  $('#km-start').onchange = async (e) => {
    S.einst.kmStart[S.monat] = e.target.value === '' ? null : Number(e.target.value);
    await einstSpeichern(); zeichneMonat();
  };
  $('#km-ende').onchange = async (e) => {
    S.einst.kmEnde[S.monat] = e.target.value === '' ? null : Number(e.target.value);
    await einstSpeichern(); zeichneMonat();
  };
  $('#monat-liste').onclick = (e) => {
    const li = e.target.closest('[data-tag]');
    if (li) tagBearbeiten(li.dataset.tag);
  };
  $('#btn-export').onclick = exportieren;
  $('#btn-backup').onclick = backup;

  $('#foto-input').onchange = async (e) => {
    const f = e.target.files[0];
    if (!f) return;
    $('#foto-text').textContent = 'Bild wird verkleinert…';
    try {
      fotoPuffer = await fotoVerkleinern(f);
      $('#foto-vorschau').src = fotoPuffer;
      $('#foto-vorschau').hidden = false;
      $('#foto-text').hidden = true;
    } catch (err) {
      $('#foto-text').textContent = 'Bild nicht lesbar, nochmal versuchen';
    }
    e.target.value = '';
  };
  $('#btn-beleg-speichern').onclick = belegSpeichern;
  $('#beleg-liste').onclick = (e) => {
    const b = e.target.closest('[data-beleg]');
    if (b) belegLoeschen(Number(b.dataset.beleg));
  };

  // Tageswechsel, wenn die App ueber Mitternacht offen bleibt
  document.addEventListener('visibilitychange', async () => {
    if (document.hidden) return;
    if (S.tag.datum !== heuteIso()) { await laden(); }
    else zeichneHeute();
  });
  setInterval(() => { if (S.tag.laeuft) zeichneHeute(); }, 30000);
}

async function monatWechseln(richtung) {
  const [j, m] = S.monat.split('-').map(Number);
  const d = new Date(j, m - 1 + richtung, 1);
  await monatLaden(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`);
  zeichneMonat(); zeichneBelege();
}

// ------------------------------------------------------------------ Start --
verdrahten();
laden().catch((e) => {
  document.body.innerHTML =
    `<div style="padding:30px;font-family:sans-serif">
       <h2>Start fehlgeschlagen</h2><p>${e.message}</p>
       <p>Im privaten Modus von Safari ist die Datenbank gesperrt.</p>
     </div>`;
});

if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => navigator.serviceWorker.register('sw.js').catch(() => {}));
}
