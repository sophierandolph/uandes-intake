/**
 * UANDES participant profiles: one Google Slide per person, built from the "Responses" tab.
 * Adds a menu to the Sheet: UANDES > Build / refresh profile slides.
 *
 * Safe to re-run any time:
 *  - Each slide is matched to its row by the response_id in the speaker notes.
 *  - On refresh, only elements titled "auto" are redrawn. Your pasted photo, the
 *    "photo" placeholder and the "ideas" box (Dream weaving notes) are never touched.
 *  - New people get a new slide at the end.
 * The deck lives in the "Intake Survey" Drive folder; its ID is kept in Script Properties.
 * The slides hold allergy/safety details: share with Hesam and the UANDES IR staffer only.
 */
const PROFILE_DECK_NAME = 'UANDES Participant Profiles';
const PROFILE_FOLDER_ID = '10CxWivZ62-3InnRuop5AnPga60hCaBWp'; // UAndes / Rice > Intake Survey
const C = {
  navy: '#0037A6', cyan: '#17B9E2', mid: '#0E6BA5', green: '#39B54A',
  orange: '#F15A29', red: '#EE243E', soft: '#F2F6FC', line: '#C9D3E6', grey: '#888888', black: '#000000'
};
const MONO = 'IBM Plex Mono';
const DISPLAY = 'Archivo Narrow';
const INSET = 7.2; // Slides' default text box padding; boxes are offset so text lands on the grid.

function onOpen() {
  SpreadsheetApp.getUi().createMenu('UANDES')
    .addItem('Build / refresh profile slides', 'buildProfiles')
    .addItem('Open profile deck', 'showProfileDeckLink')
    .addToUi();
}

function buildProfiles() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Responses');
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  const people = values
    .map(function (row) { const o = {}; headers.forEach(function (h, i) { o[h] = row[i] == null ? '' : String(row[i]).trim(); }); return o; })
    .filter(function (p) { return p.response_id && p.full_name.toUpperCase() !== 'TEST' && p.response_id.indexOf('TEST-') !== 0; });

  const deck = getOrCreateDeck_();
  const byId = {};
  deck.getSlides().forEach(function (s) {
    const m = s.getNotesPage().getSpeakerNotesShape().getText().asString().match(/response_id:\s*(\S+)/);
    if (m) byId[m[1]] = s;
  });

  let added = 0, refreshed = 0;
  people.forEach(function (p) {
    let slide = byId[p.response_id];
    if (slide) {
      slide.getPageElements().forEach(function (el) { if (el.getTitle() === 'auto') el.remove(); });
      refreshed++;
    } else {
      slide = deck.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      slide.getNotesPage().getSpeakerNotesShape().getText().setText('response_id: ' + p.response_id);
      drawFixed_(slide);
      added++;
    }
    drawProfile_(slide, p);
  });

  SpreadsheetApp.getActive().toast(added + ' new, ' + refreshed + ' refreshed. Deck: ' + deck.getUrl(), 'Profiles done', 15);
}

function showProfileDeckLink() {
  const deck = getOrCreateDeck_();
  const html = HtmlService.createHtmlOutput('<p style="font-family:sans-serif"><a href="' + deck.getUrl() + '" target="_blank">Open ' + PROFILE_DECK_NAME + '</a></p>').setWidth(320).setHeight(80);
  SpreadsheetApp.getUi().showModalDialog(html, 'Profile deck');
}

function getOrCreateDeck_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PROFILE_DECK_ID');
  if (id) {
    try { return SlidesApp.openById(id); } catch (e) { /* deleted or no access: make a new one */ }
  }
  const deck = SlidesApp.create(PROFILE_DECK_NAME);
  deck.getSlides().forEach(function (s) { s.remove(); }); // drop the default title slide
  DriveApp.getFileById(deck.getId()).moveTo(DriveApp.getFolderById(PROFILE_FOLDER_ID));
  props.setProperty('PROFILE_DECK_ID', deck.getId());
  return deck;
}

/* ---------- drawing ---------- */

// Elements drawn once per new slide and never redrawn (you can paste over / write in them).
function drawFixed_(slide) {
  const ph = rect_(slide, 20, 30, 186, 169, C.soft, C.line, true);
  ph.setTitle('photo');
  ph.getText().setText('Paste photo here\n(from LinkedIn)');
  style_(ph.getText(), { size: 7.5, color: C.mid });
  ph.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  ph.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

  const ideas = rect_(slide, 468, 206, 232, 112, '#FFFFFF', C.orange, false);
  ideas.setTitle('ideas');
  ideas.getBorder().setWeight(1.5);
  ideas.getText().setText('Ideas (Sophie + Hesam fill in)');
  style_(ideas.getText(), { size: 7, color: C.grey });
  ideas.setContentAlignment(SlidesApp.ContentAlignment.TOP);
}

function drawProfile_(slide, p) {
  // Top hatch band (solid cyan in Slides).
  rect_(slide, 0, 0, 720, 7, C.cyan);

  // ---- Left: who they are
  kicker_(slide, 20, 17, 186, 'UANDES × Rice · Mar 2027', C.cyan);
  const name = (p.full_name || p.call_me || 'Unnamed').toUpperCase();
  const nameSize = name.length <= 16 ? 26 : name.length <= 24 ? 21 : 17;
  text_(slide, 20, 205, 186, 58, name.split(' ').length > 1 ? splitName_(name) : name,
    { font: DISPLAY, size: nameSize, bold: true, color: C.navy });
  text_(slide, 20, 266, 186, 14, p.call_me ? 'Call me: ' + p.call_me : '', { size: 9.5, bold: true });
  const meta = [p.program_year, p.age ? 'Age ' + p.age : '', p.been_to_us ? 'Been to the US: ' + p.been_to_us : '']
    .filter(String).join(' · ');
  text_(slide, 20, 283, 186, 22, meta, { size: 7.5, color: '#333333' });
  if (p.linkedin) {
    const url = /^https?:\/\//i.test(p.linkedin) ? p.linkedin : 'https://' + p.linkedin;
    const li = text_(slide, 20, 304, 186, 12, 'LinkedIn ↗', { size: 7.5, bold: true, color: C.navy });
    li.getText().getTextStyle().setLinkUrl(url);
  }

  // ---- Middle: what they bring
  const MX = 222, MW = 222;
  kicker_(slide, MX, 17, MW, 'Why they joined', C.navy);
  text_(slide, MX, 28, MW, 50, withOther_(p.why_join, p.why_join_other).split('; ').filter(String).map(function (s) { return '▪ ' + s; }).join('\n'), { size: 7.5 });

  kicker_(slide, MX, 84, MW, 'A problem they care about', C.navy);
  if (p.problem) rect_(slide, MX, 95, 2.5, 36, C.navy);
  text_(slide, MX + 7, 95, MW - 7, 40, p.problem ? '"' + p.problem + '"' : '', { size: 9.5, bold: true });

  kicker_(slide, MX, 142, MW, 'Prior experience', C.navy);
  text_(slide, MX, 153, MW, 36, p.prior_experience, { size: 8 });

  kicker_(slide, MX, 196, MW, 'Self-rating, 0 to 10', C.navy);
  [['Entrepreneurship exp.', p.exp_entrepreneurship], ['Spotting problems', p.ability_identify], ['Field research', p.ability_field]]
    .forEach(function (m, i) { meter_(slide, MX, 208 + i * 14, m[0], m[1]); });

  kicker_(slide, MX, 256, MW, 'Anything else', C.navy);
  text_(slide, MX, 267, MW, 50, p.questions_else, { size: 8 });

  // ---- Right: dream weaving
  const RX = 468, RW = 232;
  rect_(slide, RX, 17, RW, 1.8, C.orange);
  kicker_(slide, RX, 24, RW, 'Dream weaving', C.orange);
  const dream = [
    ['Top 3 in Houston', p.top3, 24],
    ['Texas food to try first', withOther_(p.texas_food, p.texas_food_other).replace(/; /g, ', '), 14],
    ['Hype song', p.hype_song, 14],
    ['Recharges by', p.recharge, 14],
    ['Snacks', p.snacks, 24]
  ];
  let y = 38;
  dream.forEach(function (d) {
    text_(slide, RX, y, RW, 9, d[0].toUpperCase(), { size: 6, bold: true, color: C.mid });
    text_(slide, RX, y + 9, RW, d[2], d[1], { size: 8.5 });
    y += 11 + d[2];
  });

  // ---- Bottom: care strip
  rect_(slide, 0, 331, 720, 74, C.soft);
  rect_(slide, 0, 331, 720, 1.2, C.navy);
  const safety = p.safety;
  const safetyFlag = safety && !/^(none|no|n\/?a|nothing|nope)\.?$/i.test(safety);
  const cols = [
    [20, 130, 'Food', withOther_(p.food_needs, p.food_needs_other), C.green, null],
    [160, 170, 'Safety', safetyFlag ? '▲ ' + safety : (safety || ''), C.red, safetyFlag ? C.red : null],
    [340, 70, 'Photos', p.photos, C.cyan, null],
    [420, 55, 'T-shirt', p.tshirt, C.cyan, null],
    [485, 70, 'US phone', p.phone_us, C.cyan, null],
    [565, 140, 'WhatsApp', p.whatsapp, C.cyan, null]
  ];
  cols.forEach(function (c) {
    kicker_(slide, c[0], 340, c[1], c[2], c[4]);
    text_(slide, c[0], 351, c[1], 48, c[3], { size: 8, bold: !!c[5], color: c[5] || C.black });
  });
}

/* ---------- helpers ---------- */

function withOther_(v, other) {
  if (!v) return '';
  return v.split('; ').map(function (s) { return s === 'Other' ? (other || '') : s; }).filter(String).join('; ');
}

// Two-line name: first word on top, the rest below (MARÍA JOSÉ / FERNÁNDEZ stays readable).
function splitName_(name) {
  const parts = name.split(' ');
  return parts[0] + '\n' + parts.slice(1).join(' ');
}

function kicker_(slide, x, y, w, label, color) {
  rect_(slide, x, y + 2, 4, 4, color);
  text_(slide, x + 8, y, w - 8, 9, label.toUpperCase(), { size: 6, bold: true, color: C.mid });
}

function meter_(slide, x, y, label, raw) {
  const n = Math.max(0, Math.min(10, Number(raw) || 0));
  text_(slide, x, y, 96, 10, label, { size: 7 });
  rect_(slide, x + 98, y + 2.5, 104, 5.5, C.soft);
  if (n > 0) rect_(slide, x + 98, y + 2.5, 104 * n / 10, 5.5, C.navy);
  text_(slide, x + 207, y, 16, 10, raw === '' ? '' : String(n), { size: 7.5, bold: true });
}

function text_(slide, x, y, w, h, str, o) {
  o = o || {};
  const box = slide.insertTextBox(str && String(str).trim() ? String(str) : ' ', x - INSET, y - INSET, w + 2 * INSET, h + 2 * INSET);
  box.setTitle('auto');
  box.setContentAlignment(SlidesApp.ContentAlignment.TOP);
  style_(box.getText(), o);
  if (o.lineSpacing) box.getText().getParagraphStyle().setLineSpacing(o.lineSpacing);
  return box;
}

function style_(range, o) {
  range.getTextStyle()
    .setFontFamily(o.font || MONO)
    .setFontSize(o.size || 8)
    .setBold(!!o.bold)
    .setForegroundColor(o.color || C.black);
}

function rect_(slide, x, y, w, h, fill, border, dashed) {
  const s = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, w, h);
  s.setTitle('auto');
  s.getFill().setSolidFill(fill);
  if (border) {
    s.getBorder().getLineFill().setSolidFill(border);
    s.getBorder().setWeight(1);
    if (dashed) s.getBorder().setDashStyle(SlidesApp.DashStyle.DASH);
  } else {
    s.getBorder().setTransparent();
  }
  return s;
}
