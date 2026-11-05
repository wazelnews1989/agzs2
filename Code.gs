/**
 * Сводки АГЗС — сервер (Google Apps Script)
 *
 * 1. Поменяйте PIN владельца ниже.
 * 2. Выберите функцию «setup» и нажмите «Выполнить» (разрешите доступ).
 * 3. «Начать развёртывание» → «Новое развёртывание» → «Веб-приложение»,
 *    запуск от вашего имени, доступ «Все». Скопируйте ссылку — это адрес сервера.
 */

/* ======== НАСТРОЙКИ ======== */
const OWNER_PIN = '1234';            // PIN владельца — замените на свой (4–8 цифр)
const TZ = 'Europe/Samara';          // часовой пояс заправок
/* =========================== */

const APP = 'Сводки АГЗС';
const SLOTS = [
  { id: '08', label: 'Утро',  time: '8:00',  h: 8 },
  { id: '18', label: 'Вечер', time: '18:00', h: 18 },
  { id: '22', label: 'Ночь',  time: '22:00', h: 22 },
];
const SLOT_IDS = SLOTS.map(s => s.id);
const REMIND2_AFTER = 30;   // мин: повторное напоминание оператору
const SUMMARY_AFTER = 45;   // мин после срока: сводка владельцу
const LOW_PCT = 20;         // % — «мало газа»
const REPORT_DAYS = 45;

const NAMES = { st: 'Заправки', rep: 'Сводки', del: 'Привозы', dev: 'Телефоны' };
const HEAD = {
  st:  ['id', 'Название', 'Объём бочки, л', 'Счётчик при добавлении', 'Газ в бочке (база), л', 'База: время', 'База: счётчик', 'Добавлена', 'Удалена'],
  rep: ['Ключ', 'id заправки', 'Заправка', 'Дата', 'Смена', 'Показание счётчика', 'Продано, л', 'Отправлено'],
  del: ['id', 'id заправки', 'Заправка', 'Литры', 'Время'],
  dev: ['Телефон', 'id заправки', 'Заправка', 'Владелец', 'Обновлено', 'Адрес уведомлений'],
};
const FMT = {
  st:  ['@', '@', '0', '0.##', '0', 'dd.MM.yyyy HH:mm', '0.##', 'dd.MM.yyyy HH:mm', 'dd.MM.yyyy HH:mm'],
  rep: ['@', '@', '@', '@', '@', '0.##', '0.##', 'dd.MM.yyyy HH:mm'],
  del: ['@', '@', '@', '0', 'dd.MM.yyyy HH:mm'],
  dev: ['@', '@', '@', '@', 'dd.MM.yyyy HH:mm', '@'],
};

/* ---------- вход: API для приложения ---------- */
function doGet() {
  return HtmlService.createHtmlOutput('<p style="font:20px sans-serif;padding:20px">Сервер «' + APP + '» работает. Откройте само приложение по ссылке с сайта.</p>');
}
function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = ROUTES[req.a];
    if (!fn) fail_('Неизвестная команда.');
    out = { ok: true, data: fn(req) };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
const ROUTES = {
  getState: r => getState(),
  submitReport: r => submitReport(r.p),
  addDelivery: r => addDelivery(r.p),
  ownerLogin: r => ownerLogin(r.pin),
  getSettings: r => getSettings(r.pin),
  setEmail: r => setEmail(r.pin, r.on),
  addStation: r => addStation(r.pin, r.p),
  updateStation: r => updateStation(r.pin, r.p),
  fixTank: r => fixTank(r.pin, r.p),
  deleteStation: r => deleteStation(r.pin, r.id),
  deleteDelivery: r => deleteDelivery(r.pin, r.id),
  subscribe: r => subscribe(r),
  inbox: r => inbox(r.dev),
  testPush: r => testPush(r.pin, r.dev),
};

/* ---------- установка ---------- */
function setup() {
  const ss = book_();
  vapid_();
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'tick') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('tick').timeBased().everyMinutes(15).create();
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true }); // запрос разрешения на отправку уведомлений
  MailApp.getRemainingDailyQuota();                                         // и писем
  if (String(OWNER_PIN) === '1234') console.warn('Вы оставили PIN 1234. Поменяйте его в начале кода и запустите setup ещё раз.');
  console.log('Готово. Таблица с данными: ' + ss.getUrl());
}

/* ---------- хранилище ---------- */
let SS_ = null;
function book_() {
  if (SS_) return SS_;
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('SS_ID');
  if (id) { try { SS_ = SpreadsheetApp.openById(id); } catch (e) { SS_ = null; } }
  if (!SS_) { SS_ = SpreadsheetApp.create(APP + ' — данные'); p.setProperty('SS_ID', SS_.getId()); }
  ensureSheets_(SS_);
  return SS_;
}
function ensureSheets_(ss) {
  let created = false;
  Object.keys(NAMES).forEach(k => {
    let sh = ss.getSheetByName(NAMES[k]);
    if (!sh) { sh = ss.insertSheet(NAMES[k]); created = true; }
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, HEAD[k].length).setValues([HEAD[k]]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
  if (created) ss.getSheets().forEach(sh => {
    if (Object.values(NAMES).indexOf(sh.getName()) < 0 && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
}
function sheet_(k) { return book_().getSheetByName(NAMES[k]); }
function values_(k) {
  const sh = sheet_(k), n = sh.getLastRow();
  return n < 2 ? [] : sh.getRange(2, 1, n - 1, HEAD[k].length).getValues();
}
function writeRow_(k, row, values) {
  const sh = sheet_(k);
  const r = row || sh.getLastRow() + 1;
  const rng = sh.getRange(r, 1, 1, values.length);
  rng.setNumberFormats([FMT[k]]);
  rng.setValues([values]);
  return r;
}
const num_ = v => (v === '' || v == null) ? null : Number(v);
const ms_ = v => v instanceof Date ? v.getTime() : (v ? (Number(v) || new Date(v).getTime() || 0) : 0);
const dstr_ = v => v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v);
const slot_ = v => String(v).padStart(2, '0');
const byKey_ = (a, b) => a.key < b.key ? -1 : a.key > b.key ? 1 : 0;
const fmt_ = n => (Math.round(n * 100) / 100).toLocaleString('ru-RU');
const props_ = () => PropertiesService.getScriptProperties();
function fail_(msg) { throw new Error(msg); }
function safe_(fn) { try { return fn(); } catch (e) { console.error(e); } }
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) fail_('Сервер занят, нажмите ещё раз через минуту.');
  try { return fn(); } finally { lock.releaseLock(); }
}

function readStations_() {
  return values_('st').map((r, i) => ({
    row: i + 2, id: String(r[0]), name: String(r[1]), capacity: num_(r[2]), startMeter: Number(r[3]) || 0,
    baseLevel: Number(r[4]) || 0, baseAt: ms_(r[5]), baseMeter: Number(r[6]) || 0, createdAt: ms_(r[7]), deleted: !!r[8],
  })).filter(s => s.id);
}
function readReports_() {
  return values_('rep').map((r, i) => ({
    row: i + 2, key: String(r[0]), sid: String(r[1]), date: dstr_(r[3]), slot: slot_(r[4]), meter: Number(r[5]), at: ms_(r[7]),
  })).filter(r => r.key && r.sid);
}
function readDeliveries_() {
  return values_('del').map((r, i) => ({ row: i + 2, id: String(r[0]), sid: String(r[1]), liters: Number(r[3]) || 0, at: ms_(r[4]) })).filter(d => d.id);
}
function readDevices_() {
  return values_('dev').map((r, i) => {
    let sub = null; try { sub = JSON.parse(r[5]); } catch (e) {}
    return { row: i + 2, dev: String(r[0]), sid: String(r[1] || ''), owner: r[3] === true || r[3] === 'да', sub };
  }).filter(d => d.dev && d.sub && d.sub.endpoint);
}
function derive_(stations, reports, deliveries) {
  const bySt = {};
  reports.forEach(r => (bySt[r.sid] = bySt[r.sid] || []).push(r));
  stations.forEach(st => {
    const rs = (bySt[st.id] || []).sort(byKey_);
    let prev = st.startMeter;
    rs.forEach(r => { r.liters = r.meter - prev; prev = r.meter; });
    st.lastMeter = rs.length ? rs[rs.length - 1].meter : st.startMeter;
    st.lastKey = rs.length ? rs[rs.length - 1].key : '';
    const dl = deliveries.filter(d => d.sid === st.id && d.at > st.baseAt).reduce((a, d) => a + d.liters, 0);
    st.tank = st.baseLevel + dl - Math.max(0, st.lastMeter - st.baseMeter);
  });
}
function load_() {
  const stations = readStations_(), reports = readReports_(), deliveries = readDeliveries_();
  derive_(stations, reports, deliveries);
  return { stations, reports, deliveries };
}
function tzOffsetMs_(d) {
  const z = Utilities.formatDate(d, TZ, 'Z');               // например +0400
  const sign = z[0] === '-' ? -1 : 1;
  return sign * (Number(z.slice(1, 3)) * 60 + Number(z.slice(3, 5))) * 60000;
}
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function shiftDate_(key, days) {
  const p = key.split('-').map(Number);
  return new Date(Date.UTC(p[0], p[1] - 1, p[2] + days)).toISOString().slice(0, 10);
}

/* ---------- для всех ---------- */
function getState() {
  const d = load_();
  const cutoff = shiftDate_(today_(), -REPORT_DAYS), dlCut = Date.now() - 120 * 86400000;
  return {
    stations: d.stations.filter(s => !s.deleted).map(s => ({
      id: s.id, name: s.name, capacity: s.capacity, startMeter: s.startMeter, lastMeter: s.lastMeter, lastKey: s.lastKey, tank: s.tank, createdAt: s.createdAt,
    })),
    reports: d.reports.filter(r => r.date >= cutoff).map(r => ({ key: r.key, sid: r.sid, date: r.date, slot: r.slot, meter: r.meter, liters: r.liters, at: r.at })),
    deliveries: d.deliveries.filter(x => x.at >= dlCut).map(x => ({ id: x.id, sid: x.sid, liters: x.liters, at: x.at })),
    vapid: vapid_().pub,
    serverTime: Date.now(),
  };
}

function submitReport(p) {
  const res = withLock_(() => {
    const date = String(p && p.date), slot = String(p && p.slot), sid = String(p && p.sid), v = Number(p && p.meter);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || SLOT_IDS.indexOf(slot) < 0) fail_('Неверная дата или смена.');
    if (!isFinite(v) || v < 0) fail_('Введите показание счётчика цифрами.');
    const st = readStations_().find(s => s.id === sid && !s.deleted);
    if (!st) fail_('Заправка не найдена. Выберите её заново.');
    const key = date + '_' + slot;
    const mine = readReports_().filter(r => r.sid === sid).sort(byKey_);
    const before = mine.filter(r => r.key < key), after = mine.filter(r => r.key > key);
    const prev = before.length ? before[before.length - 1].meter : st.startMeter;
    if (v < prev) fail_('Показание меньше прошлого (' + fmt_(prev) + '). Проверьте цифры.');
    if (after.length && v > after[0].meter) fail_('Показание больше, чем в следующей сводке (' + fmt_(after[0].meter) + '). Проверьте цифры.');
    const existing = mine.find(r => r.key === key);
    writeRow_('rep', existing ? existing.row : null, [key, sid, st.name, date, slot, v, v - prev, new Date()]);
    recomputeSold_(sid, st.startMeter);
    const late = props_().getProperty('m_' + date + '_' + slot) ? { name: st.name, liters: v - prev, slot } : null;
    return { late, sid };
  });
  if (res.late) {
    const s = SLOTS.find(x => x.id === res.late.slot);
    safe_(() => notifyOwner_('Сводка пришла с опозданием', res.late.name + ', ' + s.label.toLowerCase() + ': ' + fmt_(res.late.liters) + ' л'));
  }
  safe_(() => lowGasCheck_(res.sid));
  return getState();
}
function recomputeSold_(sid, startMeter) {
  const sh = sheet_('rep'), n = sh.getLastRow();
  if (n < 2) return;
  const vals = sh.getRange(2, 1, n - 1, 7).getValues();
  const rows = vals.map((r, i) => ({ i, sid: String(r[1]), key: String(r[0]), meter: Number(r[5]) })).filter(r => r.sid === sid).sort(byKey_);
  const sold = vals.map(r => [r[6]]);
  let prev = startMeter;
  rows.forEach(r => { sold[r.i] = [r.meter - prev]; prev = r.meter; });
  sh.getRange(2, 7, n - 1, 1).setValues(sold);
}
function lowGasCheck_(sid) {
  const st = load_().stations.find(s => s.id === sid);
  if (!st || !st.capacity) return;
  const pct = st.tank / st.capacity * 100;
  const k = 'low_' + today_() + '_' + sid;
  if (pct < LOW_PCT && !props_().getProperty(k)) {
    props_().setProperty(k, '1');
    notifyOwner_('Мало газа: ' + st.name, 'В бочке ' + fmt_(Math.max(0, Math.round(st.tank))) + ' л (' + Math.round(pct) + '%). Пора заказывать газовоз.', true);
  }
}

function addDelivery(p) {
  const info = withLock_(() => {
    const sid = String(p && p.sid), v = Number(p && p.liters);
    if (!isFinite(v) || v <= 0 || v > 100000) fail_('Введите, сколько литров залили.');
    const st = readStations_().find(s => s.id === sid && !s.deleted);
    if (!st) fail_('Заправка не найдена.');
    writeRow_('del', null, [Utilities.getUuid().slice(0, 8), sid, st.name, v, new Date()]);
    return { name: st.name, v };
  });
  safe_(() => notifyOwner_('Привоз газа: ' + info.name, '+' + fmt_(info.v) + ' л'));
  return getState();
}

/* ---------- владелец ---------- */
function checkPin_(pin) {
  const cache = CacheService.getScriptCache();
  const bad = Number(cache.get('pin_bad') || 0);
  if (bad >= 10) fail_('Слишком много неверных попыток. Подождите 10 минут.');
  if (String(pin) !== String(OWNER_PIN)) { cache.put('pin_bad', String(bad + 1), 600); fail_('Неверный PIN.'); }
}
function ownerLogin(pin) { checkPin_(pin); return settings_(); }
function addStation(pin, p) {
  checkPin_(pin);
  return withLock_(() => {
    const name = String(p && p.name || '').trim(), cap = num_(p && p.capacity), lvl = Number(p && p.level), m = Number(p && p.meter);
    if (!name) fail_('Впишите название заправки.');
    if (!isFinite(lvl) || lvl < 0) fail_('Впишите, сколько газа в бочке.');
    if (!isFinite(m) || m < 0) fail_('Впишите показание счётчика.');
    const now = new Date();
    writeRow_('st', null, [Utilities.getUuid().slice(0, 8), name, cap, m, lvl, now, m, now, '']);
    return getState();
  });
}
function findStation_(id) {
  const st = readStations_().find(s => s.id === String(id) && !s.deleted);
  if (!st) fail_('Заправка не найдена.');
  return st;
}
function updateStation(pin, p) {
  checkPin_(pin);
  return withLock_(() => {
    const st = findStation_(p && p.id), name = String(p.name || '').trim();
    if (!name) fail_('Название не может быть пустым.');
    sheet_('st').getRange(st.row, 2, 1, 2).setValues([[name, num_(p.capacity)]]);
    return getState();
  });
}
function fixTank(pin, p) {
  checkPin_(pin);
  return withLock_(() => {
    const lvl = Number(p && p.level);
    if (!isFinite(lvl) || lvl < 0) fail_('Впишите уровень газа в литрах.');
    const st = load_().stations.find(s => s.id === String(p.id) && !s.deleted);
    if (!st) fail_('Заправка не найдена.');
    sheet_('st').getRange(st.row, 5, 1, 3).setValues([[lvl, new Date(), st.lastMeter]]);
    return getState();
  });
}
function deleteStation(pin, id) {
  checkPin_(pin);
  return withLock_(() => { sheet_('st').getRange(findStation_(id).row, 9).setValue(new Date()); return getState(); });
}
function deleteDelivery(pin, id) {
  checkPin_(pin);
  return withLock_(() => {
    const d = readDeliveries_().find(x => x.id === String(id));
    if (d) sheet_('del').deleteRow(d.row);
    return getState();
  });
}
function emailOn_() {
  const v = props_().getProperty('EMAIL_ON');
  if (v === 'yes') return true;
  if (v === 'no') return false;
  return !readDevices_().some(d => d.owner);   // по умолчанию: почта, пока у владельца нет телефона с уведомлениями
}
function settings_() {
  const devs = readDevices_(), sts = readStations_().filter(s => !s.deleted);
  return {
    sheetUrl: book_().getUrl(),
    email: Session.getEffectiveUser().getEmail(), emailOn: emailOn_(),
    ownerPhones: devs.filter(d => d.owner).length,
    phones: sts.map(s => ({ name: s.name, count: devs.filter(d => d.sid === s.id).length })),
    defaultPin: String(OWNER_PIN) === '1234',
  };
}
function getSettings(pin) { checkPin_(pin); return settings_(); }
function setEmail(pin, on) { checkPin_(pin); props_().setProperty('EMAIL_ON', on ? 'yes' : 'no'); return settings_(); }

/* ---------- телефоны и уведомления ---------- */
function subscribe(r) {
  const dev = String(r.dev || '');
  if (!/^[\w-]{8,64}$/.test(dev)) fail_('Неверный id телефона.');
  const sub = r.sub;
  if (!sub || !/^https:\/\//.test(String(sub.endpoint || ''))) fail_('Телефон не дал адрес для уведомлений.');
  const owner = !!r.owner;
  if (owner) checkPin_(r.pin);
  return withLock_(() => {
    const sid = String(r.sid || '');
    const st = sid ? readStations_().find(s => s.id === sid && !s.deleted) : null;
    const row = readDevices_().find(d => d.dev === dev);
    writeRow_('dev', row ? row.row : null, [dev, st ? st.id : '', st ? st.name : '', owner ? 'да' : '', new Date(),
      JSON.stringify({ endpoint: sub.endpoint, keys: sub.keys || {} })]);
    return { ok: true };
  });
}
function inbox(dev) {
  dev = String(dev || '');
  return withLock_(() => {
    const k = 'in_' + dev, raw = props_().getProperty(k);
    if (!raw) return [];
    props_().deleteProperty(k);
    let list = []; try { list = JSON.parse(raw); } catch (e) {}
    return list.filter(m => Date.now() - m.at < 24 * 3600000);
  });
}
function testPush(pin, dev) {
  checkPin_(pin);
  const d = readDevices_().find(x => x.dev === String(dev));
  if (!d) fail_('Этот телефон ещё не подключён к уведомлениям.');
  const r = push_([d], { title: 'Проверка', body: 'Уведомления работают.' });
  if (!r.sent) fail_('Не удалось отправить уведомление: ' + (r.errors[0] || 'неизвестная ошибка'));
  return 'Уведомление отправлено. Оно придёт через несколько секунд.';
}
function notifyOwner_(title, body, urgent) {
  const owners = readDevices_().filter(d => d.owner);
  if (owners.length) safe_(() => push_(owners, { title, body, urgent: !!urgent }));
  if (emailOn_()) safe_(() => MailApp.sendEmail(Session.getEffectiveUser().getEmail(), APP + ': ' + title, body));
}
/** Кладёт сообщение во «входящие» телефона и будит его пустым push-уведомлением. */
function push_(devs, msg) {
  if (!devs.length) return { sent: 0, errors: [] };
  const p = props_();
  const id = Utilities.getUuid().slice(0, 8);
  devs.forEach(d => {
    const k = 'in_' + d.dev;
    let list = []; try { list = JSON.parse(p.getProperty(k) || '[]'); } catch (e) {}
    list.push({ id, title: msg.title, body: msg.body, urgent: !!msg.urgent, at: Date.now() });
    p.setProperty(k, JSON.stringify(list.slice(-10)));
  });
  const v = vapid_();
  const reqs = devs.map(d => ({
    url: d.sub.endpoint, method: 'post', payload: '', muteHttpExceptions: true,
    headers: { Authorization: 'vapid t=' + vapidJwt_(d.sub.endpoint) + ', k=' + v.pub, TTL: '43200', Urgency: 'high' },
  }));
  const res = UrlFetchApp.fetchAll(reqs);
  let sent = 0; const errors = [], gone = [];
  res.forEach((r, i) => {
    const c = r.getResponseCode();
    if (c >= 200 && c < 300) sent++;
    else { if (c === 404 || c === 410) gone.push(devs[i]); errors.push(c + ' ' + String(r.getContentText()).slice(0, 120)); }
  });
  if (gone.length) safe_(() => {
    const sh = sheet_('dev');
    gone.map(g => g.row).sort((a, b) => b - a).forEach(r => sh.deleteRow(r));
  });
  if (errors.length) console.warn('push errors', errors);
  return { sent, errors };
}

/* ---------- по расписанию (каждые 15 минут) ---------- */
function tick() {
  const now = new Date();
  const date = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
  const hm = Utilities.formatDate(now, TZ, 'HH:mm').split(':').map(Number);
  const mins = hm[0] * 60 + hm[1];
  const p = props_();
  SLOTS.forEach(s => {
    const base = s.h * 60;
    if (mins >= base && mins < base + 20) once_(p, 'r_' + date + '_' + s.id, () => remind_(date, s, false));
    if (mins >= base + REMIND2_AFTER && mins < base + REMIND2_AFTER + 20) once_(p, 'q_' + date + '_' + s.id, () => remind_(date, s, true));
    if (mins >= base + SUMMARY_AFTER && mins < base + SUMMARY_AFTER + 120) once_(p, 'm_' + date + '_' + s.id, () => summary_(date, s));
  });
  if (hm[0] === 3) cleanupFlags_(p, date);
}
function once_(p, key, fn) { if (p.getProperty(key)) return; p.setProperty(key, '1'); safe_(fn); }
function cleanupFlags_(p, date) {
  const old = shiftDate_(date, -3);
  Object.keys(p.getProperties()).forEach(k => {
    const m = /^(?:[rqm]|low)_(\d{4}-\d{2}-\d{2})_/.exec(k);
    if (m && m[1] < old) p.deleteProperty(k);
  });
}
function remind_(date, s, again) {
  const d = load_(), devs = readDevices_();
  const p = date.split('-').map(Number);
  const dueMs = Date.UTC(p[0], p[1] - 1, p[2], s.h) - tzOffsetMs_(new Date());
  d.stations.filter(st => !st.deleted && !(st.createdAt > dueMs)).forEach(st => {
    if (d.reports.some(r => r.sid === st.id && r.date === date && r.slot === s.id)) return;
    const targets = devs.filter(x => x.sid === st.id);
    if (!targets.length) return;
    push_(targets, {
      title: again ? 'Сводка ещё не отправлена' : 'Пора отправить сводку',
      body: s.label + ', ' + s.time + ' · ' + st.name + '. Нажмите, чтобы ввести показание счётчика.',
      urgent: true,
    });
  });
}
function summary_(date, s) {
  const d = load_(), sts = d.stations.filter(st => !st.deleted);
  if (!sts.length) return;
  let total = 0, missing = 0; const low = [];
  const p = date.split('-').map(Number);
  const dueMs = Date.UTC(p[0], p[1] - 1, p[2], s.h) - tzOffsetMs_(new Date());
  const lines = sts.filter(st => !(st.createdAt > dueMs)).map(st => {
    const r = d.reports.find(x => x.sid === st.id && x.date === date && x.slot === s.id);
    const pct = st.capacity ? Math.round(st.tank / st.capacity * 100) : null;
    if (pct != null && pct < LOW_PCT) low.push(st.name);
    if (r) total += r.liters; else missing++;
    return st.name + ': ' + (r ? fmt_(r.liters) + ' л' : 'НЕТ СВОДКИ') + ' · в бочке ' + fmt_(Math.max(0, Math.round(st.tank))) + ' л' + (pct != null ? ' (' + pct + '%)' : '');
  });
  const title = 'Сводка «' + s.label + '»: ' + fmt_(total) + ' л' + (missing ? ' · нет сводок: ' + missing : '');
  let body = lines.join('\n');
  if (low.length) body += '\nМало газа: ' + low.join(', ');
  notifyOwner_(title, body, missing > 0);
}

/* ---------- ключи уведомлений (VAPID, ECDSA P-256) ---------- */
function vapid_() {
  const p = props_();
  let pub = p.getProperty('VAPID_PUB'), d = p.getProperty('VAPID_D');
  if (!pub || !d) {
    let k = BigInt(0);
    while (k <= BigInt(0) || k >= EC_.N) k = bytesToBig_(unsigned_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
      Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid() + Date.now() + Math.random())));
    const Q = ecMul_(k, [EC_.GX, EC_.GY]);
    pub = b64u_([4].concat(bigToBytes_(Q[0]), bigToBytes_(Q[1])));
    d = k.toString(16);
    p.setProperty('VAPID_PUB', pub); p.setProperty('VAPID_D', d);
  }
  return { pub, d: BigInt('0x' + d) };
}
function vapidJwt_(endpoint) {
  const aud = (/^(https:\/\/[^\/]+)/.exec(endpoint) || [])[1];
  const cache = CacheService.getScriptCache(), ck = 'jwt_' + aud.replace(/\W/g, '_');
  const hit = cache.get(ck);
  if (hit) return hit;
  const exp = Math.floor(Date.now() / 1000) + 12 * 3600;
  const enc = o => b64u_(unsigned_(Utilities.newBlob(JSON.stringify(o)).getBytes()));
  const input = enc({ typ: 'JWT', alg: 'ES256' }) + '.' + enc({ aud, exp, sub: 'mailto:' + (Session.getEffectiveUser().getEmail() || 'owner@example.com') });
  const jwt = input + '.' + b64u_(ecdsaSign_(input, vapid_().d));
  cache.put(ck, jwt, 6 * 3600);
  return jwt;
}
const EC_ = {
  P: BigInt('0xffffffff00000001000000000000000000000000ffffffffffffffffffffffff'),
  N: BigInt('0xffffffff00000000ffffffffffffffffbce6faada7179e84f3b9cac2fc632551'),
  GX: BigInt('0x6b17d1f2e12c4247f8bce6e563a440f277037d812deb33a0f4a13945d898c296'),
  GY: BigInt('0x4fe342e2fe1a7f9b8ee7eb4a7c0f9e162bce33576b315ececbb6406837bf51f5'),
};
function mod_(a, m) { const r = a % m; return r < BigInt(0) ? r + m : r; }
function inv_(a, m) {
  let lo = BigInt(1), hi = BigInt(0), low = mod_(a, m), high = m;
  while (low > BigInt(1)) { const q = high / low; [lo, hi] = [hi - lo * q, lo]; [low, high] = [high - low * q, low]; }
  return mod_(lo, m);
}
function ecAdd_(a, b) {
  if (!a) return b; if (!b) return a;
  const P = EC_.P;
  let l;
  if (a[0] === b[0]) {
    if (mod_(a[1] + b[1], P) === BigInt(0)) return null;
    l = mod_((BigInt(3) * a[0] * a[0] - BigInt(3)) * inv_(BigInt(2) * a[1], P), P);
  } else l = mod_((b[1] - a[1]) * inv_(b[0] - a[0], P), P);
  const x = mod_(l * l - a[0] - b[0], P);
  return [x, mod_(l * (a[0] - x) - a[1], P)];
}
function ecMul_(k, pt) {
  let r = null, q = pt;
  while (k > BigInt(0)) { if (k & BigInt(1)) r = ecAdd_(r, q); q = ecAdd_(q, q); k >>= BigInt(1); }
  return r;
}
function unsigned_(b) { return Array.prototype.map.call(b, x => x & 255); }
function signed_(b) { return b.map(x => x > 127 ? x - 256 : x); }
function bytesToBig_(b) { return BigInt('0x' + (b.map(x => (x & 255).toString(16).padStart(2, '0')).join('') || '0')); }
function bigToBytes_(n) { const h = n.toString(16).padStart(64, '0'); const o = []; for (let i = 0; i < 64; i += 2) o.push(parseInt(h.slice(i, i + 2), 16)); return o; }
function b64u_(bytes) { return Utilities.base64EncodeWebSafe(signed_(bytes)).replace(/=+$/, ''); }
function hmac_(key, data) { return unsigned_(Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, signed_(data), signed_(key))); }
/** ECDSA P-256 / SHA-256 с детерминированным k (RFC 6979). Возвращает r||s, 64 байта. */
function ecdsaSign_(text, d) {
  const N = EC_.N;
  const h = unsigned_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8));
  const z = mod_(bytesToBig_(h), N);
  const x = bigToBytes_(d), hz = bigToBytes_(z);
  let V = new Array(32).fill(1), K = new Array(32).fill(0);
  K = hmac_(K, V.concat([0], x, hz)); V = hmac_(K, V);
  K = hmac_(K, V.concat([1], x, hz)); V = hmac_(K, V);
  for (;;) {
    V = hmac_(K, V);
    const k = bytesToBig_(V);
    if (k > BigInt(0) && k < N) {
      const R = ecMul_(k, [EC_.GX, EC_.GY]);
      const r = mod_(R[0], N);
      const s = mod_(inv_(k, N) * (z + r * d), N);
      if (r > BigInt(0) && s > BigInt(0)) return bigToBytes_(r).concat(bigToBytes_(s));
    }
    K = hmac_(K, V.concat([0])); V = hmac_(K, V);
  }
}
