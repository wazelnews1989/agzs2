/**
 * Сводки АГЗС — логика сервера.
 * Этот файл загружает «загрузчик» в Google Apps Script (Code.gs) прямо с GitHub,
 * поэтому обновления сервера приходят сами. Секретов здесь нет: PIN хранится в свойствах скрипта.
 */

/* ======== НАСТРОЙКИ ======== */
const APP_VERSION = '05.11.2026 14:30';
const OWNER_PIN_DEFAULT = '1234';    // используется, только если PIN ещё не сохранён загрузчиком
const OPERATOR_PIN = '704747';       // код входа для операторов (можно поменять в приложении)
const TZ = 'Europe/Samara';          // часовой пояс заправок
/* =========================== */

const APP = 'Сводки АГЗС';
const SLOTS = [
  { id: '08', label: 'Утро',  time: '8:00',  h: 8 },
  { id: '14', label: 'День',  time: '14:00', h: 14 },
  { id: '18', label: 'Вечер', time: '18:00', h: 18 },
  { id: '22', label: 'Ночь',  time: '22:00', h: 22 },
];
const SLOT_IDS = SLOTS.map(s => s.id);
const REMIND2_AFTER = 30;   // мин: повторное напоминание оператору
const SUMMARY_AFTER = 45;   // мин после срока: сводка владельцу
const LOW_PCT = 20;         // % — «мало газа»
const REPORT_DAYS = 70;

const NAMES = { st: 'Заправки', rep: 'Сводки за смены', del: 'Привозы', dev: 'Устройства' };
const HEAD = {
  st:  ['id', 'Название', 'Объём бочки, л', 'Цена литра, ₽', 'Газ в бочке (замер владельца), л', 'Замер: время', 'Замер: литров с 8:00', 'Добавлена', 'Удалена'],
  rep: ['Ключ', 'id заправки', 'Заправка', 'Дата', 'Смена', 'Литры (18/22 — с 8:00, утро — за сутки)', 'Остаток в бочке, л', 'Расчётный остаток, л', 'Расхождение, л', 'Касса, ₽', 'Отправлено', 'Безнал, ₽'],
  del: ['id', 'id заправки', 'Заправка', 'Литры', 'Время'],
  dev: ['id телефона', 'Ключ', 'id заправки', 'Заправка', 'Кто', 'Владелец', 'Подключён', 'Был в сети', 'Статус', 'Адрес уведомлений', 'Устройство'],
};
const FMT = {
  st:  ['@', '@', '0', '0.00', '0', 'dd.MM.yyyy HH:mm', '0.##', 'dd.MM.yyyy HH:mm', 'dd.MM.yyyy HH:mm'],
  rep: ['@', '@', '@', '@', '@', '0.##', '0', '0', '+0;-0;0', '#,##0', 'dd.MM.yyyy HH:mm', '#,##0'],
  del: ['@', '@', '@', '0', 'dd.MM.yyyy HH:mm'],
  dev: ['@', '@', '@', '@', '@', '@', 'dd.MM.yyyy HH:mm', 'dd.MM.yyyy HH:mm', '@', '@', '@'],
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
  getState: r => { const a = auth_(r); return getState(a.role === 'op' ? a.sid : null); },
  submitReport: r => { const a = auth_(r); ownSid_(a, r.p); submitReport(r.p); return getState(a.role === 'op' ? a.sid : null); },
  addDelivery: r => { const a = auth_(r); ownSid_(a, r.p); addDelivery(r.p); return getState(a.role === 'op' ? a.sid : null); },
  opStations: r => opStations(r.code),
  opRegister: r => opRegister(r),
  ownerLogin: r => ownerLogin(r.pin, r.dev, r.ua),
  getSettings: r => getSettings(r.pin),
  setEmail: r => setEmail(r.pin, r.on),
  setOpPin: r => setOpPin(r.pin, r.code),
  setOwnerPin: r => setOwnerPin(r.pin, r.code),
  reloadServer: r => { checkPin_(r.pin); if (typeof bootReload_ !== 'function') fail_('Этот сервер ещё не умеет обновляться сам.'); return bootReload_(); },
  revokeDevice: r => setDeviceStatus(r.pin, r.target, false),
  restoreDevice: r => setDeviceStatus(r.pin, r.target, true),
  addStation: r => addStation(r.pin, r.p),
  updateStation: r => updateStation(r.pin, r.p),
  fixTank: r => fixTank(r.pin, r.p),
  deleteStation: r => deleteStation(r.pin, r.id),
  deleteDelivery: r => deleteDelivery(r.pin, r.id),
  subscribe: r => subscribe(r),
  inbox: r => inbox(r.dev, r.tok),
  testPush: r => testPush(r.pin, r.dev),
};

/* ---------- доступ: владелец по PIN, операторы по ключу своего телефона ---------- */
function auth_(r) {
  if (r.pin) { checkPin_(r.pin); return { role: 'owner' }; }
  const d = deviceByKey_(r.dev, r.tok);
  if (!d) fail_('NEED_LOGIN');
  if (d.status === 'отключён') fail_('REVOKED');
  if (d.owner) fail_('NEED_LOGIN');
  touch_(d);
  return { role: 'op', sid: d.sid, d };
}
function ownSid_(a, p) {
  if (a.role === 'op' && String(p && p.sid) !== a.sid) fail_('Можно отправлять сводки только по своей заправке.');
}
function deviceByKey_(dev, tok) {
  dev = String(dev || ''); tok = String(tok || '');
  if (!dev || !tok) return null;
  const d = readDevices_().find(x => x.dev === dev);
  return d && d.tok === tok ? d : null;
}
function touch_(d) {
  if (Date.now() - d.seen > 10 * 60000) safe_(() => sheet_('dev').getRange(d.row, 8).setNumberFormat('dd.MM.yyyy HH:mm').setValue(new Date()));
}
function ownerPin_() { return props_().getProperty('OWNER_PIN') || OWNER_PIN_DEFAULT; }
function opPin_() { return props_().getProperty('OP_PIN') || String(OPERATOR_PIN); }
function checkOpPin_(code) {
  const cache = CacheService.getScriptCache();
  const bad = Number(cache.get('op_bad') || 0);
  if (bad >= 15) fail_('Слишком много неверных попыток. Подождите 10 минут.');
  if (String(code || '').trim() !== opPin_()) { cache.put('op_bad', String(bad + 1), 600); fail_('Неверный код. Спросите код у владельца.'); }
}
function opStations(code) {
  checkOpPin_(code);
  return readStations_().filter(s => !s.deleted).map(s => ({ id: s.id, name: s.name }));
}
function opRegister(r) {
  checkOpPin_(r.code);
  const dev = String(r.dev || ''), who = String(r.who || '').trim().slice(0, 60);
  if (!/^[\w-]{8,64}$/.test(dev)) fail_('Неверный id телефона.');
  if (!who) fail_('Напишите, как вас зовут.');
  return withLock_(() => {
    const st = readStations_().find(s => s.id === String(r.sid) && !s.deleted);
    if (!st) fail_('Заправка не найдена.');
    const old = readDevices_().find(x => x.dev === dev);
    if (old && old.status === 'отключён') fail_('Этот телефон отключён владельцем. Обратитесь к владельцу.');
    const tok = Utilities.getUuid().replace(/-/g, '');
    writeRow_('dev', old ? old.row : null, [dev, tok, st.id, st.name, who, '', old ? new Date(old.at) : new Date(), new Date(), 'активен',
      old && old.sub ? JSON.stringify(old.sub) : '', String(r.ua || '').slice(0, 200)]);
    return { tok, sid: st.id };
  });
}
function setDeviceStatus(pin, target, active) {
  checkPin_(pin);
  return withLock_(() => {
    const d = readDevices_().find(x => x.dev === String(target));
    if (!d) fail_('Телефон не найден.');
    sheet_('dev').getRange(d.row, 9).setValue(active ? 'активен' : 'отключён');
    if (!active) props_().deleteProperty('in_' + d.dev);
    return getState(null);
  });
}
function setOwnerPin(pin, code) {
  checkPin_(pin);
  code = String(code || '').trim();
  if (!/^\d{4,8}$/.test(code)) fail_('PIN — от 4 до 8 цифр.');
  if (code === '1234') fail_('Придумайте PIN сложнее, чем 1234.');
  if (code === opPin_()) fail_('PIN владельца не должен совпадать с кодом операторов.');
  props_().setProperty('OWNER_PIN', code);
  return settings_();
}
function setOpPin(pin, code) {
  checkPin_(pin);
  code = String(code || '').trim();
  if (!/^\d{4,8}$/.test(code)) fail_('Код — от 4 до 8 цифр.');
  if (code === ownerPin_()) fail_('Код операторов не должен совпадать с PIN владельца.');
  props_().setProperty('OP_PIN', code);
  return settings_();
}

/* ---------- установка ---------- */
function setup() {
  const ss = book_();
  vapid_();
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'tick') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('tick').timeBased().everyMinutes(15).create();
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true }); // запрос разрешения на отправку уведомлений
  MailApp.getRemainingDailyQuota();                                         // и писем
  if (ownerPin_() === '1234') console.warn('Вы оставили PIN 1234. Поменяйте его в начале кода и запустите setup ещё раз.');
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
  const st0 = ss.getSheetByName(NAMES.st);
  if (st0 && st0.getLastRow() >= 1 && String(st0.getRange(1, 4).getValue()) === '—') st0.getRange(1, 4).setValue('Цена литра, ₽');
  if (st0 && st0.getLastRow() >= 1 && String(st0.getRange(1, 4).getValue()) === 'Счётчик при добавлении') {
    // переход со счётчика колонки на литры смены: старые показания счётчика больше не нужны
    st0.getRange(1, 1, 1, HEAD.st.length).setValues([HEAD.st]);
    const n = st0.getLastRow();
    if (n > 1) { st0.getRange(2, 4, n - 1, 1).setValue(''); st0.getRange(2, 7, n - 1, 1).setValue(0); }
  }
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
    row: i + 2, id: String(r[0]), name: String(r[1]), capacity: num_(r[2]), price: num_(r[3]),
    baseLevel: Number(r[4]) || 0, baseAt: ms_(r[5]), baseCum: Number(r[6]) || 0, createdAt: ms_(r[7]), deleted: !!r[8],
  })).filter(s => s.id);
}
function readReports_() {
  return values_('rep').map((r, i) => ({
    row: i + 2, key: String(r[0]), sid: String(r[1]), date: dstr_(r[3]), slot: slot_(r[4]), value: Number(r[5]) || 0,
    tank: num_(r[6]), cash: num_(r[9]), at: ms_(r[10]), card: num_(r[11]),
  })).filter(r => r.key && r.sid);
}
function readDeliveries_() {
  return values_('del').map((r, i) => ({ row: i + 2, id: String(r[0]), sid: String(r[1]), liters: Number(r[3]) || 0, at: ms_(r[4]) })).filter(d => d.id);
}
function readDevices_() {
  return values_('dev').map((r, i) => {
    let sub = null; try { sub = JSON.parse(r[9]); } catch (e) {}
    if (!sub || !sub.endpoint) sub = null;
    return { row: i + 2, dev: String(r[0]), tok: String(r[1] || ''), sid: String(r[2] || ''), stName: String(r[3] || ''), who: String(r[4] || ''),
      owner: r[5] === true || r[5] === 'да', at: ms_(r[6]), seen: ms_(r[7]), status: String(r[8] || 'активен'), sub, ua: String(r[10] || '') };
  }).filter(d => d.dev);
}
const pushable_ = d => d.sub && d.status !== 'отключён';
function workDayOfMs_(ms) { return Utilities.formatDate(new Date(ms - 8 * 3600000), TZ, 'yyyy-MM-dd'); }
function derive_(stations, reports, deliveries) {
  stations.forEach(st => {
    const reps = reports.filter(r => r.sid === st.id), dels = deliveries.filter(d => d.sid === st.id);
    st.days = computeStation(st, reps, dels, workDayOfMs_);
    st.cumNow = dayCum(st.days[workDayOfMs_(Date.now())]);
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
function dm_(key) { const p = key.split('-'); return Number(p[2]) + '.' + p[1]; }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function shiftDate_(key, days) {
  const p = key.split('-').map(Number);
  return new Date(Date.UTC(p[0], p[1] - 1, p[2] + days)).toISOString().slice(0, 10);
}

/* ---------- для всех ---------- */
function getState(onlySid) {
  const d = load_();
  const cutoff = shiftDate_(today_(), -REPORT_DAYS), dlCut = Date.now() - 120 * 86400000;
  const mine = x => !onlySid || x.sid === onlySid;
  return {
    stations: d.stations.filter(s => !s.deleted && (!onlySid || s.id === onlySid)).map(s => ({
      id: s.id, name: s.name, capacity: s.capacity, price: s.price, tank: s.tank, tankFrom: s.tankFrom, tankAt: s.tankAt, createdAt: s.createdAt,
    })),
    reports: d.reports.filter(r => r.date >= cutoff && mine(r)).map(r => ({
      key: r.key, sid: r.sid, date: r.date, slot: r.slot, value: r.value, sold: r.sold, day: r.day, at: r.at,
      tank: r.tank, cash: r.cash, card: r.card, expected: r.expected == null ? null : r.expected, diff: r.diff == null ? null : r.diff,
    })),
    deliveries: d.deliveries.filter(x => x.at >= dlCut && mine(x)).map(x => ({ id: x.id, sid: x.sid, liters: x.liters, at: x.at })),
    devices: onlySid ? [] : readDevices_().map(x => ({ dev: x.dev, sid: x.sid, who: x.who, owner: x.owner, at: x.at, seen: x.seen,
      active: x.status !== 'отключён', push: !!x.sub, ua: x.ua })),
    mySid: onlySid || null,
    vapid: vapid_().pub,
    serverTime: Date.now(),
  };
}

function submitReport(p) {
  const res = withLock_(() => {
    const date = String(p && p.date), slot = String(p && p.slot), sid = String(p && p.sid), v = Number(p && p.liters);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || SLOT_IDS.indexOf(slot) < 0) fail_('Неверная дата или смена.');
    if (!isFinite(v) || v < 0 || v > 1000000) fail_('Введите проданные литры цифрами.');
    const st = readStations_().find(s => s.id === sid && !s.deleted);
    if (!st) fail_('Заправка не найдена. Выберите её заново.');
    let tank = null, cash = null, card = null;
    if (slot === '08') {
      tank = Number(p.tank); cash = Number(p.cash);
      if (p.tank === '' || p.tank == null || !isFinite(tank) || tank < 0) fail_('Введите остаток газа в бочке.');
      if (p.cash === '' || p.cash == null || !isFinite(cash) || cash < 0) fail_('Введите сумму денег в кассе.');
      card = Number(p.card);
      if (p.card === '' || p.card == null || !isFinite(card) || card < 0) fail_('Введите сумму безнала за сутки.');
    }
    const key = date + '_' + slot;
    const mine = readReports_().filter(r => r.sid === sid);
    const b = shiftBounds(mine, date, slot);
    const nm = s => (SLOTS.find(x => x.id === s) || {}).time;
    if (v < b.lo) fail_('Меньше, чем в сводке ' + nm(b.loSlot) + ' (' + fmt_(b.lo) + ' л). Литры считаются с 8:00 и не уменьшаются. Проверьте цифры.');
    if (b.hi != null && v > b.hi) fail_('Больше, чем в сводке ' + nm(b.hiSlot) + ' (' + fmt_(b.hi) + ' л). Проверьте цифры.');
    const existing = mine.find(r => r.key === key);
    const row = writeRow_('rep', existing ? existing.row : null, [key, sid, st.name, date, slot, v, tank, '', '', cash, new Date(), card]);
    let info = { name: st.name, slot, liters: v, late: !!props_().getProperty('m_' + date + '_' + slot) };
    if (slot === '08') {
      const d = load_();
      const r = d.reports.find(x => x.sid === sid && x.key === key);
      if (r && r.expected != null) {
        sheet_('rep').getRange(row, 8, 1, 2).setValues([[Math.round(r.expected), Math.round(r.tank - r.expected)]]);
        info.diff = r.tank - r.expected; info.tank = r.tank; info.cash = cash; info.card = card;
      }
    }
    return info;
  });
  if (res.late) {
    const s = SLOTS.find(x => x.id === res.slot);
    safe_(() => notifyOwner_('Сводка пришла с опозданием', res.name + ', ' + s.label.toLowerCase() + ': ' + fmt_(res.liters) + ' л' +
      (res.slot === '08' ? ' за сутки, касса ' + fmt_(res.cash || 0) + ' ₽, безнал ' + fmt_(res.card || 0) + ' ₽' : ' с 8:00')));
  }
  safe_(() => lowGasCheck_(String(p.sid)));
  return true;
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
  return true;
}

/* ---------- владелец ---------- */
function checkPin_(pin) {
  const cache = CacheService.getScriptCache();
  const bad = Number(cache.get('pin_bad') || 0);
  if (bad >= 10) fail_('Слишком много неверных попыток. Подождите 10 минут.');
  if (String(pin) !== ownerPin_()) { cache.put('pin_bad', String(bad + 1), 600); fail_('Неверный PIN.'); }
}
function ownerLogin(pin, dev, ua) {
  checkPin_(pin);
  dev = String(dev || '');
  let tok = null;
  if (/^[\w-]{8,64}$/.test(dev)) tok = withLock_(() => {
    const old = readDevices_().find(x => x.dev === dev);
    const t = old && old.owner && old.tok ? old.tok : Utilities.getUuid().replace(/-/g, '');
    writeRow_('dev', old ? old.row : null, [dev, t, '', '', 'Владелец', 'да', old ? new Date(old.at) : new Date(), new Date(), 'активен',
      old && old.sub ? JSON.stringify(old.sub) : '', String(ua || '').slice(0, 200)]);
    return t;
  });
  const s = settings_(); s.tok = tok; return s;
}
function addStation(pin, p) {
  checkPin_(pin);
  return withLock_(() => {
    const name = String(p && p.name || '').trim(), cap = num_(p && p.capacity), lvl = Number(p && p.level), sold = Number(p && p.sold) || 0;
    if (!name) fail_('Впишите название заправки.');
    if (!isFinite(lvl) || lvl < 0) fail_('Впишите, сколько газа в бочке.');
    const now = new Date();
    const price = num_(p && p.price);
    writeRow_('st', null, [Utilities.getUuid().slice(0, 8), name, cap, price > 0 ? price : '', lvl, now, sold, now, '']);
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
    const price = num_(p.price);
    sheet_('st').getRange(st.row, 2, 1, 3).setValues([[name, num_(p.capacity), price > 0 ? price : '']]);
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
    sheet_('st').getRange(st.row, 5, 1, 3).setValues([[lvl, new Date(), st.cumNow || 0]]);
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
  return !readDevices_().some(d => d.owner && pushable_(d));   // по умолчанию: почта, пока у владельца нет телефона с уведомлениями
}
function settings_() {
  const devs = readDevices_(), sts = readStations_().filter(s => !s.deleted);
  return {
    sheetUrl: book_().getUrl(),
    email: Session.getEffectiveUser().getEmail(), emailOn: emailOn_(),
    ownerPhones: devs.filter(d => d.owner && pushable_(d)).length,
    phones: sts.map(s => ({ name: s.name, count: devs.filter(d => !d.owner && d.sid === s.id && d.status !== 'отключён').length })),
    opPin: opPin_(),
    defaultPin: ownerPin_() === '1234',
    serverVersion: APP_VERSION, selfUpdate: typeof bootReload_ === 'function',
  };
}
function getSettings(pin) { checkPin_(pin); return settings_(); }
function setEmail(pin, on) { checkPin_(pin); props_().setProperty('EMAIL_ON', on ? 'yes' : 'no'); return settings_(); }

/* ---------- телефоны и уведомления ---------- */
function subscribe(r) {
  const sub = r.sub;
  if (!sub || !/^https:\/\//.test(String(sub.endpoint || ''))) fail_('Телефон не дал адрес для уведомлений.');
  return withLock_(() => {
    const d = deviceByKey_(r.dev, r.tok);
    if (!d) fail_('NEED_LOGIN');
    if (d.status === 'отключён') fail_('REVOKED');
    sheet_('dev').getRange(d.row, 10).setValue(JSON.stringify({ endpoint: sub.endpoint, keys: sub.keys || {} }));
    return { ok: true };
  });
}
function inbox(dev, tok) {
  const d = deviceByKey_(dev, tok);
  if (!d || d.status === 'отключён') return [];
  return withLock_(() => {
    const k = 'in_' + d.dev, raw = props_().getProperty(k);
    if (!raw) return [];
    props_().deleteProperty(k);
    let list = []; try { list = JSON.parse(raw); } catch (e) {}
    return list.filter(m => Date.now() - m.at < 24 * 3600000);
  });
}
function testPush(pin, dev) {
  checkPin_(pin);
  const d = readDevices_().find(x => x.dev === String(dev));
  if (!d || !d.sub) fail_('Этот телефон ещё не подключён к уведомлениям.');
  const r = push_([d], { title: 'Проверка', body: 'Уведомления работают.' });
  if (!r.sent) fail_('Не удалось отправить уведомление: ' + (r.errors[0] || 'неизвестная ошибка'));
  return 'Уведомление отправлено. Оно придёт через несколько секунд.';
}
function notifyOwner_(title, body, urgent) {
  const owners = readDevices_().filter(d => d.owner && pushable_(d));
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
  if (gone.length) safe_(() => gone.forEach(g => sheet_('dev').getRange(g.row, 10).setValue('')));   // телефон отписался от уведомлений
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
    const targets = devs.filter(x => !x.owner && x.sid === st.id && pushable_(x));
    if (!targets.length) return;
    push_(targets, {
      title: again ? 'Сводка ещё не отправлена' : 'Пора отправить сводку',
      body: s.label + ', ' + s.time + ' · ' + st.name + '. Нажмите, чтобы отправить.',
      urgent: true,
    });
  });
}
function summary_(date, s) {
  const d = load_(), sts = d.stations.filter(st => !st.deleted);
  if (!sts.length) return;
  const p = date.split('-').map(Number);
  const dueMs = Date.UTC(p[0], p[1] - 1, p[2], s.h) - tzOffsetMs_(new Date());
  let total = 0, cash = 0, card = 0, missing = 0, bigDiff = false, rev = 0, anyPrice = false; const low = [];
  const lines = sts.filter(st => !(st.createdAt > dueMs)).map(st => {
    const r = d.reports.find(x => x.sid === st.id && x.date === date && x.slot === s.id);
    const pct = st.capacity ? Math.round(st.tank / st.capacity * 100) : null;
    if (pct != null && pct < LOW_PCT) low.push(st.name);
    if (!r) { missing++; return st.name + ': НЕТ СВОДКИ'; }
    total += r.value;
    if (st.price > 0) { rev += r.value * st.price; anyPrice = true; }
    if (s.id !== '08') return st.name + ': ' + fmt_(r.value) + ' л с 8:00 · в бочке ≈ ' + fmt_(Math.max(0, Math.round(st.tank))) + ' л' + (pct != null ? ' (' + pct + '%)' : '');
    cash += r.cash || 0; card += r.card || 0;
    const diff = r.diff == null ? null : Math.round(r.diff);
    if (diff != null && Math.abs(diff) >= 50) bigDiff = true;
    return st.name + ': ' + fmt_(r.value) + ' л за сутки' + (st.price > 0 ? ' · выручка ' + fmt_(Math.round(r.value * st.price)) + ' (' + fmt_(r.card || 0) + ' безнал)' : ' · безнал ' + fmt_(r.card || 0) + ' ₽') +
      ' · касса ' + fmt_(r.cash || 0) + ' ₽ · бочка ' + fmt_(r.tank) + ' л' +
      (diff != null ? ' (расчёт ' + fmt_(Math.round(r.expected)) + ', ' + (diff > 0 ? '+' : '') + fmt_(diff) + ')' : '');
  });
  if (!lines.length) return;
  const title = s.id === '08'
    ? 'Выручка за ' + dm_(shiftDate_(date, -1)) + ': ' + (anyPrice ? fmt_(Math.round(rev)) + ' ₽ (' + fmt_(card) + ' безнал)' : 'безнал ' + fmt_(card) + ' ₽') + ' · ' + fmt_(total) + ' л' + (missing ? ' · нет сводок: ' + missing : '')
    : 'Сводка ' + s.time + ': ' + fmt_(total) + ' л с 8:00' + (anyPrice ? ' · ' + fmt_(Math.round(rev)) + ' ₽' : '') + (missing ? ' · нет сводок: ' + missing : '');
  let body = lines.join('\n');
  if (low.length) body += '\nМало газа: ' + low.join(', ');
  if (bigDiff) body += '\nЕсть расхождение по бочке — проверьте.';
  notifyOwner_(title, body, missing > 0 || bigDiff);
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
/* ===== Логика смен: одинаковая на сервере и в приложении =====
 * Рабочие сутки идут с 8:00 до 8:00 следующего дня.
 * 14:00, 18:00 и 22:00 — литры, проданные с 8:00 (после обнуления смены).
 * 8:00 — итог за прошедшие сутки (до обнуления), фактический остаток в бочке и деньги в кассе.
 */
var SLOT_RANK = { '14': 1, '18': 2, '22': 3, '08': 4 };
function addDays(key, n) {
  var p = key.split('-').map(Number);
  return new Date(Date.UTC(p[0], p[1] - 1, p[2] + n)).toISOString().slice(0, 10);
}
/** К каким рабочим суткам относится сводка: утренняя закрывает вчерашние. */
function workDayOf(r) { return r.slot === '08' ? addDays(r.date, -1) : r.date; }
/** Раскладывает сводки одной заправки по суткам и считает продажи за каждый период. */
function shiftDays(reps) {
  var days = {};
  reps.forEach(function (r) { var d = workDayOf(r); (days[d] = days[d] || []).push(r); });
  Object.keys(days).forEach(function (d) {
    var prev = 0;
    days[d].sort(function (a, b) { return SLOT_RANK[a.slot] - SLOT_RANK[b.slot]; }).forEach(function (r) {
      r.day = d; r.sold = r.value - prev; prev = r.value;
    });
  });
  return days;
}
function dayCum(list) { return list && list.length ? list[list.length - 1].value : 0; }
/** Остаток в бочке: последняя известная точка (утренний замер оператора или ввод владельца)
 *  плюс привозы после неё минус продажи после неё. */
function tankEstimate(st, reps, dels, days, dayOfMs, cutAt, upToDay) {
  var base = { level: st.baseLevel, at: st.baseAt, day: dayOfMs(st.baseAt), cum0: st.baseCum || 0, kind: 'owner' };
  reps.forEach(function (r) {
    if (r.slot === '08' && r.tank != null && r.at < cutAt && r.at > base.at) base = { level: r.tank, at: r.at, day: r.date, cum0: 0, kind: 'operator' };
  });
  var v = base.level;
  dels.forEach(function (d) { if (d.at > base.at && d.at <= cutAt) v += d.liters; });
  Object.keys(days).forEach(function (d) {
    if (d < base.day || d > upToDay) return;
    var c = dayCum(days[d]);
    v -= d === base.day ? Math.max(0, c - base.cum0) : c;
  });
  return { value: v, base: base };
}
/** Пересчитывает всё по заправке: продажи по периодам, остаток сейчас и расчётный остаток к каждому утру. */
function computeStation(st, reps, dels, dayOfMs) {
  var days = shiftDays(reps);
  var now = tankEstimate(st, reps, dels, days, dayOfMs, Infinity, '9999-12-31');
  st.tank = now.value;
  st.tankFrom = now.base.kind;
  st.tankAt = now.base.at;
  reps.forEach(function (r) {
    if (r.slot === '08') {
      r.expected = tankEstimate(st, reps, dels, days, dayOfMs, r.at, r.day).value;
      r.diff = r.tank != null ? r.tank - r.expected : null;
    }
  });
  return days;
}
/** Границы для нового значения литров в сводке: не меньше предыдущей за те же сутки, не больше следующей. */
function shiftBounds(reps, date, slot) {
  var probe = { date: date, slot: slot }, day = workDayOf(probe), rank = SLOT_RANK[slot];
  var lo = 0, hi = null, loSlot = null, hiSlot = null;
  reps.forEach(function (r) {
    if (workDayOf(r) !== day || r.slot === slot) return;
    if (SLOT_RANK[r.slot] < rank && r.value >= lo) { lo = r.value; loSlot = r.slot; }
    if (SLOT_RANK[r.slot] > rank && (hi == null || r.value < hi)) { hi = r.value; hiSlot = r.slot; }
  });
  return { lo: lo, hi: hi, loSlot: loSlot, hiSlot: hiSlot, day: day };
}
