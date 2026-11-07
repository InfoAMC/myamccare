// AMC Care Group — portal de agentes y Daily Report.
// Toda la validación de acceso y de datos ocurre en el Apps Script.
var API_URL = 'https://script.google.com/macros/s/AKfycbyNLaR1nbwCRVKaeNrcNqxiEmLfOxHi6forzTz-Xu8nPaV1pLORPOSpVG4_p4TiCT55/exec';   // misma URL /exec que los formularios (Formularios-Web.gs)
var LOGOS = 'logos/';
var FILL_LOGOS = ['ssa-lis.webp', 'zoom.webp', 'whatsapp-web.webp', 'wellcare.webp'];
var PAGE_SIZE = 15;
var XLSX_CDN = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';

var I18N = {
  es: {
    logout: 'Salir', loginTitle: 'Portal de agentes',
    loginHint: 'Escriba su correo y le enviaremos un código para entrar.',
    codeHint: function (e) { return 'Si ' + e + ' tiene acceso, le llegará un código en unos minutos. Revise también el correo no deseado.'; },
    emailLabel: 'Correo', sendCode: 'Enviarme el código', codeLabel: 'Código de 6 dígitos', enter: 'Entrar', useOther: 'Usar otro correo',
    badEmail: 'Escriba un correo válido.', badCode: 'El código no es correcto o ya venció. Pida uno nuevo si hace falta.',
    noApi: 'El portal todavía no está conectado.', netErr: 'No pudimos conectar. Revise su internet e intente de nuevo.', sessionOut: 'Su sesión terminó. Entre de nuevo.',
    hello: function (n) { return 'Hola, ' + n; }, portalSub: 'Sus herramientas y portales, en un solo lugar.',
    search: 'Buscar enlace…', noLinks: 'No encontramos ese enlace.', favorite: 'Marcar como favorito',
    daily: 'Reporte Diario de Ventas', backPortal: '← Volver al portal',
    reportTitle: 'Reporte Diario de Ventas', reportSub: 'Registre sus ventas del día. No escriba datos de clientes.',
    salesToday: 'ventas hoy (equipo)', tabReg: 'Registrar venta', tabHist: 'Mi historial', tabAdmin: 'Administración',
    howMany: '¿Cuántas ventas va a registrar?', saleDate: 'Fecha de esta venta', dateHint: 'Cámbiela si la venta fue otro día.',
    enrollHelpTitle: '¿Qué significa cada tipo de inscripción?',
    enrollHelp: ['New to Medicare: nuevo en Medicare o en su inscripción inicial (IEP).', 'Same Carrier Plan Change: su cliente cambia de plan con la misma aseguradora.', 'Same Carrier Transfer: cliente nuevo para usted que sigue con su aseguradora, pero cambia de plan.', 'New Carrier Switch: su cliente pasa a otra aseguradora.', 'New Client & Carrier: cliente nuevo que viene de otra aseguradora.'],
    saleN: function (n) { return 'Venta ' + n; }, carrier: 'Aseguradora', plan: 'Tipo de plan', enrollment: 'Inscripción', lead: 'Origen del lead', select: 'Elija una opción',
    completeHint: 'Complete todos los campos para activar el botón.',
    submit: function (n) { return n === 1 ? 'Registrar 1 venta' : 'Registrar ' + n + ' ventas'; },
    success: function (n) { return '✓ ' + (n === 1 ? '1 venta registrada.' : n + ' ventas registradas.'); },
    saveErr: 'No se pudo guardar. Intente de nuevo.',
    allCarriers: 'Todas las aseguradoras', allPlans: 'Todos los planes', allAgents: 'Todos los agentes',
    date: 'Fecha', agent: 'Agente', registered: 'Registrado', none: 'No hay ventas para mostrar.',
    histTotal: function (n) { return 'Total de los últimos 30 días: ' + n; },
    mToday: 'Ventas de hoy', mMonth: 'Ventas del mes', mAgents: 'Agentes activos', mTop: 'Aseguradora líder del mes',
    chartTitle: 'Ventas por plan o agente', chartHint: 'Filtre por aseguradora, plan o fechas. Muestra el año actual si no elige fechas.',
    groupPlan: 'Agrupar por plan', groupAgent: 'Agrupar por agente', downloadCsv: 'Descargar CSV', downloadXlsx: 'Descargar Excel', downloadChart: 'Descargar gráfica',
    fullHist: 'Historial completo', page: function (a, b) { return 'Página ' + a + ' de ' + b; }, prev: 'Anterior', next: 'Siguiente'
  },
  en: {
    logout: 'Sign out', loginTitle: 'Agent portal',
    loginHint: 'Enter your email and we will send you a sign-in code.',
    codeHint: function (e) { return 'If ' + e + ' has access, a code will arrive within a few minutes. Please check your spam folder too.'; },
    emailLabel: 'Email', sendCode: 'Send me the code', codeLabel: '6-digit code', enter: 'Sign in', useOther: 'Use another email',
    badEmail: 'Please enter a valid email.', badCode: 'That code is wrong or has expired. Request a new one if needed.',
    noApi: 'The portal is not connected yet.', netErr: 'We could not connect. Check your internet and try again.', sessionOut: 'Your session ended. Please sign in again.',
    hello: function (n) { return 'Hello, ' + n; }, portalSub: 'Your tools and portals, all in one place.',
    search: 'Search links…', noLinks: 'We could not find that link.', favorite: 'Mark as favorite',
    daily: 'Daily Sales Report', backPortal: '← Back to portal',
    reportTitle: 'Daily Sales Report', reportSub: 'Log your sales for the day. Do not enter client information.',
    salesToday: 'sales today (team)', tabReg: 'Register sale', tabHist: 'My history', tabAdmin: 'Administration',
    howMany: 'How many sales are you registering?', saleDate: 'Date of this sale', dateHint: 'Change it if the sale was on another day.',
    enrollHelpTitle: 'What does each enrollment type mean?',
    enrollHelp: ['New to Medicare: brand new to Medicare or in their Initial Enrollment Period (IEP).', 'Same Carrier Plan Change: your existing client changes plans with the same carrier.', 'Same Carrier Transfer: new client to you, staying with their carrier but changing plans.', 'New Carrier Switch: your existing client moves to a new carrier.', 'New Client & Carrier: brand new client coming from another carrier.'],
    saleN: function (n) { return 'Sale ' + n; }, carrier: 'Carrier', plan: 'Plan type', enrollment: 'Enrollment', lead: 'Lead source', select: 'Select an option',
    completeHint: 'Complete every field to enable the button.',
    submit: function (n) { return n === 1 ? 'Register 1 sale' : 'Register ' + n + ' sales'; },
    success: function (n) { return '✓ ' + (n === 1 ? '1 sale registered.' : n + ' sales registered.'); },
    saveErr: 'We could not save it. Please try again.',
    allCarriers: 'All carriers', allPlans: 'All plans', allAgents: 'All agents',
    date: 'Date', agent: 'Agent', registered: 'Registered', none: 'No sales to show.',
    histTotal: function (n) { return 'Total for the last 30 days: ' + n; },
    mToday: 'Sales today', mMonth: 'Sales this month', mAgents: 'Active agents', mTop: 'Top carrier this month',
    chartTitle: 'Sales by plan or agent', chartHint: 'Filter by carrier, plan or dates. It shows the current year if you pick no dates.',
    groupPlan: 'Group by plan', groupAgent: 'Group by agent', downloadCsv: 'Download CSV', downloadXlsx: 'Download Excel', downloadChart: 'Download chart',
    fullHist: 'Full history', page: function (a, b) { return 'Page ' + a + ' of ' + b; }, prev: 'Previous', next: 'Next'
  }
};

var lang = store('get', 'amc-lang') === 'en' ? 'en' : 'es';
var state = { token: null, me: null, links: [], options: null, today: '', count: 1, sales: [], history: [], admin: null, adminPage: 1, emailSent: '' };

function $(id) { return document.getElementById(id); }
function t(k) { return I18N[lang][k]; }
function esc(v) { return String(v == null ? '' : v).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
function store(op, k, v) {
  try {
    if (op === 'get') return localStorage.getItem(k);
    if (op === 'set') localStorage.setItem(k, v);
    if (op === 'del') localStorage.removeItem(k);
  } catch (e) {}
  return null;
}

function api(action, data) {
  var body = Object.assign({ action: action, token: state.token }, data || {});
  if (window.AMC_DEMO_API) return window.AMC_DEMO_API(body);
  if (!API_URL) return Promise.reject(new Error('noApi'));
  return fetch(API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(body)
  }).then(function (r) { return r.json(); });
}

// Llama al script y devuelve al inicio de sesión si la sesión ya no vale.
function authed(action, data) {
  return api(action, data).then(function (res) {
    if (res.error === 'session') {
      endSession();
      showMsg('loginMsg', t('sessionOut'), true);
      throw new Error('session');
    }
    return res;
  });
}

function showView(id) {
  ['vLoading', 'vLogin', 'vPortal', 'vReport'].forEach(function (v) { $(v).hidden = v !== id; });
  var inside = id === 'vPortal' || id === 'vReport';
  $('who').hidden = !inside;
  $('logoutBtn').hidden = !inside;
}

function showMsg(id, text, isErr) {
  var m = $(id);
  m.textContent = text;
  m.classList.toggle('err', !!isErr);
  m.hidden = false;
}

function applyLang() {
  document.documentElement.lang = lang;
  document.querySelectorAll('[data-t]').forEach(function (el) { el.textContent = t(el.dataset.t); });
  document.querySelectorAll('.lang button').forEach(function (b) {
    var on = b.dataset.lang === lang;
    b.classList.toggle('active', on);
    b.setAttribute('aria-pressed', on);
  });
  $('search').setAttribute('placeholder', t('search'));
  $('loginHint').textContent = state.emailSent ? t('codeHint')(state.emailSent) : t('loginHint');
  $('enrollHelp').innerHTML = t('enrollHelp').map(function (l) { return '<li>' + esc(l) + '</li>'; }).join('');
  if (state.me) {
    $('hello').textContent = t('hello')(state.me.name.split(' ')[0]);
    renderLinks();
    renderSales();
    fillFilters();
    renderHistory();
    renderAdmin();
  }
}

// ---------- entrada ----------

function endSession() {
  state.token = null;
  state.me = null;
  state.admin = null;
  state.emailSent = '';
  store('del', 'amc-session');
  $('codeForm').hidden = true;
  $('emailForm').hidden = false;
  $('loginMsg').hidden = true;
  $('code').value = '';
  $('loginHint').textContent = t('loginHint');
  showView('vLogin');
}

$('emailForm').addEventListener('submit', function (e) {
  e.preventDefault();
  var email = $('email').value.trim().toLowerCase();
  $('loginMsg').hidden = true;
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return showMsg('loginMsg', t('badEmail'), true);
  var btn = e.target.querySelector('button');
  btn.disabled = true;
  api('requestCode', { email: email }).then(function (res) {
    if (!res || !res.ok) throw new Error('server');
    state.emailSent = email;
    $('loginHint').textContent = t('codeHint')(email);
    $('emailForm').hidden = true;
    $('codeForm').hidden = false;
    $('code').focus();
  }).catch(function (err) {
    showMsg('loginMsg', err.message === 'noApi' ? t('noApi') : t('netErr'), true);
  }).then(function () { btn.disabled = false; });
});

$('codeForm').addEventListener('submit', function (e) {
  e.preventDefault();
  var btn = e.target.querySelector('.primary');
  $('loginMsg').hidden = true;
  btn.disabled = true;
  api('verifyCode', { email: state.emailSent, code: $('code').value }).then(function (res) {
    if (!res.ok) return showMsg('loginMsg', t('badCode'), true);
    state.token = res.token;
    store('set', 'amc-session', JSON.stringify({ t: res.token, x: res.expires }));
    return start();
  }).catch(function () {
    showMsg('loginMsg', t('netErr'), true);
  }).then(function () { btn.disabled = false; });
});

$('backBtn').addEventListener('click', function () {
  state.emailSent = '';
  $('codeForm').hidden = true;
  $('emailForm').hidden = false;
  $('loginMsg').hidden = true;
  $('code').value = '';
  $('loginHint').textContent = t('loginHint');
});

$('logoutBtn').addEventListener('click', function () {
  var token = state.token;
  api('logout', { token: token }).catch(function () {});
  endSession();
});

function start() {
  showView('vLoading');
  return authed('boot').then(function (res) {
    if (!res.ok) throw new Error('boot');
    state.me = res.me;
    state.links = res.links;
    state.options = res.options;
    state.today = res.today;
    $('todayN').textContent = res.todayCount;
    $('whoName').textContent = res.me.name;
    $('whoEmail').textContent = res.me.email;
    $('tabAdmin').hidden = !res.me.admin;
    state.count = 1;
    state.sales = [newSale()];
    applyLang();
    route();
  }).catch(function (err) {
    if (err.message === 'session') return;
    endSession();
    showMsg('loginMsg', t('netErr'), true);
  });
}

// ---------- portal ----------

function favorites() {
  try { return JSON.parse(store('get', 'amc-fav-' + state.me.email) || '[]'); } catch (e) { return []; }
}

function renderLinks() {
  var favs = favorites();
  var q = $('search').value.trim().toLowerCase();
  var list = state.links.slice().sort(function (a, b) {
    var fa = favs.indexOf(a.name) > -1, fb = favs.indexOf(b.name) > -1;
    if (fa !== fb) return fa ? -1 : 1;
    return a.name.localeCompare(b.name);
  });
  var html = '';
  if (!q || t('daily').toLowerCase().indexOf(q) > -1) {
    html += '<div class="lk"><button class="lk-main" type="button" id="openReport">' +
      '<span class="lk-circle"><img src="' + LOGOS + 'reporte-diario.svg" alt=""></span>' +
      '<span class="lk-name">' + esc(t('daily')) + '</span></button></div>';
  }
  list.forEach(function (l) {
    if (q && l.name.toLowerCase().indexOf(q) === -1) return;
    var fav = favs.indexOf(l.name) > -1;
    var fill = FILL_LOGOS.indexOf(l.logo) > -1 ? ' fill' : '';
    html += '<div class="lk"><a class="lk-main" href="' + esc(l.href) + '" target="_blank" rel="noopener noreferrer">' +
      '<span class="lk-circle' + fill + '"><img src="' + LOGOS + esc(l.logo) + '" alt="" loading="lazy"></span>' +
      '<span class="lk-name">' + esc(l.name) + '</span></a>' +
      '<button class="lk-star" type="button" data-fav="' + esc(l.name) + '" aria-pressed="' + fav + '" aria-label="' + esc(t('favorite') + ': ' + l.name) + '" title="' + esc(t('favorite')) + '">★</button></div>';
  });
  $('links').innerHTML = html;
  $('linksEmpty').hidden = html !== '';
}

$('search').addEventListener('input', renderLinks);
$('links').addEventListener('click', function (e) {
  var star = e.target.closest('[data-fav]');
  if (star) {
    var favs = favorites(), i = favs.indexOf(star.dataset.fav);
    if (i > -1) favs.splice(i, 1); else favs.push(star.dataset.fav);
    store('set', 'amc-fav-' + state.me.email, JSON.stringify(favs));
    renderLinks();
    return;
  }
  if (e.target.closest('#openReport')) location.hash = '#reporte';
});

function route() {
  if (!state.me) return;
  if (location.hash === '#reporte') {
    showView('vReport');
    setTab(state.tab || 'reg');
  } else {
    showView('vPortal');
  }
  window.scrollTo(0, 0);
}
window.addEventListener('hashchange', route);
$('homeBtn').addEventListener('click', function () { if (state.me) location.hash = ''; });
$('backHome').addEventListener('click', function () { location.hash = ''; });

// ---------- pestañas ----------

function setTab(tab) {
  if (tab === 'admin' && !state.me.admin) tab = 'reg';
  state.tab = tab;
  document.querySelectorAll('.tabs button').forEach(function (b) { b.setAttribute('aria-selected', b.dataset.tab === tab); });
  ['reg', 'hist', 'admin'].forEach(function (p) { $('p-' + p).hidden = p !== tab; });
  if (tab === 'hist') loadHistory();
  if (tab === 'admin') loadAdmin();
}
document.querySelector('.tabs').addEventListener('click', function (e) {
  var b = e.target.closest('[data-tab]');
  if (b) setTab(b.dataset.tab);
});

// ---------- registrar venta ----------

function newSale() {
  return { sale_date: state.today };
}

function optionTags(list, selected) {
  return '<option value="">' + esc(t('select')) + '</option>' + list.map(function (o) {
    return '<option value="' + esc(o) + '"' + (o === selected ? ' selected' : '') + '>' + esc(o) + '</option>';
  }).join('');
}

function renderSales() {
  if (!state.options) return;
  $('count').textContent = state.count;
  var o = state.options;
  $('saleBlocks').innerHTML = state.sales.map(function (s, i) {
    var cars = o.carriers.map(function (c) {
      var fill = FILL_LOGOS.indexOf(c.logo) > -1 ? ' style="width:100%;height:100%;object-fit:cover"' : '';
      return '<button class="car" type="button" data-i="' + i + '" data-car="' + esc(c.name) + '" aria-pressed="' + (s.carrier === c.name) + '">' +
        '<span class="car-circle"><img src="' + LOGOS + esc(c.logo) + '" alt=""' + fill + '></span><span>' + esc(c.name) + '</span></button>';
    }).join('');
    return '<div class="sale"><h3>' + esc(t('saleN')(i + 1)) + '</h3><div class="sale-date"><label for="da' + i + '">' + esc(t('saleDate')) + '</label><input id="da' + i + '" type="date" max="' + esc(state.today) + '" data-i="' + i + '" data-f="sale_date" value="' + esc(s.sale_date || '') + '"><p class="hint">' + esc(t('dateHint')) + '</p></div><label>' + esc(t('carrier')) + '</label><div class="car-grid">' + cars + '</div>' +
      '<div class="sale-fields">' +
      '<div><label for="pl' + i + '">' + esc(t('plan')) + '</label><select id="pl' + i + '" data-i="' + i + '" data-f="plan_type">' + optionTags(o.plan_types, s.plan_type) + '</select></div>' +
      '<div><label for="en' + i + '">' + esc(t('enrollment')) + '</label><select id="en' + i + '" data-i="' + i + '" data-f="enrollment_type">' + optionTags(o.enrollment_types, s.enrollment_type) + '</select></div>' +
      '<div><label for="le' + i + '">' + esc(t('lead')) + '</label><select id="le' + i + '" data-i="' + i + '" data-f="lead_source">' + optionTags(o.lead_sources, s.lead_source) + '</select></div>' +
      '</div></div>';
  }).join('');
  updateSubmit();
}

function updateSubmit() {
  var done = state.sales.every(function (s) { return s.sale_date && s.carrier && s.plan_type && s.enrollment_type && s.lead_source; });
  var b = $('submitSales');
  b.disabled = !done;
  b.textContent = t('submit')(state.count);
  $('completeHint').hidden = done;
}

function setCount(n) {
  n = Math.max(1, Math.min(15, n));
  state.count = n;
  while (state.sales.length < n) state.sales.push(newSale());
  state.sales.length = n;
  renderSales();
}
$('minus').addEventListener('click', function () { setCount(state.count - 1); });
$('plus').addEventListener('click', function () { setCount(state.count + 1); });
$('saleBlocks').addEventListener('click', function (e) {
  var b = e.target.closest('[data-car]');
  if (!b) return;
  state.sales[+b.dataset.i].carrier = b.dataset.car;
  document.querySelectorAll('.car[data-i="' + b.dataset.i + '"]').forEach(function (c) { c.setAttribute('aria-pressed', c === b); });
  updateSubmit();
});
$('saleBlocks').addEventListener('change', function (e) {
  if (!e.target.dataset.f) return;
  state.sales[+e.target.dataset.i][e.target.dataset.f] = e.target.value;
  updateSubmit();
});
$('submitSales').addEventListener('click', function () {
  var b = $('submitSales');
  b.disabled = true;
  $('regMsg').hidden = true;
  authed('submitSales', { sales: state.sales }).then(function (res) {
    if (!res.ok) return showMsg('regMsg', t('saveErr'), true);
    showMsg('regMsg', t('success')(res.count));
    $('todayN').textContent = res.todayCount;
    state.admin = null;
    state.count = 1;
    state.sales = [newSale()];
    renderSales();
  }).catch(function (err) {
    if (err.message !== 'session') showMsg('regMsg', t('saveErr'), true);
    updateSubmit();
  });
});

// ---------- tablas ----------

function table(rows, withAgent) {
  if (!rows.length) return '<p class="none">' + esc(t('none')) + '</p>';
  var head = [t('date')].concat(withAgent ? [t('agent')] : [], [t('carrier'), t('plan'), t('enrollment'), t('lead'), t('registered')]);
  return '<table><thead><tr>' + head.map(function (h) { return '<th>' + esc(h) + '</th>'; }).join('') + '</tr></thead><tbody>' +
    rows.map(function (r) {
      var cells = [r.sale_date].concat(withAgent ? [r.agent] : [], [r.carrier, r.plan_type, r.enrollment_type, r.lead_source, r.registered]);
      return '<tr>' + cells.map(function (c) { return '<td>' + esc(c) + '</td>'; }).join('') + '</tr>';
    }).join('') + '</tbody></table>';
}

// ---------- mi historial ----------

function loadHistory() {
  authed('myHistory').then(function (res) {
    if (!res.ok) return;
    state.history = res.rows;
    renderHistory();
  }).catch(function () {});
}

function renderHistory() {
  var c = $('histCarrier').value;
  var rows = state.history.filter(function (r) { return !c || r.carrier === c; });
  $('histTable').innerHTML = table(rows);
  $('histTotal').textContent = t('histTotal')(rows.length);
}
$('histCarrier').addEventListener('change', renderHistory);

function fillSelect(el, first, list, keep) {
  var v = keep ? el.value : '';
  el.innerHTML = '<option value="">' + esc(first) + '</option>' + list.map(function (o) { return '<option value="' + esc(o) + '">' + esc(o) + '</option>'; }).join('');
  el.value = v;
}

function fillFilters() {
  var names = state.options.carriers.map(function (c) { return c.name; });
  fillSelect($('histCarrier'), t('allCarriers'), names, true);
  fillSelect($('fCarrier'), t('allCarriers'), names, true);
  fillSelect($('fPlan'), t('allPlans'), state.options.plan_types, true);
  var g = $('fGroup').value || 'plan_type';
  $('fGroup').innerHTML = '<option value="plan_type">' + esc(t('groupPlan')) + '</option><option value="agent">' + esc(t('groupAgent')) + '</option>';
  $('fGroup').value = g;
  fillSelect($('hAgent'), t('allAgents'), state.admin ? state.admin.agents : [], true);
}

// ---------- administración ----------

function loadAdmin() {
  if (state.admin) return renderAdmin();
  authed('adminData').then(function (res) {
    if (!res.ok) return;
    state.admin = res;
    state.adminPage = 1;
    fillFilters();
    renderAdmin();
  }).catch(function () {});
}

function chartRows() {
  var c = $('fCarrier').value, p = $('fPlan').value, a = $('fFrom').value, z = $('fTo').value;
  return state.admin.rows.filter(function (r) {
    return (!c || r.carrier === c) && (!p || r.plan_type === p) && (!a || r.sale_date >= a) && (!z || r.sale_date <= z);
  });
}

function fullRows() {
  var a = $('hFrom').value, z = $('hTo').value, g = $('hAgent').value;
  return state.admin.rows.filter(function (r) {
    return (!a || r.sale_date >= a) && (!z || r.sale_date <= z) && (!g || r.agent === g);
  });
}

function chartGroups() {
  var key = $('fGroup').value || 'plan_type', counts = {};
  chartRows().forEach(function (r) { counts[r[key]] = (counts[r[key]] || 0) + 1; });
  return Object.keys(counts).map(function (k) { return { label: k, n: counts[k] }; }).sort(function (a, b) { return b.n - a.n; });
}

function renderAdmin() {
  var d = state.admin;
  if (!d) return;
  $('metrics').innerHTML = [[d.metrics.today, t('mToday')], [d.metrics.month, t('mMonth')], [d.metrics.agents, t('mAgents')], [d.metrics.top, t('mTop')]].map(function (m) {
    return '<div class="metric"><b>' + esc(m[0]) + '</b><span>' + esc(m[1]) + '</span></div>';
  }).join('');
  var groups = chartGroups();
  var max = groups.reduce(function (m, g) { return Math.max(m, g.n); }, 0);
  $('chart').innerHTML = groups.length ? groups.map(function (g) {
    return '<div class="bar"><span>' + esc(g.label) + '</span><div><i style="width:' + Math.round(g.n / max * 100) + '%"></i></div><b>' + g.n + '</b></div>';
  }).join('') : '<p class="none">' + esc(t('none')) + '</p>';
  var rows = fullRows();
  var pages = Math.max(1, Math.ceil(rows.length / PAGE_SIZE));
  state.adminPage = Math.min(state.adminPage, pages);
  var start = (state.adminPage - 1) * PAGE_SIZE;
  $('fullTable').innerHTML = table(rows.slice(start, start + PAGE_SIZE), true);
  $('pager').innerHTML = pages > 1 ? '<button class="ghost" type="button" data-p="-1"' + (state.adminPage === 1 ? ' disabled' : '') + '>' + esc(t('prev')) + '</button><span>' + esc(t('page')(state.adminPage, pages)) + '</span><button class="ghost" type="button" data-p="1"' + (state.adminPage === pages ? ' disabled' : '') + '>' + esc(t('next')) + '</button>' : '';
}

['fCarrier', 'fPlan', 'fFrom', 'fTo', 'fGroup'].forEach(function (id) { $(id).addEventListener('change', renderAdmin); });
['hFrom', 'hTo', 'hAgent'].forEach(function (id) {
  $(id).addEventListener('change', function () { state.adminPage = 1; renderAdmin(); });
});
$('pager').addEventListener('click', function (e) {
  var b = e.target.closest('[data-p]');
  if (!b) return;
  state.adminPage += +b.dataset.p;
  renderAdmin();
});

// ---------- descargas ----------

function exportMatrix(rows) {
  return [['Date', 'Agent', 'Email', 'Carrier', 'Plan', 'Enrollment', 'Lead source', 'Registered']].concat(rows.map(function (r) {
    return [r.sale_date, r.agent, r.email, r.carrier, r.plan_type, r.enrollment_type, r.lead_source, r.registered];
  }));
}

function saveBlob(blob, name) {
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
}

function downloadCsv(rows, name) {
  var csv = exportMatrix(rows).map(function (line) {
    return line.map(function (c) {
      var v = String(c == null ? '' : c);
      if (/^[=+\-@]/.test(v)) v = "'" + v;
      return '"' + v.replace(/"/g, '""') + '"';
    }).join(',');
  }).join('\r\n');
  saveBlob(new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' }), name + '.csv');
}

function downloadXlsx(rows, name) {
  var go = function () {
    var wb = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.aoa_to_sheet(exportMatrix(rows)), 'Sales');
    window.XLSX.writeFile(wb, name + '.xlsx');
  };
  if (window.XLSX) return go();
  var s = document.createElement('script');
  s.src = XLSX_CDN;
  s.onload = go;
  s.onerror = function () { downloadCsv(rows, name); };
  document.head.appendChild(s);
}

function downloadChart() {
  var groups = chartGroups();
  var c = document.createElement('canvas');
  c.width = 900;
  c.height = 110 + Math.max(groups.length, 1) * 44;
  var x = c.getContext('2d');
  var max = groups.reduce(function (m, g) { return Math.max(m, g.n); }, 1);
  x.fillStyle = '#fff';
  x.fillRect(0, 0, c.width, c.height);
  x.fillStyle = '#000E66';
  x.font = 'bold 26px Georgia';
  x.fillText('AMC Care Group · ' + t('chartTitle'), 28, 50);
  x.font = '16px Arial';
  groups.forEach(function (g, i) {
    var y = 88 + i * 44;
    x.fillStyle = '#121633';
    x.fillText(g.label, 28, y + 20);
    x.fillStyle = '#C9A250';
    x.fillRect(250, y, Math.max(6, g.n / max * 580), 28);
    x.fillStyle = '#000E66';
    x.fillText(String(g.n), 250 + Math.max(6, g.n / max * 580) + 8, y + 20);
  });
  c.toBlob(function (b) { saveBlob(b, 'amc-sales-chart.png'); });
}

$('chartCsv').addEventListener('click', function () { downloadCsv(chartRows(), 'amc-sales-chart'); });
$('chartXlsx').addEventListener('click', function () { downloadXlsx(chartRows(), 'amc-sales-chart'); });
$('chartPng').addEventListener('click', downloadChart);
$('fullCsv').addEventListener('click', function () { downloadCsv(fullRows(), 'amc-sales'); });
$('fullXlsx').addEventListener('click', function () { downloadXlsx(fullRows(), 'amc-sales'); });

// ---------- inicio ----------

document.querySelectorAll('.lang button').forEach(function (b) {
  b.addEventListener('click', function () {
    lang = b.dataset.lang;
    store('set', 'amc-lang', lang);
    applyLang();
  });
});

applyLang();
(function () {
  var saved = null;
  try { saved = JSON.parse(store('get', 'amc-session')); } catch (e) {}
  if (saved && saved.t && saved.x > Date.now()) {
    state.token = saved.t;
    start();
  } else {
    store('del', 'amc-session');
    showView('vLogin');
  }
})();
