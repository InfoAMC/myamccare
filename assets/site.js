// AMC Care Group — comportamiento del sitio público.
// Datos que se completan en un solo lugar cuando estén listos:
var CONFIG = {
  phone: '+18623474364',     // número principal al público
  phoneDisplay: '(862) 347-4364',
  formEndpoint: 'https://script.google.com/macros/s/AKfycbyNLaR1nbwCRVKaeNrcNqxiEmLfOxHi6forzTz-Xu8nPaV1pLORPOSpVG4_p4TiCT55/exec',
  facebook: 'https://www.facebook.com/AMCCareGroup/',
  instagram: 'https://www.instagram.com/amccg/',
  whatsapp: ''               // solo dígitos con código de país, por ejemplo 18623474364. Vacío = oculto
};

var LANG = 'es';
var T = {
  es: {
    open: 'Abierto ahora',
    iep: function (a, b) { return 'De ' + a + ' a ' + b; },
    sent: '¡Gracias! Recibimos su solicitud. Un agente con licencia le llamará pronto, normalmente el mismo día hábil.',
    joinSent: '¡Gracias! Recibimos sus datos. Le escribiremos pronto.',
    missing: 'Revise los campos marcados en rojo.',
    badPhone: 'Escriba un teléfono de 10 dígitos, por ejemplo (305) 555-0123.',
    noEndpoint: 'El formulario todavía no está conectado. Llámenos al (862) 347-4364.',
    fail: 'No pudimos enviar su solicitud. Intente de nuevo o llámenos.',
    sending: 'Enviando…',
    menu: 'Menú', close: 'Cerrar'
  },
  en: {
    open: 'Open now',
    iep: function (a, b) { return a + ' to ' + b; },
    sent: 'Thank you! We received your request. A licensed agent will call you soon, usually the same business day.',
    joinSent: 'Thank you! We received your information. We will contact you soon.',
    missing: 'Please check the fields marked in red.',
    badPhone: 'Enter a 10-digit phone number, for example (305) 555-0123.',
    noEndpoint: 'The form is not connected yet. Please call us at (862) 347-4364.',
    fail: 'We could not send your request. Please try again or call us.',
    sending: 'Sending…',
    menu: 'Menu', close: 'Close'
  }
};

function updateTextLabel() {
  var b = document.getElementById('textSize');
  if (!b) return;
  var big = document.documentElement.classList.contains('big-text');
  var label = LANG === 'en' ? (big ? 'Make text smaller' : 'Make text larger') : (big ? 'Reducir la letra' : 'Agrandar la letra');
  b.setAttribute('aria-label', label);
  b.title = label;
}

function setLang(l) {
  LANG = l;
  document.documentElement.lang = l;
  document.querySelectorAll('[data-en]').forEach(function (el) {
    if (el.dataset.es === undefined) el.dataset.es = el.innerHTML;
    el.innerHTML = l === 'en' ? el.dataset.en : el.dataset.es;
  });
  document.querySelectorAll('[data-en-ph]').forEach(function (el) {
    if (el.dataset.esPh === undefined) el.dataset.esPh = el.getAttribute('placeholder') || '';
    el.setAttribute('placeholder', l === 'en' ? el.dataset.enPh : el.dataset.esPh);
  });
  document.querySelectorAll('[data-en-alt]').forEach(function (el) {
    if (el.dataset.esAlt === undefined) el.dataset.esAlt = el.getAttribute('alt') || '';
    el.setAttribute('alt', l === 'en' ? el.dataset.enAlt : el.dataset.esAlt);
  });
  document.querySelectorAll('.lang button').forEach(function (b) {
    var on = b.dataset.lang === l;
    b.classList.toggle('active', on);
    b.setAttribute('aria-pressed', on);
  });
  document.querySelectorAll('.open-badge').forEach(function (b) { b.textContent = T[l].open; });
  document.querySelectorAll('.form-msg').forEach(function (m) { m.hidden = true; });
  var lf = document.getElementById('f-lang');
  if (lf) lf.value = l;
  updateIep();
  updateTextLabel();
  try { localStorage.setItem('amc-lang', l); } catch (e) {}
}

// Periodo de inscripción inicial: 3 meses antes del mes en que cumple 65, ese mes y 3 después.
// Si nació el día 1, todo se adelanta un mes. La fecha se calcula aquí y no se envía a ningún lado.
function updateIep(force) {
  var input = document.getElementById('bday65');
  var out = document.getElementById('iepOut');
  if (!input || !out) return;
  if (force !== true && !out.textContent) return;
  if (!input.value) { out.textContent = ''; return; }
  var p = input.value.split('-');
  var y = +p[0] + 65, m = +p[1] - 1 - (+p[2] === 1 ? 1 : 0);
  var fmt = new Intl.DateTimeFormat(LANG === 'en' ? 'en-US' : 'es-US', { month: 'long', year: 'numeric' });
  out.textContent = T[LANG].iep(fmt.format(new Date(y, m - 3, 1)), fmt.format(new Date(y, m + 3, 1)));
}

function markOpenPeriods() {
  var now = new Date();
  var md = (now.getMonth() + 1) * 100 + now.getDate(); // 1015 = 15 oct
  var open = { aep: md >= 1015 && md <= 1207, oep: md >= 101 && md <= 331 };
  document.querySelectorAll('.timeline li').forEach(function (li) {
    if (!open[li.dataset.period]) return;
    li.classList.add('is-open');
    var b = document.createElement('span');
    b.className = 'open-badge';
    b.textContent = T[LANG].open;
    li.querySelector('.when').insertAdjacentElement('afterend', b);
  });
}

function showMsg(form, text, isError) {
  var msg = form.querySelector('.form-msg');
  msg.textContent = text;
  msg.classList.toggle('error', !!isError);
  msg.hidden = false;
}

function validate(form) {
  var ok = true;
  form.querySelectorAll('[required]').forEach(function (el) {
    var bad = el.type === 'checkbox' ? !el.checked : !el.value.trim() || !el.checkValidity();
    var box = el.closest('.field') || el.closest('.check');
    if (box) box.classList.toggle('invalid', bad);
    if (bad) ok = false;
  });
  return ok;
}

function send(form, okText) {
  if (form.querySelector('.hp').value) return; // bot
  form.querySelectorAll('input[type=tel]').forEach(function (el) { el.dispatchEvent(new Event('change')); });
  if (!validate(form)) {
    var tel = form.querySelector('input[type=tel]');
    showMsg(form, tel && tel.validationMessage === T[LANG].badPhone ? T[LANG].badPhone : T[LANG].missing, true);
    return;
  }
  if (!CONFIG.formEndpoint) { showMsg(form, T[LANG].noEndpoint, true); return; }
  var data = new URLSearchParams(new FormData(form));
  data.set('form', form.id);
  data.set('page_lang', LANG);
  data.set('page', location.href.split('#')[0]);
  data.set('user_agent', navigator.userAgent.slice(0, 300));
  var btn = form.querySelector('[type=submit]');
  btn.disabled = true;
  showMsg(form, T[LANG].sending);
  fetch(CONFIG.formEndpoint, { method: 'POST', mode: 'no-cors', body: data })
    .then(function () { form.reset(); showMsg(form, okText); })
    .catch(function () { showMsg(form, T[LANG].fail, true); })
    .then(function () { btn.disabled = false; });
}

document.addEventListener('DOMContentLoaded', function () {
  // Teléfono y redes desde CONFIG.
  document.querySelectorAll('[data-phone]').forEach(function (a) {
    if (CONFIG.phone) a.href = 'tel:' + CONFIG.phone;
    if (a.hasAttribute('data-phone-text')) a.textContent = CONFIG.phoneDisplay;
  });
  document.querySelectorAll('[data-social]').forEach(function (a) {
    var url = CONFIG[a.dataset.social];
    if (url) a.href = url; else a.hidden = true;
  });
  // WhatsApp: oculto mientras CONFIG.whatsapp esté vacío.
  if (CONFIG.whatsapp) {
    document.querySelectorAll('[data-whatsapp]').forEach(function (el) {
      el.hidden = false;
      (el.tagName === 'A' ? el : el.querySelector('a')).href = 'https://wa.me/' + CONFIG.whatsapp;
    });
  }
  // Botón flotante de llamada: llama directo; con WhatsApp activo abre un menú con las dos opciones.
  var fabBtn = document.getElementById('fabBtn');
  var fabMenu = document.getElementById('fabMenu');
  if (fabBtn) {
    var fabToggle = function (open) {
      fabMenu.hidden = !open;
      fabBtn.setAttribute('aria-expanded', open ? 'true' : 'false');
    };
    fabBtn.addEventListener('click', function () {
      if (!CONFIG.whatsapp) { location.href = 'tel:' + CONFIG.phone; return; }
      fabToggle(fabMenu.hidden);
    });
    document.addEventListener('click', function (e) { if (!e.target.closest('.fab')) fabToggle(false); });
    document.addEventListener('keydown', function (e) { if (e.key === 'Escape') fabToggle(false); });
  }
  // Teléfonos de los formularios: 10 dígitos de EE. UU.; se muestran como (000) 000-0000.
  document.querySelectorAll('form input[type=tel]').forEach(function (el) {
    var check = function () {
      var d = el.value.replace(/\D/g, '');
      if (d.length === 11 && d.charAt(0) === '1') d = d.slice(1);
      var ok = !el.value.trim() || /^[2-9]\d{2}[2-9]\d{6}$/.test(d);
      el.setCustomValidity(ok ? '' : T[LANG].badPhone);
      if (ok && d.length === 10) el.value = '(' + d.slice(0, 3) + ') ' + d.slice(3, 6) + '-' + d.slice(6);
    };
    el.addEventListener('change', check);
    el.addEventListener('blur', check);
  });
  // Aviso de la inscripción anual: del 15 de octubre al 7 de diciembre, hora del Este.
  var aep = document.getElementById('aepBar');
  if (aep) {
    var et = new Date().toLocaleDateString('en-CA', { timeZone: 'America/New_York' }).slice(5);
    var closed = false;
    try { closed = sessionStorage.getItem('amc-aep') === '1'; } catch (e) {}
    aep.hidden = closed || et < '10-15' || et > '12-07';
    document.getElementById('aepClose').addEventListener('click', function () {
      aep.hidden = true;
      try { sessionStorage.setItem('amc-aep', '1'); } catch (e) {}
    });
  }
  // Botón para agrandar la letra.
  var ts = document.getElementById('textSize');
  if (ts) {
    var bigOn = false;
    try { bigOn = localStorage.getItem('amc-text') === 'big'; } catch (e) {}
    var applyText = function () {
      document.documentElement.classList.toggle('big-text', bigOn);
      ts.setAttribute('aria-pressed', bigOn);
      ts.textContent = bigOn ? 'A−' : 'A+';
      updateTextLabel();
    };
    ts.addEventListener('click', function () {
      bigOn = !bigOn;
      try { localStorage.setItem('amc-text', bigOn ? 'big' : 'normal'); } catch (e) {}
      applyText();
    });
    applyText();
  }
  var year = document.getElementById('year');
  if (year) year.textContent = new Date().getFullYear();

  // Menú en celular.
  var toggle = document.querySelector('.menu-toggle');
  toggle.addEventListener('click', function () {
    var open = document.body.classList.toggle('nav-open');
    toggle.setAttribute('aria-expanded', open);
  });
  document.querySelectorAll('.menu a').forEach(function (a) {
    a.addEventListener('click', function () {
      document.body.classList.remove('nav-open');
      toggle.setAttribute('aria-expanded', false);
    });
  });

  // Formulario: los campos cambian según la necesidad.
  var form = document.getElementById('leadForm');
  var chips = form.querySelectorAll('[data-line]');
  function pick(line) {
    chips.forEach(function (c) {
      var on = c.dataset.line === line;
      c.classList.toggle('active', on);
      c.setAttribute('aria-pressed', on);
    });
    form.querySelectorAll('[data-for]').forEach(function (f) { f.hidden = f.dataset.for !== line; });
    form.elements.need.value = line;
  }
  chips.forEach(function (c) { c.addEventListener('click', function () { pick(c.dataset.line); }); });
  pick('nuevo');
  // Las opciones de la portada abren el formulario con la necesidad ya elegida.
  document.querySelectorAll('[data-pick]').forEach(function (a) {
    a.addEventListener('click', function () { pick(a.dataset.pick); });
  });
  form.addEventListener('submit', function (e) { e.preventDefault(); send(form, T[LANG].sent); });
  form.addEventListener('input', function (e) {
    var box = e.target.closest('.invalid');
    if (box) box.classList.remove('invalid');
  });

  // Pestañas de Medicare (con flechas del teclado).
  var tabs = Array.prototype.slice.call(document.querySelectorAll('[role=tab]'));
  function selectTab(tab, focus) {
    tabs.forEach(function (t) {
      var on = t === tab;
      t.setAttribute('aria-selected', on);
      t.tabIndex = on ? 0 : -1;
      document.getElementById(t.getAttribute('aria-controls')).hidden = !on;
    });
    if (focus) tab.focus();
  }
  tabs.forEach(function (t, i) {
    t.addEventListener('click', function () { selectTab(t); });
    t.addEventListener('keydown', function (e) {
      var j = e.key === 'ArrowRight' ? i + 1 : e.key === 'ArrowLeft' ? i - 1 : null;
      if (j === null) return;
      e.preventDefault();
      selectTab(tabs[(j + tabs.length) % tabs.length], true);
    });
  });

  // Fechas clave.
  markOpenPeriods();
  document.getElementById('iepBtn').addEventListener('click', function () { updateIep(true); });

  // "Únase a nuestro equipo" en ventana.
  var dlg = document.getElementById('joinDlg');
  var join = document.getElementById('joinForm');
  document.querySelectorAll('[data-open="joinDlg"]').forEach(function (b) {
    b.addEventListener('click', function () { dlg.showModal(); });
  });
  dlg.querySelector('[data-close]').addEventListener('click', function () { dlg.close(); });
  dlg.addEventListener('click', function (e) { if (e.target === dlg) dlg.close(); });
  join.addEventListener('submit', function (e) { e.preventDefault(); send(join, T[LANG].joinSent); });

  // Idioma: el guardado; si no hay, el del navegador; si no, español.
  document.querySelectorAll('.lang button').forEach(function (b) {
    b.addEventListener('click', function () { setLang(b.dataset.lang); });
  });
  var saved = null;
  try { saved = localStorage.getItem('amc-lang'); } catch (e) {}
  if (!saved && /^en/i.test(navigator.language || '')) saved = 'en';
  if (saved === 'en') setLang('en');
});
