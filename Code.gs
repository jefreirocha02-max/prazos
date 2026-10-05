/**
 * SECRETÁRIO JURÍDICO — backend em Google Apps Script
 *
 * - Grava audiências, prazos e demais atos numa planilha do seu Google Drive;
 * - cria cada evento numa agenda própria do Google Agenda ("Prazos e Audiências"),
 *   com lembretes no celular 2 dias antes, 1 dia antes e no próprio dia;
 * - envia todo dia, às 7h, um e-mail-resumo com o que vence hoje, amanhã, em 2 dias
 *   e o que está atrasado (às segundas, inclui a semana inteira);
 * - lê sua agenda principal para mostrar compromissos e conflitos no app.
 *
 * Primeiro uso: selecione a função "setup" e clique em Executar.
 */

const TZ = 'America/Fortaleza';
const CAL_NAME = 'Prazos e Audiências';
const HEAD = ['ID', 'Data', 'Hora', 'Cliente', 'Evento', 'Tipo', 'Processo', 'Órgão/Vara',
  'Observações', 'Status', 'Origem', 'ID Google Agenda', 'Criado em'];
const NCOL = HEAD.length;
const COL = {}; HEAD.forEach(function (h, i) { COL[h] = i; });
const CORES = { 'Audiência': 'RED', 'Prazo': 'ORANGE', 'Perícia': 'MAUVE', 'Sessão': 'PALE_RED',
  'Atendimento': 'PALE_GREEN', 'Diligência': 'CYAN', 'Outro': 'GRAY' };

/* ------------------------------------------------------------------ */
/* Configuração inicial                                                */
/* ------------------------------------------------------------------ */

function setup() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  const savedId = props.getProperty('SHEET_ID');
  if (savedId) { try { ss = SpreadsheetApp.openById(savedId); } catch (e) { ss = null; } }
  if (!ss) { try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ss = null; } }
  if (!ss) ss = SpreadsheetApp.create('Secretário Jurídico — Agenda de Prazos e Audiências');
  props.setProperty('SHEET_ID', ss.getId());
  ss.setSpreadsheetLocale('pt_BR');
  ss.setSpreadsheetTimeZone(TZ);

  // agenda própria no Google Agenda
  let cal = null;
  const calId = props.getProperty('CAL_ID');
  if (calId) cal = CalendarApp.getCalendarById(calId);
  if (!cal) cal = CalendarApp.getOwnedCalendarsByName(CAL_NAME)[0];
  if (!cal) cal = CalendarApp.createCalendar(CAL_NAME, { color: CalendarApp.Color.RED, timeZone: TZ });
  props.setProperty('CAL_ID', cal.getId());

  let token = props.getProperty('TOKEN');
  if (!token) {
    token = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
    props.setProperty('TOKEN', token);
  }

  agendaSheet_(ss);
  feriadosSheet_(ss);
  const cfg = configSheet_(ss);
  cfg.getRange('B8').setValue(token);
  resumoSheet_(ss);

  ss.getSheets().forEach(function (sh) {
    if (/^(Sheet1|Página1|Planilha1|Página 1)$/.test(sh.getName()) && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });

  // gatilho do e-mail diário às 7h
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'resumoDiario') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('resumoDiario').timeBased().everyDays(1).atHour(7).inTimezone(TZ).create();

  Logger.log('Planilha: ' + ss.getUrl());
  Logger.log('Agenda criada/usada: ' + CAL_NAME);
  Logger.log('TOKEN de acesso para o app: ' + token);
  return { url: ss.getUrl(), token: token };
}

function trocarToken() {
  const token = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
  PropertiesService.getScriptProperties().setProperty('TOKEN', token);
  configSheet_(ss_()).getRange('B8').setValue(token);
  Logger.log('Novo TOKEN: ' + token);
}

/** Recria no Google Agenda os eventos da planilha que estiverem sem vínculo (útil após importar linhas). */
function sincronizarAgenda() {
  const ss = ss_(); const sh = agendaSheet_(ss); const cfg = readConfig_(ss); const cal = cal_();
  const last = lastDataRow_(sh); if (last < 2) return;
  const rows = sh.getRange(2, 1, last - 1, NCOL).getValues();
  rows.forEach(function (r, i) {
    if (!r[COL['ID']]) return;
    const it = rowToItem_(r);
    const id = upsertEvent_(cal, it, cfg, it.gcalId);
    sh.getRange(i + 2, COL['ID Google Agenda'] + 1).setValue(id);
  });
}

/* ------------------------------------------------------------------ */
/* HTTP                                                                */
/* ------------------------------------------------------------------ */

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action) return json_(handle_(p));
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Secretário Jurídico')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, error: 'Requisição inválida.' }); }
  return json_(handle_(req));
}

function api(req) { return handle_(req || {}); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function handle_(req) {
  try {
    const token = PropertiesService.getScriptProperties().getProperty('TOKEN');
    if (!token) return { ok: false, error: 'Backend não configurado: execute a função setup() no editor do Apps Script.' };
    if (String(req.token || '') !== token) return { ok: false, error: 'Token de acesso inválido.' };
    const ss = ss_();

    switch (req.action) {
      case 'ping':
        return { ok: true, sheetUrl: ss.getUrl() };

      case 'list':
        return { ok: true, items: listAll_(ss), config: readConfig_(ss), sheetUrl: ss.getUrl(),
          google: req.from && req.to ? googleEvents_(req.from, req.to) : [] };

      case 'google':
        return { ok: true, google: googleEvents_(req.from, req.to) };

      case 'add': {
        const lock = LockService.getScriptLock(); lock.waitLock(25000);
        try {
          const cfg = readConfig_(ss); const cal = cal_();
          return { ok: true, items: (req.items || []).map(function (it) { return addItem_(ss, cal, cfg, it); }) };
        } finally { lock.releaseLock(); }
      }

      case 'update': {
        const lock = LockService.getScriptLock(); lock.waitLock(25000);
        try { return { ok: true, item: updateItem_(ss, req.item || {}) }; }
        finally { lock.releaseLock(); }
      }

      case 'delete': {
        const lock = LockService.getScriptLock(); lock.waitLock(25000);
        try { return { ok: deleteItem_(ss, String(req.id)) }; }
        finally { lock.releaseLock(); }
      }

      case 'setConfig':
        writeConfig_(ss, req.config || {});
        return { ok: true, config: readConfig_(ss) };

      case 'testEmail':
        resumoDiario(true);
        return { ok: true };

      default:
        return { ok: false, error: 'Ação desconhecida: ' + req.action };
    }
  } catch (e) {
    return { ok: false, error: String((e && e.message) || e) };
  }
}

/* ------------------------------------------------------------------ */
/* Planilha                                                            */
/* ------------------------------------------------------------------ */

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) { setup(); return ss_(); }
  return SpreadsheetApp.openById(id);
}

function cal_() {
  const id = PropertiesService.getScriptProperties().getProperty('CAL_ID');
  const cal = id ? CalendarApp.getCalendarById(id) : null;
  if (!cal) throw new Error('Agenda "' + CAL_NAME + '" não encontrada. Execute setup() novamente.');
  return cal;
}

function agendaSheet_(ss) {
  let sh = ss.getSheetByName('Agenda');
  if (sh) return sh;
  sh = ss.insertSheet('Agenda', 0);
  sh.getRange(1, 1, 1, NCOL).setValues([HEAD]).setFontWeight('bold').setBackground('#1f3a5f').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.getRange('B:B').setNumberFormat('dd/MM/yyyy (ddd)');
  sh.getRange('C:C').setNumberFormat('@');
  sh.getRange('G:G').setNumberFormat('@');
  sh.getRange('M:M').setNumberFormat('dd/MM/yyyy HH:mm');
  sh.setColumnWidth(1, 80); sh.setColumnWidth(2, 130); sh.setColumnWidth(4, 200); sh.setColumnWidth(5, 220);
  sh.setColumnWidth(7, 200); sh.setColumnWidth(8, 180); sh.setColumnWidth(9, 260);
  sh.hideColumns(COL['ID Google Agenda'] + 1);
  const rules = sh.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=($J2="Pendente")*($B2<TODAY())')
    .setBackground('#f8d7da').setRanges([sh.getRange('A2:M')]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=($J2="Pendente")*($B2>=TODAY())*($B2<=TODAY()+2)')
    .setBackground('#fff3cd').setRanges([sh.getRange('A2:M')]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$J2="Concluído"')
    .setFontColor('#8a8a8a').setRanges([sh.getRange('A2:M')]).build());
  sh.setConditionalFormatRules(rules);
  return sh;
}

function resumoSheet_(ss) {
  let sh = ss.getSheetByName('Resumo');
  if (sh) return sh;
  sh = ss.insertSheet('Resumo', 1);
  sh.getRange('A1').setValue('Atrasados (pendentes com data já vencida)').setFontWeight('bold').setFontColor('#b02a37');
  sh.getRange('A2').setFormula('=IFERROR(QUERY(Agenda!A2:J,"select B, C, D, E, F, G where J = \'Pendente\' and B < date \'"&TEXT(TODAY(),"yyyy-mm-dd")&"\' order by B label B \'Data\', C \'Hora\', D \'Cliente\', E \'Evento\', F \'Tipo\', G \'Processo\'",0),"Nenhum")');
  sh.getRange('H1').setValue('Próximos 30 dias').setFontWeight('bold');
  sh.getRange('H2').setFormula('=IFERROR(QUERY(Agenda!A2:J,"select B, C, D, E, F, G where J = \'Pendente\' and B >= date \'"&TEXT(TODAY(),"yyyy-mm-dd")&"\' and B <= date \'"&TEXT(TODAY()+30,"yyyy-mm-dd")&"\' order by B, C label B \'Data\', C \'Hora\', D \'Cliente\', E \'Evento\', F \'Tipo\', G \'Processo\'",0),"Nenhum")');
  sh.getRange('O1').setValue('Eventos por mês').setFontWeight('bold');
  sh.getRange('O2').setFormula('=IFERROR(QUERY(Agenda!A2:J,"select year(B), month(B)+1, count(A) where B is not null group by year(B), month(B)+1 order by year(B), month(B)+1 label year(B) \'Ano\', month(B)+1 \'Mês\', count(A) \'Eventos\'",0),"")');
  sh.getRange('A:A').setNumberFormat('dd/MM/yyyy');
  sh.getRange('H:H').setNumberFormat('dd/MM/yyyy');
  return sh;
}

function feriadosSheet_(ss) {
  let sh = ss.getSheetByName('Feriados');
  if (sh) return sh;
  sh = ss.insertSheet('Feriados');
  sh.getRange('A1:B1').setValues([['Data', 'Descrição']]).setFontWeight('bold');
  sh.getRange('A:A').setNumberFormat('dd/MM/yyyy');
  sh.getRange('D1').setValue('Inclua aqui feriados locais e suspensões de expediente do TJCE, TRT-7, TRF-5 e da sua comarca. Feriados nacionais e móveis já são considerados pelo app.');
  sh.setColumnWidth(2, 260);
  return sh;
}

function configSheet_(ss) {
  let sh = ss.getSheetByName('Config');
  if (sh) return sh;
  sh = ss.insertSheet('Config');
  sh.getRange('A1:B1').setValues([['Parâmetro', 'Valor']]).setFontWeight('bold');
  sh.getRange('A2:B7').setValues([
    ['Horário padrão dos prazos (sem hora informada)', '08:00'],
    ['Lembrete no dia de audiências: minutos antes', 120],
    ['Contagem padrão de prazos (úteis ou corridos)', 'úteis'],
    ['Suspender prazos entre 20/12 e 20/01 (Sim/Não)', 'Sim'],
    ['Enviar e-mail diário às 7h (Sim/Não)', 'Sim'],
    ['E-mail para o resumo diário', Session.getEffectiveUser().getEmail()]
  ]);
  sh.getRange('A8').setValue('Token de acesso do app');
  sh.getRange('B2').setNumberFormat('@');
  sh.setColumnWidth(1, 340); sh.setColumnWidth(2, 260);
  return sh;
}

function readConfig_(ss) {
  const sh = configSheet_(ss);
  const v = sh.getRange('B2:B7').getValues().map(function (r) { return r[0]; });
  let hora = v[0];
  if (hora instanceof Date) hora = Utilities.formatDate(hora, TZ, 'HH:mm');
  const fer = feriadosSheet_(ss);
  const lastF = fer.getLastRow();
  const feriados = lastF >= 2 ? fer.getRange(2, 1, lastF - 1, 2).getValues()
    .filter(function (r) { return r[0] instanceof Date; })
    .map(function (r) { return { date: Utilities.formatDate(r[0], TZ, 'yyyy-MM-dd'), desc: String(r[1] || '') }; }) : [];
  return {
    horaPrazo: /^\d{1,2}:\d{2}$/.test(String(hora)) ? String(hora) : '08:00',
    antecedencia: Number(v[1]) || 120,
    contagem: /corr/i.test(String(v[2])) ? 'corridos' : 'úteis',
    recesso: !/^n/i.test(String(v[3])),
    emailDiario: !/^n/i.test(String(v[4])),
    email: String(v[5] || Session.getEffectiveUser().getEmail()),
    feriados: feriados
  };
}

function writeConfig_(ss, c) {
  const sh = configSheet_(ss);
  if (c.horaPrazo) sh.getRange('B2').setValue(String(c.horaPrazo));
  if (c.antecedencia != null) sh.getRange('B3').setValue(Number(c.antecedencia) || 120);
  if (c.contagem) sh.getRange('B4').setValue(c.contagem);
  if (c.recesso != null) sh.getRange('B5').setValue(c.recesso ? 'Sim' : 'Não');
  if (c.emailDiario != null) sh.getRange('B6').setValue(c.emailDiario ? 'Sim' : 'Não');
  if (c.email) sh.getRange('B7').setValue(String(c.email));
}

function lastDataRow_(sh) {
  const n = sh.getMaxRows();
  const vals = sh.getRange(1, 1, n, 1).getValues();
  for (let i = n - 1; i >= 0; i--) if (vals[i][0] !== '' && vals[i][0] !== null) return i + 1;
  return 1;
}

function toDate_(iso) {
  const p = String(iso).split('-').map(Number);
  return new Date(p[0], p[1] - 1, p[2], 12, 0, 0);
}

function rowToItem_(r) {
  const d = r[COL['Data']] instanceof Date ? r[COL['Data']] : toDate_(r[COL['Data']]);
  let hora = r[COL['Hora']];
  if (hora instanceof Date) hora = Utilities.formatDate(hora, TZ, 'HH:mm');
  return {
    id: String(r[COL['ID']]), date: Utilities.formatDate(d, TZ, 'yyyy-MM-dd'), hora: String(hora || ''),
    cliente: String(r[COL['Cliente']]), evento: String(r[COL['Evento']]), tipo: String(r[COL['Tipo']]),
    processo: String(r[COL['Processo']]), orgao: String(r[COL['Órgão/Vara']]), obs: String(r[COL['Observações']]),
    status: String(r[COL['Status']] || 'Pendente'), origem: String(r[COL['Origem']]), gcalId: String(r[COL['ID Google Agenda']] || '')
  };
}

function itemToRow_(it, gcalId, criado) {
  return [it.id, toDate_(it.date), String(it.hora || ''), String(it.cliente || ''), String(it.evento || ''),
    String(it.tipo || 'Outro'), String(it.processo || ''), String(it.orgao || ''), String(it.obs || ''),
    it.status === 'Concluído' ? 'Concluído' : 'Pendente', String(it.origem || 'texto'), gcalId || '', criado || new Date()];
}

function listAll_(ss) {
  const sh = agendaSheet_(ss);
  const last = lastDataRow_(sh);
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, NCOL).getValues()
    .filter(function (r) { return r[0] !== ''; }).map(rowToItem_);
}

function findRow_(sh, id) {
  const last = lastDataRow_(sh);
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return -1;
}

function addItem_(ss, cal, cfg, it) {
  const sh = agendaSheet_(ss);
  it.id = String(it.id || Utilities.getUuid().slice(0, 10));
  it.status = it.status === 'Concluído' ? 'Concluído' : 'Pendente';
  let gcalId = '';
  try { gcalId = upsertEvent_(cal, it, cfg, null); } catch (e) { it.obs = (it.obs ? it.obs + ' | ' : '') + 'Falha ao criar na agenda: ' + e.message; }
  const r = lastDataRow_(sh) + 1;
  if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 50);
  sh.getRange(r, 1, 1, NCOL).setValues([itemToRow_(it, gcalId)]);
  it.gcalId = gcalId;
  return it;
}

function updateItem_(ss, it) {
  const sh = agendaSheet_(ss);
  const r = findRow_(sh, String(it.id));
  if (r < 0) throw new Error('Evento não encontrado na planilha.');
  const old = rowToItem_(sh.getRange(r, 1, 1, NCOL).getValues()[0]);
  const merged = Object.assign({}, old, it);
  const gcalId = upsertEvent_(cal_(), merged, readConfig_(ss), old.gcalId);
  const criado = sh.getRange(r, COL['Criado em'] + 1).getValue();
  sh.getRange(r, 1, 1, NCOL).setValues([itemToRow_(merged, gcalId, criado)]);
  merged.gcalId = gcalId;
  return merged;
}

function deleteItem_(ss, id) {
  const sh = agendaSheet_(ss);
  const r = findRow_(sh, id);
  if (r < 0) return false;
  const gcalId = String(sh.getRange(r, COL['ID Google Agenda'] + 1).getValue() || '');
  if (gcalId) { try { const ev = cal_().getEventById(gcalId); if (ev) ev.deleteEvent(); } catch (e) {} }
  sh.deleteRow(r);
  return true;
}

/* ------------------------------------------------------------------ */
/* Google Agenda                                                       */
/* ------------------------------------------------------------------ */

function upsertEvent_(cal, it, cfg, gcalId) {
  const temHora = /^\d{1,2}:\d{2}$/.test(String(it.hora || ''));
  const hora = temHora ? it.hora : cfg.horaPrazo;
  const start = Utilities.parseDate(it.date + ' ' + hora, TZ, 'yyyy-MM-dd HH:mm');
  const dur = it.tipo === 'Audiência' || it.tipo === 'Sessão' || it.tipo === 'Perícia' ? 60 : 30;
  const end = new Date(start.getTime() + dur * 60000);
  const concluido = it.status === 'Concluído';
  const title = (concluido ? '[CONCLUÍDO] ' : '') + (it.evento || it.tipo || 'Evento') + (it.cliente ? ' — ' + it.cliente : '');
  const desc = [
    'Cliente: ' + (it.cliente || '—'),
    'Evento: ' + (it.evento || '—') + ' (' + (it.tipo || 'Outro') + ')',
    it.processo ? 'Processo: ' + it.processo : '',
    it.orgao ? 'Órgão/Vara: ' + it.orgao : '',
    temHora ? '' : 'Prazo sem horário definido: lançado às ' + hora + '.',
    it.obs ? 'Obs.: ' + it.obs : ''
  ].filter(String).join('\n') + '\n\nRegistrado pelo Secretário Jurídico.';

  let ev = null;
  if (gcalId) { try { ev = cal.getEventById(gcalId); } catch (e) { ev = null; } }
  if (ev) {
    ev.setTitle(title); ev.setTime(start, end); ev.setDescription(desc); ev.setLocation(it.orgao || '');
  } else {
    ev = cal.createEvent(title, start, end, { description: desc, location: it.orgao || '' });
  }

  ev.removeAllReminders();
  if (!concluido) {
    ev.addPopupReminder(2 * 24 * 60);   // 2 dias antes
    ev.addPopupReminder(24 * 60);       // 1 dia antes
    // no próprio dia: na hora do prazo, ou X minutos antes da audiência (sem voltar para o dia anterior)
    const hm = hora.split(':').map(Number);
    const minutosDoDia = hm[0] * 60 + hm[1];
    const noDia = temHora ? Math.min(cfg.antecedencia, Math.max(0, minutosDoDia - 1)) : 0;
    ev.addPopupReminder(noDia);
  }
  const cor = CalendarApp.EventColor[CORES[it.tipo] || 'GRAY'];
  if (cor) ev.setColor(cor);
  return ev.getId();
}

/** Compromissos da agenda principal (somente leitura), para exibir no app e apontar conflitos. */
function googleEvents_(from, to) {
  if (!from || !to) return [];
  const ini = Utilities.parseDate(from + ' 00:00', TZ, 'yyyy-MM-dd HH:mm');
  const fim = Utilities.parseDate(to + ' 23:59', TZ, 'yyyy-MM-dd HH:mm');
  return CalendarApp.getDefaultCalendar().getEvents(ini, fim).slice(0, 400).map(function (ev) {
    const s = ev.getStartTime(), e = ev.getEndTime();
    return {
      title: ev.getTitle(), date: Utilities.formatDate(s, TZ, 'yyyy-MM-dd'),
      hora: ev.isAllDayEvent() ? '' : Utilities.formatDate(s, TZ, 'HH:mm'),
      fim: ev.isAllDayEvent() ? '' : Utilities.formatDate(e, TZ, 'HH:mm'),
      allDay: ev.isAllDayEvent()
    };
  });
}

/* ------------------------------------------------------------------ */
/* E-mail diário (gatilho às 7h)                                       */
/* ------------------------------------------------------------------ */

function resumoDiario(forcar) {
  const ss = ss_();
  const cfg = readConfig_(ss);
  if (!cfg.emailDiario && forcar !== true) return;
  const hoje = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const add = function (n) { const d = toDate_(hoje); d.setDate(d.getDate() + n); return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); };
  const amanha = add(1), dois = add(2), semana = add(7);
  const segunda = Utilities.formatDate(new Date(), TZ, 'u') === '1';
  const pend = listAll_(ss).filter(function (i) { return i.status !== 'Concluído'; })
    .sort(function (a, b) { return (a.date + a.hora).localeCompare(b.date + b.hora); });

  const grupos = [
    ['Atrasados', pend.filter(function (i) { return i.date < hoje; }), '#b02a37'],
    ['Hoje', pend.filter(function (i) { return i.date === hoje; }), '#d9480f'],
    ['Amanhã', pend.filter(function (i) { return i.date === amanha; }), '#c88a04'],
    ['Em 2 dias', pend.filter(function (i) { return i.date === dois; }), '#1f3a5f']
  ];
  if (segunda) grupos.push(['Restante da semana', pend.filter(function (i) { return i.date > dois && i.date <= semana; }), '#555555']);

  const total = grupos.reduce(function (a, g) { return a + g[1].length; }, 0);
  if (!total && forcar !== true) return;

  const fmt = function (iso) { return iso.split('-').reverse().join('/'); };
  let html = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#222">';
  grupos.forEach(function (g) {
    if (!g[1].length) return;
    html += '<h3 style="color:' + g[2] + ';margin:18px 0 6px">' + g[0] + ' (' + g[1].length + ')</h3><table cellpadding="6" style="border-collapse:collapse;width:100%">';
    g[1].forEach(function (i) {
      html += '<tr style="border-bottom:1px solid #eee"><td style="white-space:nowrap">' + fmt(i.date) + (i.hora ? ' ' + i.hora : '') +
        '</td><td><b>' + esc_(i.evento) + '</b><br>' + esc_(i.cliente || '—') + (i.processo ? ' · ' + esc_(i.processo) : '') +
        (i.orgao ? ' · ' + esc_(i.orgao) : '') + '</td></tr>';
    });
    html += '</table>';
  });
  if (!total) html += '<p>Nenhum prazo ou audiência nos próximos dois dias.</p>';
  html += '<p style="color:#888;font-size:12px;margin-top:20px">Planilha: <a href="' + ss.getUrl() + '">abrir</a> · Secretário Jurídico</p></div>';

  const n = function (k) { return grupos[k][1].length; };
  const assunto = 'Agenda jurídica: ' + n(1) + ' hoje, ' + n(2) + ' amanhã, ' + n(3) + ' em 2 dias' + (n(0) ? ' · ' + n(0) + ' atrasado(s)' : '');
  MailApp.sendEmail({ to: cfg.email, subject: assunto, htmlBody: html, name: 'Secretário Jurídico' });
}

function esc_(s) {
  return String(s || '').replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; });
}
