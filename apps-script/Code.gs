/**
 * Assinatura lista digital - serviço (Google Apps Script)
 *
 * Guarda os PDFs no Drive, controla os links numa planilha e recebe o PDF assinado.
 * Uso: rodar configurar() uma vez; depois Implantar > Nova implantação > App da Web
 * (Executar como: eu | Quem tem acesso: qualquer pessoa).
 */

var VALIDADE_HORAS = 72;                       // validade de cada link
var NOME_PASTA = 'Assinatura Lista Digital';
var MAX_PDF_BYTES = 8 * 1024 * 1024;           // limite por PDF
var MAX_FALHAS = 8;                            // erros de PIN seguidos ...
var BLOQUEIO_SEG = 900;                        // ... bloqueiam o acesso por 15 min
var ALFABETO_PIN = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';

var COLS = ['Token', 'Criado em', 'Expira em', 'Status', 'Unidade', 'Data da lista', 'Turno',
  'Arquivo', 'Hash original (SHA-256)', 'Aberto em', 'Assinado em', 'Assinante', 'Dispositivo',
  'Hash assinado (SHA-256)', 'ID original', 'ID assinado'];
var C = {};
COLS.forEach(function (n, i) { C[n] = i; });

/* ---------- Configuração (rodar uma vez no editor) ---------- */
function configurar() {
  var p = PropertiesService.getScriptProperties();
  var raiz;
  if (p.getProperty('PASTA_ID')) {
    raiz = DriveApp.getFolderById(p.getProperty('PASTA_ID'));
  } else {
    raiz = DriveApp.createFolder(NOME_PASTA);
    p.setProperty('PASTA_ID', raiz.getId());
  }
  ['Pendentes', 'Originais', 'Assinados'].forEach(function (n) {
    if (!p.getProperty('PASTA_' + n)) p.setProperty('PASTA_' + n, raiz.createFolder(n).getId());
  });
  if (!p.getProperty('PLANILHA_ID')) {
    var ss = SpreadsheetApp.create('Controle - ' + NOME_PASTA);
    var sh = ss.getSheets()[0];
    sh.setName('Controle');
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    ['Criado em', 'Expira em', 'Aberto em', 'Assinado em'].forEach(function (n) {
      sh.getRange(2, C[n] + 1, 1000, 1).setNumberFormat('dd/MM/yyyy HH:mm');
    });
    DriveApp.getFileById(ss.getId()).moveTo(raiz);
    p.setProperty('PLANILHA_ID', ss.getId());
  }
  if (!p.getProperty('PIN')) p.setProperty('PIN', gerarPin(10));
  Logger.log('Pronto. CÓDIGO DE ACESSO (PIN): ' + formatarPin(p.getProperty('PIN')) +
    '  (anote; se esquecer, rode a função novoPin)');
}

/** Gera um novo código de acesso (use se esquecer o atual). */
function novoPin() {
  var p = PropertiesService.getScriptProperties();
  p.setProperty('PIN', gerarPin(10));
  CacheService.getScriptCache().remove('falhas');
  Logger.log('Novo CÓDIGO DE ACESSO (PIN): ' + formatarPin(p.getProperty('PIN')));
}

/** Libera o acesso se ele foi bloqueado por erros de PIN. */
function desbloquear() {
  CacheService.getScriptCache().remove('falhas');
  Logger.log('Acesso liberado.');
}

/* ---------- Entrada HTTP ---------- */
function doGet() { return saida({ ok: true, servico: 'assinatura-lista-digital' }); }

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    var a = d.acao;
    if (a === 'criar') return saida(criar(d));
    if (a === 'buscar') return saida(buscar(d));
    if (a === 'assinar') return saida(assinar(d));
    if (a === 'listar') return saida(listar(d));
    if (a === 'baixar') return saida(baixar(d));
    if (a === 'cancelar') return saida(cancelar(d));
    return saida({ ok: false, erro: 'acao' });
  } catch (err) {
    return saida({ ok: false, erro: 'interno', detalhe: String(err && err.message || err) });
  }
}

function saida(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Utilidades ---------- */
function prop(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function pasta(n) { return DriveApp.getFolderById(prop('PASTA_' + n)); }
function folha() { return SpreadsheetApp.openById(prop('PLANILHA_ID')).getSheetByName('Controle'); }

function normPin(s) { return String(s == null ? '' : s).toUpperCase().replace(/[^A-Z0-9]/g, ''); }
function gerarPin(n) {
  var h = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  var out = '';
  for (var i = 0; i < n; i++) out += ALFABETO_PIN.charAt(parseInt(h.substr(i * 2, 2), 16) % 32);
  return out;
}
function formatarPin(p) { return p.length === 10 ? p.slice(0, 5) + '-' + p.slice(5) : p; }

/** Devolve null se o PIN estiver certo; senão o objeto de erro. Bloqueia após muitas falhas. */
function autenticar(d) {
  var cache = CacheService.getScriptCache();
  var falhas = Number(cache.get('falhas') || 0);
  if (falhas >= MAX_FALHAS) return { ok: false, erro: 'bloqueado' };
  var pin = normPin(prop('PIN'));
  if (pin && normPin(d.pin) === pin) {
    if (falhas) cache.remove('falhas');
    return null;
  }
  cache.put('falhas', String(falhas + 1), BLOQUEIO_SEG);
  return { ok: false, erro: 'pin' };
}

function hex(bytes) {
  return bytes.map(function (b) { return ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2); }).join('');
}
function sha256(bytes) { return hex(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes)); }
function ehPdf(bytes) { return bytes.length > 4 && bytes[0] === 37 && bytes[1] === 80 && bytes[2] === 68 && bytes[3] === 70; }
function novoToken() { return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').slice(0, 40); }
function limpa(s, n) { return String(s == null ? '' : s).replace(/[\r\n\t]+/g, ' ').trim().slice(0, n); }
function nomeSeguro(s) {
  return limpa(s, 60).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Za-z0-9_\-]+/g, '_').replace(/^_+|_+$/g, '');
}
function iso(v) { return v instanceof Date ? v.toISOString() : (v ? String(v) : ''); }

/** Acha a linha (número 1-based) do token; 0 se não existir. */
function achar(sh, token) {
  token = String(token || '');
  if (token.length < 30) return 0;
  var n = sh.getLastRow() - 1;
  if (n < 1) return 0;
  var col = sh.getRange(2, 1, n, 1).getValues();
  for (var i = 0; i < n; i++) if (String(col[i][0]) === token) return i + 2;
  return 0;
}
function lerLinha(sh, r) { return sh.getRange(r, 1, 1, COLS.length).getValues()[0]; }

/** Status efetivo: Pendente que passou da validade vira Vencido. */
function statusEfetivo(l) {
  var s = String(l[C['Status']]);
  var exp = l[C['Expira em']];
  if (s === 'Pendente' && exp instanceof Date && exp.getTime() < Date.now()) return 'Vencido';
  return s;
}

/* ---------- Ações do administrador (exigem o PIN) ---------- */
function criar(d) {
  var e = autenticar(d); if (e) return e;
  if (!d.pdf) return { ok: false, erro: 'pdf' };
  var bytes = Utilities.base64Decode(d.pdf);
  if (!ehPdf(bytes)) return { ok: false, erro: 'pdf' };
  if (bytes.length > MAX_PDF_BYTES) return { ok: false, erro: 'grande' };

  var token = novoToken();
  var agora = new Date();
  var expira = new Date(agora.getTime() + VALIDADE_HORAS * 3600 * 1000);
  var partes = [nomeSeguro(d.unidade), nomeSeguro(d.data), nomeSeguro(d.turno), token.slice(0, 6)].filter(function (x) { return x; });
  var base = partes.join('_');
  var hash = sha256(bytes);
  var arq = pasta('Pendentes').createFile(Utilities.newBlob(bytes, 'application/pdf', base + '.pdf'));

  var linha = [];
  for (var i = 0; i < COLS.length; i++) linha.push('');
  linha[C['Token']] = token;
  linha[C['Criado em']] = agora;
  linha[C['Expira em']] = expira;
  linha[C['Status']] = 'Pendente';
  linha[C['Unidade']] = limpa(d.unidade, 80);
  linha[C['Data da lista']] = limpa(d.data, 10);
  linha[C['Turno']] = limpa(d.turno, 10);
  linha[C['Arquivo']] = base + '.pdf';
  linha[C['Hash original (SHA-256)']] = hash;
  linha[C['ID original']] = arq.getId();
  folha().appendRow(linha);
  return { ok: true, token: token, expira: expira.toISOString(), hash: hash };
}

function listar(d) {
  var e = autenticar(d); if (e) return e;
  var sh = folha();
  var n = sh.getLastRow() - 1;
  if (n < 1) return { ok: true, itens: [] };
  var ini = Math.max(2, sh.getLastRow() - 99);
  var vals = sh.getRange(ini, 1, sh.getLastRow() - ini + 1, COLS.length).getValues();
  var itens = vals.reverse().map(function (l) {
    return {
      token: String(l[C['Token']]), status: statusEfetivo(l),
      unidade: l[C['Unidade']], data: l[C['Data da lista']], turno: l[C['Turno']],
      criado: iso(l[C['Criado em']]), expira: iso(l[C['Expira em']]),
      aberto: iso(l[C['Aberto em']]), assinadoEm: iso(l[C['Assinado em']]), assinante: l[C['Assinante']]
    };
  });
  return { ok: true, itens: itens };
}

function baixar(d) {
  var e = autenticar(d); if (e) return e;
  var sh = folha(), r = achar(sh, d.token);
  if (!r) return { ok: false, erro: 'invalido' };
  var l = lerLinha(sh, r);
  var id = d.tipo === 'original' ? l[C['ID original']] : l[C['ID assinado']];
  if (!id) return { ok: false, erro: 'sem_arquivo' };
  var arq = DriveApp.getFileById(String(id));
  return { ok: true, nome: arq.getName(), pdf: Utilities.base64Encode(arq.getBlob().getBytes()) };
}

function cancelar(d) {
  var e = autenticar(d); if (e) return e;
  var sh = folha(), r = achar(sh, d.token);
  if (!r) return { ok: false, erro: 'invalido' };
  var l = lerLinha(sh, r);
  if (String(l[C['Status']]) !== 'Pendente') return { ok: false, erro: 'status' };
  sh.getRange(r, C['Status'] + 1).setValue('Cancelado');
  return { ok: true };
}

/* ---------- Ações do contratante (exigem o token do link) ---------- */
function buscar(d) {
  var sh = folha(), r = achar(sh, d.token);
  if (!r) return { ok: false, erro: 'invalido' };
  var l = lerLinha(sh, r);
  var st = statusEfetivo(l);
  if (st === 'Assinado') return { ok: false, erro: 'usado', assinadoEm: iso(l[C['Assinado em']]) };
  if (st === 'Vencido') return { ok: false, erro: 'expirado' };
  if (st !== 'Pendente') return { ok: false, erro: 'invalido' };
  if (!l[C['Aberto em']]) sh.getRange(r, C['Aberto em'] + 1).setValue(new Date());
  var arq = DriveApp.getFileById(String(l[C['ID original']]));
  return {
    ok: true,
    pdf: Utilities.base64Encode(arq.getBlob().getBytes()),
    unidade: l[C['Unidade']], data: l[C['Data da lista']], turno: l[C['Turno']],
    hash: String(l[C['Hash original (SHA-256)']]),
    expira: iso(l[C['Expira em']]), agora: new Date().toISOString()
  };
}

function assinar(d) {
  var nome = limpa(d.nome, 80);
  if (nome.length < 3) return { ok: false, erro: 'nome' };
  if (!d.pdf) return { ok: false, erro: 'pdf' };
  var bytes = Utilities.base64Decode(d.pdf);
  if (!ehPdf(bytes)) return { ok: false, erro: 'pdf' };
  if (bytes.length > MAX_PDF_BYTES) return { ok: false, erro: 'grande' };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = folha(), r = achar(sh, d.token);
    if (!r) return { ok: false, erro: 'invalido' };
    var l = lerLinha(sh, r);
    var st = statusEfetivo(l);
    if (st === 'Assinado') return { ok: false, erro: 'usado', assinadoEm: iso(l[C['Assinado em']]) };
    if (st === 'Vencido') return { ok: false, erro: 'expirado' };
    if (st !== 'Pendente') return { ok: false, erro: 'invalido' };

    var base = String(l[C['Arquivo']]).replace(/\.pdf$/i, '');
    var agora = new Date();
    var assinado = pasta('Assinados').createFile(Utilities.newBlob(bytes, 'application/pdf', base + '_ASSINADO.pdf'));
    try { DriveApp.getFileById(String(l[C['ID original']])).moveTo(pasta('Originais')); } catch (e) {}

    var linha = lerLinha(sh, r);
    linha[C['Status']] = 'Assinado';
    linha[C['Assinado em']] = agora;
    linha[C['Assinante']] = nome;
    linha[C['Dispositivo']] = limpa(d.ua, 300);
    linha[C['Hash assinado (SHA-256)']] = sha256(bytes);
    linha[C['ID assinado']] = assinado.getId();
    sh.getRange(r, 1, 1, COLS.length).setValues([linha]);
    return { ok: true, assinadoEm: agora.toISOString() };
  } finally {
    lock.releaseLock();
  }
}
