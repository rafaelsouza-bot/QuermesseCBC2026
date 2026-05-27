// ============================================================
//  QUERMESSE — Backend (Google Apps Script)
//  Cole este código em: script.google.com → novo projeto
//  Depois: Implantar → Nova implantação → App da Web
//    - Executar como: Eu mesmo
//    - Quem tem acesso: Qualquer pessoa
// ============================================================

var SHEET_ID          = '174zZLsKEhbpcfLXtnWn2X8xJ0KQBA9LYkBEEbMTVC4I';
var SHEET_NAME        = 'Ingressos';
var SENHA_PORTEIRO    = 'Cbc@2026';
var SENHA_ORGANIZADOR = 'Org@2026';

// ── Ponto de entrada HTTP ────────────────────────────────────
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var acao = data.acao;
    if (acao === 'comprar')   return resposta(comprarIngresso(data));
    if (acao === 'validar')   return resposta(validarIngresso(data));
    if (acao === 'dashboard') return resposta(getDashboard(data));
    if (acao === 'listar')    return resposta(listarIngressos(data));
    if (acao === 'aprovar')   return resposta(aprovarIngresso(data));
    if (acao === 'rejeitar')  return resposta(rejeitarIngresso(data));
    if (acao === 'cancelar')  return resposta(cancelarIngresso(data));
    return resposta({ ok: false, erro: 'Ação desconhecida' });
  } catch (err) {
    return resposta({ ok: false, erro: err.message });
  }
}

function doGet(e) {
  return resposta({ ok: true, msg: 'API Quermesse no ar!' });
}

// ── Comprar ingresso (suporta múltiplos) ─────────────────────
function comprarIngresso(data) {
  var sheet    = getSheet();
  var agora    = new Date();
  var quantidade = parseInt(data.quantidade) || 1;
  if (quantidade < 1 || quantidade > 10) return { ok: false, erro: 'Quantidade inválida' };
  if (!data.nome || !data.contato || !data.tipo || !data.pagamento) {
    return { ok: false, erro: 'Campos obrigatórios faltando' };
  }

  var ids = [];
  var valorUnitario = parseFloat(data.valor) / quantidade;

  for (var i = 0; i < quantidade; i++) {
    var id = gerarID();
    ids.push(id);
    var linha = [
      id,
      data.nome,
      data.contato,
      data.tipo,
      data.pagamento,
      valorUnitario.toFixed(2),
      'PENDENTE',
      Utilities.formatDate(agora, 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm'),
      '',
      '',
      quantidade > 1 ? (i+1) + '/' + quantidade : ''
    ];
    sheet.appendRow(linha);
  }

  return { ok: true, ids: ids, nome: data.nome, tipo: data.tipo, quantidade: quantidade, total: data.valor };
}

// ── Validar ingresso (porteiro) ──────────────────────────────
function validarIngresso(data) {
  if (data.senha !== SENHA_PORTEIRO) {
    return { ok: false, erro: 'Senha do porteiro incorreta', codigo: 'SENHA' };
  }
  var sheet  = getSheet();
  var id     = data.id;
  var linhas = sheet.getDataRange().getValues();

  for (var i = 1; i < linhas.length; i++) {
    if (linhas[i][0] === id) {
      var status = linhas[i][6];
      var nome   = linhas[i][1];
      var tipo   = linhas[i][3];
      var seq    = linhas[i][10] || '';

      if (status === 'PENDENTE')  return { ok: false, codigo: 'PENDENTE',      nome: nome, tipo: tipo, erro: 'Pagamento ainda não confirmado pela organização' };
      if (status === 'USADO')     return { ok: false, codigo: 'JA_USADO',      nome: nome, tipo: tipo, dataUso: linhas[i][8], erro: 'Ingresso já utilizado em ' + linhas[i][8] };
      if (status === 'CANCELADO') return { ok: false, codigo: 'CANCELADO',     nome: nome, erro: 'Ingresso cancelado' };

      var agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
      sheet.getRange(i + 1, 7).setValue('USADO');
      sheet.getRange(i + 1, 9).setValue(agora);

      return { ok: true, codigo: 'LIBERADO', nome: nome, tipo: tipo, seq: seq, dataUso: agora };
    }
  }
  return { ok: false, codigo: 'NAO_ENCONTRADO', erro: 'Ingresso não encontrado' };
}

// ── Listar ingressos (organizador) ───────────────────────────
function listarIngressos(data) {
  if (data.senha !== SENHA_ORGANIZADOR) return { ok: false, erro: 'Senha incorreta' };
  var sheet  = getSheet();
  var linhas = sheet.getDataRange().getValues();
  var lista  = [];
  var filtro = data.filtro || 'TODOS';

  for (var i = 1; i < linhas.length; i++) {
    if (!linhas[i][0]) continue;
    var status = linhas[i][6];
    if (filtro !== 'TODOS' && status !== filtro) continue;
    lista.push({
      id: linhas[i][0], nome: linhas[i][1], contato: linhas[i][2],
      tipo: linhas[i][3], pagamento: linhas[i][4], valor: linhas[i][5],
      status: status, dataCompra: linhas[i][7], dataUso: linhas[i][8],
      obs: linhas[i][10] || ''
    });
  }
  lista.reverse();
  return { ok: true, lista: lista, total: lista.length };
}

// ── Aprovar ──────────────────────────────────────────────────
function aprovarIngresso(data) {
  if (data.senha !== SENHA_ORGANIZADOR) return { ok: false, erro: 'Senha incorreta' };
  return mudarStatus(data.id, 'ATIVO', 'Aprovado pela organização');
}

// ── Rejeitar ─────────────────────────────────────────────────
function rejeitarIngresso(data) {
  if (data.senha !== SENHA_ORGANIZADOR) return { ok: false, erro: 'Senha incorreta' };
  return mudarStatus(data.id, 'CANCELADO', data.obs || 'Rejeitado pela organização');
}

// ── Cancelar ─────────────────────────────────────────────────
function cancelarIngresso(data) {
  if (data.senha !== SENHA_ORGANIZADOR) return { ok: false, erro: 'Senha incorreta' };
  return mudarStatus(data.id, 'CANCELADO', 'Cancelado pela organização');
}

function mudarStatus(id, novoStatus, obs) {
  var sheet  = getSheet();
  var linhas = sheet.getDataRange().getValues();
  for (var i = 1; i < linhas.length; i++) {
    if (linhas[i][0] === id) {
      sheet.getRange(i + 1, 7).setValue(novoStatus);
      sheet.getRange(i + 1, 11).setValue(obs);
      return { ok: true, id: id, status: novoStatus };
    }
  }
  return { ok: false, erro: 'Ingresso não encontrado' };
}

// ── Dashboard ────────────────────────────────────────────────
function getDashboard(data) {
  if (data.senha !== SENHA_PORTEIRO && data.senha !== SENHA_ORGANIZADOR) {
    return { ok: false, erro: 'Senha incorreta' };
  }
  var sheet  = getSheet();
  var linhas = sheet.getDataRange().getValues();
  var total = 0, pendentes = 0, ativos = 0, usados = 0, cancelados = 0;

  for (var i = 1; i < linhas.length; i++) {
    if (!linhas[i][0]) continue;
    total++;
    var s = linhas[i][6];
    if (s === 'PENDENTE')  pendentes++;
    if (s === 'ATIVO')     ativos++;
    if (s === 'USADO')     usados++;
    if (s === 'CANCELADO') cancelados++;
  }
  return { ok: true, total: total, pendentes: pendentes, ativos: ativos, usados: usados, cancelados: cancelados };
}

// ── Helpers ──────────────────────────────────────────────────
function getSheet() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
}

function gerarID() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var id = 'QRM-';
  for (var i = 0; i < 8; i++) id += chars.charAt(Math.floor(Math.random() * chars.length));
  return id;
}

function resposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ── Configurar planilha (rode uma vez) ───────────────────────
function configurarPlanilha() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  var headers = ['ID','Nome','Contato','Tipo','Pagamento','Valor','Status','Data Compra','Data Uso','ID Pagamento','Seq/Obs'];
  sheet.getRange(1,1,1,headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#0a3550').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  Logger.log('Planilha configurada!');
}
