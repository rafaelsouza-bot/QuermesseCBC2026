// ============================================================
//  QUERMESSE — Backend (Google Apps Script)
//  Cole este código em: script.google.com → novo projeto
//  Depois: Implantar → Nova implantação → App da Web
//    - Executar como: Eu mesmo
//    - Quem tem acesso: Qualquer pessoa
// ============================================================

var SHEET_ID        = '174zZLsKEhbpcfLXtnWn2X8xJ0KQBA9LYkBEEbMTVC4I';
var SHEET_NAME      = 'Ingressos';
var SENHA_PORTEIRO  = 'Cbc@2026';
var SENHA_ORGANIZADOR = 'Org@2026*'; // Troque por uma senha segura

// ── Ponto de entrada HTTP ────────────────────────────────────
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var acao = data.acao;

    if (acao === 'comprar')    return resposta(comprarIngresso(data));
    if (acao === 'validar')    return resposta(validarIngresso(data));
    if (acao === 'dashboard')  return resposta(getDashboard(data));
    if (acao === 'listar')     return resposta(listarIngressos(data));
    if (acao === 'aprovar')    return resposta(aprovarIngresso(data));
    if (acao === 'rejeitar')   return resposta(rejeitarIngresso(data));
    if (acao === 'cancelar')   return resposta(cancelarIngresso(data));

    return resposta({ ok: false, erro: 'Ação desconhecida' });
  } catch (err) {
    return resposta({ ok: false, erro: err.message });
  }
}

function doGet(e) {
  return resposta({ ok: true, msg: 'API Quermesse no ar!' });
}

// ── Comprar ingresso ─────────────────────────────────────────
function comprarIngresso(data) {
  var sheet = getSheet();
  var id    = gerarID();
  var agora = new Date();

  if (!data.nome || !data.contato || !data.tipo || !data.pagamento) {
    return { ok: false, erro: 'Campos obrigatórios faltando' };
  }

  var limite = getLimite(data.tipo);
  if (limite > 0) {
    var vendidos = contarPorTipo(sheet, data.tipo);
    if (vendidos >= limite) {
      return { ok: false, erro: 'Ingressos esgotados para ' + data.tipo };
    }
  }

  // PIX e Dinheiro ficam PENDENTE; Débito/Crédito ficam PENDENTE também
  // A organização aprova todos via painel
  var status = 'PENDENTE';

  var linha = [
    id,
    data.nome,
    data.contato,
    data.tipo,
    data.pagamento,
    data.valor || '',
    status,
    Utilities.formatDate(agora, 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm'),
    '',   // Data uso
    data.paymentId || '',
    ''    // Observação
  ];

  sheet.appendRow(linha);

  return { ok: true, id: id, nome: data.nome, tipo: data.tipo, valor: data.valor, status: status };
}

// ── Listar ingressos (organizador) ───────────────────────────
function listarIngressos(data) {
  if (data.senha !== SENHA_ORGANIZADOR) {
    return { ok: false, erro: 'Senha incorreta' };
  }

  var sheet  = getSheet();
  var linhas = sheet.getDataRange().getValues();
  var lista  = [];
  var filtro = data.filtro || 'TODOS'; // TODOS, PENDENTE, ATIVO, USADO, CANCELADO

  for (var i = 1; i < linhas.length; i++) {
    if (!linhas[i][0]) continue;
    var status = linhas[i][6];
    if (filtro !== 'TODOS' && status !== filtro) continue;

    lista.push({
      id:        linhas[i][0],
      nome:      linhas[i][1],
      contato:   linhas[i][2],
      tipo:      linhas[i][3],
      pagamento: linhas[i][4],
      valor:     linhas[i][5],
      status:    status,
      dataCompra: linhas[i][7],
      dataUso:   linhas[i][8],
      obs:       linhas[i][10] || ''
    });
  }

  // Mais recentes primeiro
  lista.reverse();
  return { ok: true, lista: lista, total: lista.length };
}

// ── Aprovar ingresso ─────────────────────────────────────────
function aprovarIngresso(data) {
  if (data.senha !== SENHA_ORGANIZADOR) {
    return { ok: false, erro: 'Senha incorreta' };
  }
  return mudarStatus(data.id, 'ATIVO', data.obs || 'Aprovado pela organização');
}

// ── Rejeitar ingresso ────────────────────────────────────────
function rejeitarIngresso(data) {
  if (data.senha !== SENHA_ORGANIZADOR) {
    return { ok: false, erro: 'Senha incorreta' };
  }
  return mudarStatus(data.id, 'CANCELADO', data.obs || 'Rejeitado pela organização');
}

// ── Cancelar ingresso ────────────────────────────────────────
function cancelarIngresso(data) {
  if (data.senha !== SENHA_ORGANIZADOR) {
    return { ok: false, erro: 'Senha incorreta' };
  }
  return mudarStatus(data.id, 'CANCELADO', data.obs || 'Cancelado pela organização');
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

      if (status === 'PENDENTE') {
        return { ok: false, codigo: 'PENDENTE', nome: nome, tipo: tipo, erro: 'Ingresso ainda não aprovado pela organização' };
      }
      if (status === 'USADO') {
        return { ok: false, codigo: 'JA_USADO', nome: nome, tipo: tipo, dataUso: linhas[i][8], erro: 'Ingresso já utilizado em ' + linhas[i][8] };
      }
      if (status === 'CANCELADO') {
        return { ok: false, codigo: 'CANCELADO', nome: nome, erro: 'Ingresso cancelado' };
      }

      // ATIVO → marca como USADO
      var agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
      sheet.getRange(i + 1, 7).setValue('USADO');
      sheet.getRange(i + 1, 9).setValue(agora);

      return { ok: true, codigo: 'LIBERADO', nome: nome, tipo: tipo, dataUso: agora };
    }
  }

  return { ok: false, codigo: 'NAO_ENCONTRADO', erro: 'Ingresso não encontrado' };
}

// ── Dashboard ────────────────────────────────────────────────
function getDashboard(data) {
  if (data.senha !== SENHA_PORTEIRO && data.senha !== SENHA_ORGANIZADOR) {
    return { ok: false, erro: 'Senha incorreta' };
  }

  var sheet  = getSheet();
  var linhas = sheet.getDataRange().getValues();
  var total = 0, pendentes = 0, ativos = 0, usados = 0, cancelados = 0;
  var porTipo = {};

  for (var i = 1; i < linhas.length; i++) {
    if (!linhas[i][0]) continue;
    total++;
    var status = linhas[i][6];
    var tipo   = linhas[i][3];
    if (status === 'PENDENTE')  pendentes++;
    if (status === 'ATIVO')     ativos++;
    if (status === 'USADO')     usados++;
    if (status === 'CANCELADO') cancelados++;
    porTipo[tipo] = (porTipo[tipo] || 0) + 1;
  }

  return { ok: true, total: total, pendentes: pendentes, ativos: ativos, usados: usados, cancelados: cancelados, porTipo: porTipo };
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

function getLimite(tipo) {
  var limites = { 'Inteira': 0, 'Meia': 0, 'VIP': 50, 'Criança': 0 };
  return limites[tipo] || 0;
}

function contarPorTipo(sheet, tipo) {
  var linhas = sheet.getDataRange().getValues();
  var count  = 0;
  for (var i = 1; i < linhas.length; i++) {
    if (linhas[i][3] === tipo && linhas[i][6] !== 'CANCELADO') count++;
  }
  return count;
}

function resposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ── Configurar planilha (rode uma vez manualmente) ───────────
function configurarPlanilha() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  var headers = ['ID', 'Nome', 'Contato', 'Tipo', 'Pagamento', 'Valor', 'Status', 'Data Compra', 'Data Uso', 'ID Pagamento', 'Observação'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  Logger.log('Planilha configurada!');
}
