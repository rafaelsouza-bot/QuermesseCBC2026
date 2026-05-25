// ============================================================
//  QUERMESSE — Backend (Google Apps Script)
//  Cole este código em: script.google.com → novo projeto
//  Depois: Implantar → Nova implantação → App da Web
//    - Executar como: Eu mesmo
//    - Quem tem acesso: Qualquer pessoa
// ============================================================

var SHEET_ID = 'COLE_O_ID_DA_SUA_PLANILHA_AQUI';
var SHEET_NAME = 'Ingressos';
var SENHA_PORTEIRO = '1234'; // Troque por uma senha segura

// ── Ponto de entrada HTTP ────────────────────────────────────
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var acao = data.acao;

    if (acao === 'comprar')    return resposta(comprarIngresso(data));
    if (acao === 'validar')    return resposta(validarIngresso(data));
    if (acao === 'dashboard')  return resposta(getDashboard(data));

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
  var id = gerarID();
  var agora = new Date();

  // Valida campos obrigatórios
  if (!data.nome || !data.contato || !data.tipo || !data.pagamento) {
    return { ok: false, erro: 'Campos obrigatórios faltando' };
  }

  // Verifica limite de ingressos por tipo
  var limite = getLimite(data.tipo);
  if (limite > 0) {
    var vendidos = contarPorTipo(sheet, data.tipo);
    if (vendidos >= limite) {
      return { ok: false, erro: 'Ingressos esgotados para ' + data.tipo };
    }
  }

  var linha = [
    id,                          // A - ID único
    data.nome,                   // B - Nome
    data.contato,                // C - WhatsApp ou e-mail
    data.tipo,                   // D - Tipo (inteira, meia, VIP)
    data.pagamento,              // E - Forma de pagamento
    data.valor || '',            // F - Valor pago
    'ATIVO',                     // G - Status
    Utilities.formatDate(agora, 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm'), // H - Data compra
    '',                          // I - Data uso
    data.paymentId || ''         // J - ID pagamento Mercado Pago
  ];

  sheet.appendRow(linha);

  return {
    ok: true,
    id: id,
    nome: data.nome,
    tipo: data.tipo,
    valor: data.valor,
    dataCompra: linha[7]
  };
}

// ── Validar ingresso (porteiro) ──────────────────────────────
function validarIngresso(data) {
  if (data.senha !== SENHA_PORTEIRO) {
    return { ok: false, erro: 'Senha do porteiro incorreta', codigo: 'SENHA' };
  }

  var sheet = getSheet();
  var id = data.id;
  var linhas = sheet.getDataRange().getValues();

  for (var i = 1; i < linhas.length; i++) {
    if (linhas[i][0] === id) {
      var status = linhas[i][6];
      var nome   = linhas[i][1];
      var tipo   = linhas[i][3];

      if (status === 'USADO') {
        var dataUso = linhas[i][8];
        return {
          ok: false,
          codigo: 'JA_USADO',
          nome: nome,
          tipo: tipo,
          dataUso: dataUso,
          erro: 'Ingresso já utilizado em ' + dataUso
        };
      }

      if (status === 'CANCELADO') {
        return { ok: false, codigo: 'CANCELADO', nome: nome, erro: 'Ingresso cancelado' };
      }

      // Marca como USADO
      var agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
      sheet.getRange(i + 1, 7).setValue('USADO');
      sheet.getRange(i + 1, 9).setValue(agora);

      return {
        ok: true,
        codigo: 'LIBERADO',
        nome: nome,
        tipo: tipo,
        dataUso: agora
      };
    }
  }

  return { ok: false, codigo: 'NAO_ENCONTRADO', erro: 'Ingresso não encontrado' };
}

// ── Dashboard resumo ─────────────────────────────────────────
function getDashboard(data) {
  if (data.senha !== SENHA_PORTEIRO) {
    return { ok: false, erro: 'Senha incorreta' };
  }

  var sheet = getSheet();
  var linhas = sheet.getDataRange().getValues();
  var total = 0, ativos = 0, usados = 0, cancelados = 0;
  var porTipo = {};

  for (var i = 1; i < linhas.length; i++) {
    if (!linhas[i][0]) continue;
    total++;
    var status = linhas[i][6];
    var tipo = linhas[i][3];
    if (status === 'ATIVO')     ativos++;
    if (status === 'USADO')     usados++;
    if (status === 'CANCELADO') cancelados++;
    porTipo[tipo] = (porTipo[tipo] || 0) + 1;
  }

  return { ok: true, total: total, ativos: ativos, usados: usados, cancelados: cancelados, porTipo: porTipo };
}

// ── Helpers ──────────────────────────────────────────────────
function getSheet() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
}

function gerarID() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var id = 'QRM-';
  for (var i = 0; i < 8; i++) {
    id += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return id;
}

function getLimite(tipo) {
  // 0 = sem limite. Ajuste conforme necessário.
  var limites = { 'Inteira': 0, 'Meia': 0, 'VIP': 50, 'Criança': 0 };
  return limites[tipo] || 0;
}

function contarPorTipo(sheet, tipo) {
  var linhas = sheet.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < linhas.length; i++) {
    if (linhas[i][3] === tipo && linhas[i][6] !== 'CANCELADO') count++;
  }
  return count;
}

function resposta(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Configurar planilha (rode uma vez manualmente) ───────────
function configurarPlanilha() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  var headers = ['ID', 'Nome', 'Contato', 'Tipo', 'Pagamento', 'Valor', 'Status', 'Data Compra', 'Data Uso', 'ID Pagamento'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1a1a2e')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 200);

  Logger.log('Planilha configurada!');
}
