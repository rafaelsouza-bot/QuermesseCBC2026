// ============================================================
//  QUERMESSE — Backend (Google Apps Script)
//  Cole este código em: Extensões → Apps Script (na planilha)
//  Depois: Implantar → Nova implantação → App da Web
//    - Executar como: Eu mesmo
//    - Quem tem acesso: Qualquer pessoa
// ============================================================

var SHEET_ID          = '174zZLsKEhbpcfLXtnWn2X8xJ0KQBA9LYkBEEbMTVC4I';
var SHEET_NAME        = 'Ingressos';
var SENHA_PORTEIRO    = 'Cbc@2026';
var SENHA_ORGANIZADOR = 'Org@2026';

// ── CONFIGURAÇÕES DE E-MAIL ──────────────────────────────────
// Troque pelo e-mail real da conta que hospeda o script:
var EMAIL_REMETENTE_NOME = 'Quermesse CBC';
var EMAIL_REPLY_TO       = 'SEU_EMAIL@escola.com.br'; // <-- TROQUE AQUI

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
    return resposta({ ok: false, erro: 'Acao desconhecida' });
  } catch (err) {
    return resposta({ ok: false, erro: err.message });
  }
}

function doGet(e) {
  return resposta({ ok: true, msg: 'API Quermesse no ar!' });
}

// ── Comprar ingresso (suporta multiplos) ─────────────────────
function comprarIngresso(data) {
  var sheet      = getSheet();
  var agora      = new Date();
  var quantidade = parseInt(data.quantidade) || 1;
  if (quantidade < 1 || quantidade > 10) return { ok: false, erro: 'Quantidade invalida' };
  if (!data.nome || !data.contato || !data.tipo || !data.pagamento) {
    return { ok: false, erro: 'Campos obrigatorios faltando' };
  }

  var ids          = [];
  var valorUnitario = parseFloat(data.valor) / quantidade;

  for (var i = 0; i < quantidade; i++) {
    var id = gerarID();
    ids.push(id);
    var linha = [
      id,                    // col A - ID
      data.nome,             // col B - Nome
      data.contato,          // col C - WhatsApp
      data.tipo,             // col D - Tipo
      data.pagamento,        // col E - Pagamento
      valorUnitario.toFixed(2), // col F - Valor
      'PENDENTE',            // col G - Status
      Utilities.formatDate(agora, 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm'), // col H - Data Compra
      '',                    // col I - Data Uso
      data.email || '',      // col J - E-mail
      quantidade > 1 ? (i+1) + '/' + quantidade : '' // col K - Seq/Obs
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

      if (status === 'PENDENTE')  return { ok: false, codigo: 'PENDENTE',    nome: nome, tipo: tipo, erro: 'Pagamento ainda nao confirmado pela organizacao' };
      if (status === 'USADO')     return { ok: false, codigo: 'JA_USADO',    nome: nome, tipo: tipo, dataUso: linhas[i][8], erro: 'Ingresso ja utilizado em ' + linhas[i][8] };
      if (status === 'CANCELADO') return { ok: false, codigo: 'CANCELADO',   nome: nome, erro: 'Ingresso cancelado' };

      var agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
      sheet.getRange(i + 1, 7).setValue('USADO');
      sheet.getRange(i + 1, 9).setValue(agora);

      return { ok: true, codigo: 'LIBERADO', nome: nome, tipo: tipo, seq: seq, dataUso: agora };
    }
  }
  return { ok: false, codigo: 'NAO_ENCONTRADO', erro: 'Ingresso nao encontrado' };
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
      id:        linhas[i][0],
      nome:      linhas[i][1],
      contato:   linhas[i][2],
      tipo:      linhas[i][3],
      pagamento: linhas[i][4],
      valor:     linhas[i][5],
      status:    status,
      dataCompra:linhas[i][7],
      dataUso:   linhas[i][8],
      email:     linhas[i][9] || '',
      obs:       linhas[i][10] || ''
    });
  }
  lista.reverse();
  return { ok: true, lista: lista, total: lista.length };
}

// ── Aprovar ──────────────────────────────────────────────────
function aprovarIngresso(data) {
  if (data.senha !== SENHA_ORGANIZADOR) return { ok: false, erro: 'Senha incorreta' };

  // Busca dados do ingresso antes de alterar
  var sheet      = getSheet();
  var linhas     = sheet.getDataRange().getValues();
  var ingresso   = null;

  for (var i = 1; i < linhas.length; i++) {
    if (linhas[i][0] === data.id) {
      ingresso = {
        id:        linhas[i][0],
        nome:      linhas[i][1],
        contato:   linhas[i][2],
        tipo:      linhas[i][3],
        valor:     linhas[i][5],
        dataCompra:linhas[i][7],
        email:     linhas[i][9] || ''
      };
      break;
    }
  }

  if (!ingresso) return { ok: false, erro: 'Ingresso nao encontrado' };

  // Muda status para ATIVO
  var resultado = mudarStatus(data.id, 'ATIVO', 'Aprovado pela organizacao');
  if (!resultado.ok) return resultado;

  // Envia e-mail com QR Code se houver e-mail cadastrado
  var emailEnviado = false;
  var avisoEmail   = '';
  if (ingresso.email) {
    try {
      enviarEmailQRCode(ingresso);
      emailEnviado = true;
    } catch(e) {
      avisoEmail = 'Erro ao enviar e-mail: ' + e.message;
    }
  }

  // Retorna dados para o painel disparar WhatsApp
  resultado.nome         = ingresso.nome;
  resultado.contato      = ingresso.contato;
  resultado.email        = ingresso.email;
  resultado.emailEnviado = emailEnviado;
  resultado.avisoEmail   = avisoEmail;
  return resultado;
}

// ── Enviar e-mail com QR Code de entrada ─────────────────────
function enviarEmailQRCode(ingresso) {
  var qrUrl  = 'https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=' + encodeURIComponent(ingresso.id);
  var valor  = parseFloat(ingresso.valor).toFixed(2).replace('.', ',');
  var assunto = 'Seu ingresso esta aprovado! | 104a Quermesse Solidaria CBC';

  var html =
    '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<style>' +
      'body{margin:0;padding:0;background:#f0f4f8;font-family:Segoe UI,Arial,sans-serif}' +
      '.wrap{max-width:520px;margin:32px auto;background:#fff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.10)}' +
      '.header{background:linear-gradient(135deg,#0a3550,#1a7aab);padding:32px 24px;text-align:center}' +
      '.header h1{color:#fff;margin:0 0 6px;font-size:22px;font-weight:900}' +
      '.header p{color:#7ec8e3;margin:0;font-size:14px}' +
      '.badge{display:inline-block;background:#29ABE2;color:#fff;font-size:11px;font-weight:800;letter-spacing:2px;text-transform:uppercase;padding:4px 14px;border-radius:99px;margin-bottom:12px}' +
      '.body{padding:28px 24px}' +
      '.ok-box{background:#e6fff3;border:1.5px solid #00c97a;border-radius:10px;padding:14px 18px;color:#007a45;font-weight:700;font-size:15px;margin-bottom:24px;text-align:center}' +
      '.qr-box{text-align:center;margin-bottom:20px}' +
      '.qr-box img{border-radius:12px;border:3px solid #0a3550;padding:8px;background:#fff}' +
      '.qr-dica{font-size:12px;color:#888;margin-top:8px}' +
      '.id-box{background:#f0f4f8;border-radius:8px;padding:10px 16px;font-family:monospace;font-size:20px;font-weight:900;letter-spacing:3px;color:#0a3550;text-align:center;margin-bottom:24px}' +
      'table{width:100%;border-collapse:collapse;margin-bottom:24px;font-size:14px}' +
      'td{padding:9px 4px;border-bottom:1px solid #f0f4f8;color:#444}' +
      'td:first-child{color:#888;width:42%}' +
      'td:last-child{font-weight:700;color:#0a3550}' +
      '.inst{background:#fff8e1;border-left:4px solid #f59e0b;border-radius:0 8px 8px 0;padding:16px 18px;margin-bottom:24px}' +
      '.inst h3{margin:0 0 10px;color:#92610a;font-size:13px;text-transform:uppercase;letter-spacing:1px}' +
      '.inst ol{margin:0;padding-left:18px;color:#6b4c0a;font-size:14px;line-height:2}' +
      '.aviso{background:#fff3cd;border:1px solid #ffc107;border-radius:8px;padding:12px 16px;font-size:13px;color:#856404;margin-bottom:24px;text-align:center}' +
      '.footer{background:#f8f9fa;padding:18px 24px;text-align:center;font-size:12px;color:#aaa;border-top:1px solid #eee}' +
    '</style></head><body>' +
    '<div class="wrap">' +
      '<div class="header">' +
        '<div class="badge">Quermesse 2026</div>' +
        '<h1>Ingresso Aprovado!</h1>' +
        '<p>104a Quermesse Solidaria CBC &mdash; 04 de Julho de 2026</p>' +
      '</div>' +
      '<div class="body">' +
        '<div class="ok-box">Pagamento confirmado pela organizacao!</div>' +
        '<div class="qr-box">' +
          '<img src="' + qrUrl + '" width="220" height="220" alt="QR Code de Entrada"><br>' +
          '<div class="qr-dica">Apresente este QR Code ao porteiro na entrada</div>' +
        '</div>' +
        '<div class="id-box">' + ingresso.id + '</div>' +
        '<table>' +
          '<tr><td>Nome</td><td>' + ingresso.nome + '</td></tr>' +
          '<tr><td>Tipo</td><td>' + ingresso.tipo + '</td></tr>' +
          '<tr><td>Valor pago</td><td>R$ ' + valor + '</td></tr>' +
          '<tr><td>Data da compra</td><td>' + ingresso.dataCompra + '</td></tr>' +
          '<tr><td>Evento</td><td>04/07/2026 &mdash; 11h as 16h</td></tr>' +
          '<tr><td>Local</td><td>Ramiro Barcelos, 996 &mdash; POA/RS</td></tr>' +
        '</table>' +
        '<div class="inst">' +
          '<h3>Como usar seu ingresso no dia</h3>' +
          '<ol>' +
            '<li>Abra este e-mail no celular ao chegar</li>' +
            '<li>Apresente o <strong>QR Code acima</strong> ao porteiro</li>' +
            '<li>Aguarde o porteiro escanear &mdash; a entrada sera liberada</li>' +
            '<li>Apos a leitura o ingresso e marcado como utilizado</li>' +
            '<li><strong>Cada QR Code e de uso unico</strong> &mdash; nao compartilhe</li>' +
          '</ol>' +
        '</div>' +
        '<div class="aviso">Guarde este e-mail! Ele e o seu ingresso de entrada.</div>' +
      '</div>' +
      '<div class="footer">' +
        '104a Quermesse Solidaria CBC &bull; 04 de Julho de 2026<br>' +
        'Ramiro Barcelos, 996 &mdash; Porto Alegre/RS<br><br>' +
        'Duvidas? Responda este e-mail.' +
      '</div>' +
    '</div>' +
    '</body></html>';

  MailApp.sendEmail({
    to:       ingresso.email,
    subject:  assunto,
    htmlBody: html,
    name:     EMAIL_REMETENTE_NOME,
    replyTo:  EMAIL_REPLY_TO
  });
}

// ── Rejeitar ─────────────────────────────────────────────────
function rejeitarIngresso(data) {
  if (data.senha !== SENHA_ORGANIZADOR) return { ok: false, erro: 'Senha incorreta' };
  return mudarStatus(data.id, 'CANCELADO', data.obs || 'Rejeitado pela organizacao');
}

// ── Cancelar ─────────────────────────────────────────────────
function cancelarIngresso(data) {
  if (data.senha !== SENHA_ORGANIZADOR) return { ok: false, erro: 'Senha incorreta' };
  return mudarStatus(data.id, 'CANCELADO', 'Cancelado pela organizacao');
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
  return { ok: false, erro: 'Ingresso nao encontrado' };
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
  var headers = ['ID','Nome','WhatsApp','Tipo','Pagamento','Valor','Status','Data Compra','Data Uso','E-mail','Seq/Obs'];
  sheet.getRange(1,1,1,headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#0a3550').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  Logger.log('Planilha configurada!');
}
