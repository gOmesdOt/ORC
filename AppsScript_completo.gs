// ============================================================
// LG Visual Media — App Orçamentos
// Google Apps Script — Backend
// ============================================================

const SPREADSHEET_ID = '1SFfJ7IJgtMiOUQdplUaM2G3nXlalnBy3iqI55b7msQg';
const DRIVE_FOLDER_ID = '1Arp6FIA2wCaNU5meZjnhq8crCwRSdrhg';
const CONDICOES_FILE_ID = '1fx7ZJm8mBgXvbAKEpt5FNMrMseG3CWTA';
const NOTION_TOKEN = 'COLA_AQUI_O_TEU_TOKEN_NOTION'; // NUNCA colocar a chave real aqui se este ficheiro for para o GitHub — cola a chave real apenas dentro do editor do Apps Script
const NOTION_CLIENTES_DS = '32f027b4-6dc5-80ef-a72b-000bc07bbfdb';
const APP_TOKEN = 'psxAMJ88yrsbXXKe5osvo789iHJ6JxvhbkBP1oPn14k'; // tem de ser igual ao APP_TOKEN no index.html — protege o script de pedidos externos

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);

    if (payload.token !== APP_TOKEN) {
      return resposta({ ok: false, erro: 'Não autorizado.' });
    }

    const action = payload.action;
    let result;

    if (action === 'guardarOrcamento')        result = guardarOrcamento(payload.data);
    else if (action === 'listarOrcamentos')   result = listarOrcamentos(payload.ano);
    else if (action === 'atualizarEstado')    result = atualizarEstado(payload.id, payload.estado, payload.ano);
    else if (action === 'uploadPDF')          result = uploadPDF(payload.nome, payload.pdfBase64);
    else if (action === 'prepararEmail')      result = prepararEmail(payload.data);
    else if (action === 'criarProjetoNotion') result = criarProjetoNotion(payload.data);
    else if (action === 'listarClientes')     result = listarClientes();
    else if (action === 'gerarPDFServidor')   result = gerarPDFServidor(payload.nome, payload.html);
    else result = { ok: false, erro: 'Ação desconhecida: ' + action };

    return resposta(result);
  } catch (err) {
    return resposta({ ok: false, erro: err.toString() });
  }
}

function doGet(e) {
  return ContentService.createTextOutput(
    JSON.stringify({ ok: true, msg: 'LG Visual Media Script ativo' })
  ).setMimeType(ContentService.MimeType.JSON);
}

// ------------------------------------------------------------
// GERAR PDF NO SERVIDOR (via PDFShift - renderizacao fiel)
// ------------------------------------------------------------
const PDFSHIFT_API_KEY = 'COLA_AQUI_A_TUA_CHAVE_PDFSHIFT'; // NUNCA colocar a chave real aqui se este ficheiro for para o GitHub — cola a chave real apenas dentro do editor do Apps Script

function gerarPDFServidor(nomeBase, htmlContent) {
  try {
    var payload = JSON.stringify({
      source: htmlContent,
      format: '680xauto',
      margin: '0',
      use_print: false,
      sandbox: false
    });

    var options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode('api:' + PDFSHIFT_API_KEY)
      },
      payload: payload,
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch('https://api.pdfshift.io/v3/convert/pdf', options);
    var code = response.getResponseCode();

    if (code !== 200) {
      return { ok: false, erro: 'PDFShift respondeu ' + code + ': ' + response.getContentText().substring(0, 300) };
    }

    var pdfBlob = response.getBlob();
    pdfBlob.setName(nomeBase + '.pdf');

    // versionar e guardar na pasta de orçamentos
    var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    var nomeCompleto = nomeBase + '.pdf';
    var existentes = folder.getFilesByName(nomeCompleto);
    var versao = 1;
    while (existentes.hasNext()) { existentes.next(); versao++; }
    var nomeFinal = versao > 1 ? nomeBase + '-v' + versao + '.pdf' : nomeCompleto;

    pdfBlob.setName(nomeFinal);
    var file = folder.createFile(pdfBlob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    var pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());

    return {
      ok: true,
      nomeFinal: nomeFinal.replace('.pdf', ''),
      linkPDF: file.getUrl(),
      pdfBase64: pdfBase64
    };
  } catch (err) {
    return { ok: false, erro: 'Erro PDFShift: ' + err.toString() };
  }
}

// ------------------------------------------------------------
// GUARDAR ORÇAMENTO NO SHEETS
// ------------------------------------------------------------
function guardarOrcamento(d) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const ano = d.ano || new Date().getFullYear().toString();
  let sheet = ss.getSheetByName(ano);

  if (!sheet) {
    sheet = ss.insertSheet(ano);
    const headers = [
      'ID', 'Nº Orçamento', 'Versão', 'Data', 'Validade (dias)',
      'Nome Projeto (interno)', 'Data Realização',
      'Empresa Cliente', 'NIF Cliente', 'Morada Cliente',
      'Telefone Cliente', 'Email Cliente',
      'Descrição Projeto', 'Sinal (%)', 'Sinal (€)',
      'Serviços (JSON)', 'Total (€)',
      'Nota Prazo', 'Nota Revisões', 'Nota Pagamento', 'Nota Extra',
      'Estado', 'Link PDF Drive', 'Data Criação'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#F5A800').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  const id = d.id || Date.now().toString();
  const agora = new Date().toISOString();

  const row = [
    id, d.num, d.versao || 'v1', d.data, d.validade,
    d.nomeProjeto, d.dataRealizacao || '',
    d.cliEmpresa, d.cliNif || '', d.cliMorada || '',
    d.cliTel || '', d.cliEmail || '',
    d.projDesc || '', d.sinalPct || 0, d.sinalEur || 0,
    JSON.stringify(d.linhas || []), d.total,
    d.notaPrazo || '', d.notaRev || '', d.notaPag || '', d.notaExtra || '',
    'Pendente', d.linkPDF || '', agora
  ];

  sheet.appendRow(row);
  return { ok: true, id };
}

// ------------------------------------------------------------
// LISTAR ORÇAMENTOS DE UM ANO
// ------------------------------------------------------------
function listarOrcamentos(ano) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(ano || new Date().getFullYear().toString());
  if (!sheet) return { ok: true, orcamentos: [] };

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { ok: true, orcamentos: [] };

  const headers = data[0];
  const orcamentos = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  }).reverse();

  return { ok: true, orcamentos };
}

// ------------------------------------------------------------
// ATUALIZAR ESTADO DE UM ORÇAMENTO
// ------------------------------------------------------------
function atualizarEstado(id, novoEstado, ano) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(ano || new Date().getFullYear().toString());
  if (!sheet) return { ok: false, erro: 'Aba não encontrada' };

  const data = sheet.getDataRange().getValues();
  const colId = 0;
  const colEstado = 21;

  for (let i = 1; i < data.length; i++) {
    if (data[i][colId].toString() === id.toString()) {
      sheet.getRange(i + 1, colEstado + 1).setValue(novoEstado);
      return { ok: true };
    }
  }
  return { ok: false, erro: 'Orçamento não encontrado' };
}

// ------------------------------------------------------------
// UPLOAD DIRETO DE PDF PARA A DRIVE (mantido para compatibilidade)
// ------------------------------------------------------------
function uploadPDF(nomeBase, pdfBase64) {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const nomeCompleto = nomeBase + '.pdf';

  const existentes = folder.getFilesByName(nomeCompleto);
  let versao = 1;
  while (existentes.hasNext()) { existentes.next(); versao++; }

  const nomeFinal = versao > 1 ? nomeBase + '-v' + versao + '.pdf' : nomeCompleto;

  const bytes = Utilities.base64Decode(pdfBase64);
  const blob = Utilities.newBlob(bytes, 'application/pdf', nomeFinal);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return { ok: true, nomeFinal, fileId: file.getId(), linkPDF: file.getUrl() };
}

// ------------------------------------------------------------
// PREPARAR RASCUNHO NO GMAIL
// ------------------------------------------------------------
function prepararEmail(d) {
  const assunto = 'Orçamento ' + d.nomeProjeto + ' para ' + d.cliEmpresa + ' - Luís Gomes';

  const hora = parseInt(Utilities.formatDate(new Date(), 'Europe/Lisbon', 'H'));
  let saudacao = 'Bom dia,';
  if (hora >= 13 && hora < 20) saudacao = 'Boa tarde,';
  else if (hora >= 20 || hora < 6) saudacao = 'Boa noite,';

  const dataReal = d.dataRealizacao ? 'com data de realização ' + d.dataRealizacao + ',' : '';

  const corpoTexto =
    saudacao + '\n\n' +
    'Venho por este meio enviar a proposta de orçamento ' + d.nomeProjeto +
    ', ' + dataReal + ' conforme solicitado.\n\n' +
    'Caso exista alguma questão ou necessidade de revisão, encontro-me ao dispor.\n\n' +
    'Com os melhores cumprimentos,';

  let assinatura = '';
  try {
    var sendAsList = Gmail.Users.Settings.SendAs.list('me').sendAs;
    var principal = sendAsList.filter(function(s){ return s.isDefault; })[0] || sendAsList[0];
    if (principal && principal.signature) assinatura = principal.signature;
  } catch (err) {
    console.log('Erro ao buscar assinatura:', err);
  }

  const corpoHtml = corpoTexto.replace(/\n/g, '<br>') + (assinatura ? '<br><br>' + assinatura : '');

  const anexos = [];
  if (d.pdfBase64) {
    const bytes = Utilities.base64Decode(d.pdfBase64);
    const blob = Utilities.newBlob(bytes, 'application/pdf', d.nomePDF || 'Orcamento.pdf');
    anexos.push(blob);
  }
  try {
    const condFile = DriveApp.getFileById(CONDICOES_FILE_ID);
    anexos.push(condFile.getBlob());
  } catch (err) {
    console.log('Erro ao anexar condições gerais:', err);
  }

  GmailApp.createDraft(d.cliEmail || '', assunto, corpoTexto, { htmlBody: corpoHtml, attachments: anexos });

  return { ok: true, assunto };
}

// ------------------------------------------------------------
// CRIAR PROJETO NO NOTION
// ------------------------------------------------------------
function criarProjetoNotion(d) {
  var headers = {
    'Authorization': 'Bearer ' + NOTION_TOKEN,
    'Notion-Version': '2022-06-28',
    'Content-Type': 'application/json'
  };

  var anoCurto = d.ano.toString().slice(-2);
  var searchRes = UrlFetchApp.fetch('https://api.notion.com/v1/search', {
    method: 'post', headers: headers,
    payload: JSON.stringify({ query: "Projetos '" + anoCurto, filter: { property: 'object', value: 'database' } }),
    muteHttpExceptions: true
  });
  var searchData = JSON.parse(searchRes.getContentText());
  if (!searchData.results || !searchData.results.length) {
    return { ok: false, erro: "Base 'Projetos '" + anoCurto + "' não encontrada. Confirma que está partilhada com a integração." };
  }
  var dbId = searchData.results[0].id;

  var clienteRelId = null;
  var cliRes = UrlFetchApp.fetch('https://api.notion.com/v1/data_sources/' + NOTION_CLIENTES_DS + '/query', {
    method: 'post', headers: headers,
    payload: JSON.stringify({
      filter: { or: [
        { property: 'Empresa', rich_text: { contains: d.clienteNome } },
        { property: 'Nome', title: { contains: d.clienteNome } }
      ]},
      page_size: 1
    }),
    muteHttpExceptions: true
  });
  var cliData = JSON.parse(cliRes.getContentText());
  if (cliData.results && cliData.results.length) clienteRelId = cliData.results[0].id;

  var properties = {
    'Nome do projeto': { title: [{ text: { content: d.nomeProjeto || 'Sem nome' } }] },
    'Estado': { status: { name: 'Pendente' } },
    'Prazo de entrega': { select: { name: d.prazoEntrega || '5 dias' } },
    'Preço': { number: parseFloat(d.preco) || 0 }
  };
  if (d.dataInicio) properties['Data Inicio'] = { date: { start: d.dataInicio } };
  if (clienteRelId) properties['Cliente'] = { relation: [{ id: clienteRelId }] };

  var createRes = UrlFetchApp.fetch('https://api.notion.com/v1/pages', {
    method: 'post', headers: headers,
    payload: JSON.stringify({ parent: { database_id: dbId }, properties: properties }),
    muteHttpExceptions: true
  });
  var createData = JSON.parse(createRes.getContentText());

  if (createData.id) {
    var avisoCliente = clienteRelId ? '' : ' (cliente não encontrado — projeto criado sem ligação)';
    return { ok: true, url: createData.url, pageId: createData.id, aviso: avisoCliente };
  } else {
    return { ok: false, erro: createData.message || 'Erro desconhecido ao criar página.' };
  }
}

// ------------------------------------------------------------
// LISTAR CLIENTES DO NOTION (Info Clientes)
// ------------------------------------------------------------
function listarClientes() {
  var headers = {
    'Authorization': 'Bearer ' + NOTION_TOKEN,
    'Notion-Version': '2022-06-28',
    'Content-Type': 'application/json'
  };

  var clientes = [];
  var hasMore = true;
  var startCursor = null;

  while (hasMore) {
    var payload = { page_size: 100 };
    if (startCursor) payload.start_cursor = startCursor;

    var res = UrlFetchApp.fetch('https://api.notion.com/v1/data_sources/' + NOTION_CLIENTES_DS + '/query', {
      method: 'post', headers: headers, payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    var data = JSON.parse(res.getContentText());

    if (code !== 200) {
      return { ok: false, erro: 'Notion respondeu ' + code + ': ' + (data.message || JSON.stringify(data)) };
    }
    if (!data.results) break;

    data.results.forEach(function(page) {
      var p = page.properties;
      clientes.push({
        nome: getTitle(p['Nome']),
        empresa: getText(p['Empresa']),
        nif: p['NIF | Contribuinte'] && p['NIF | Contribuinte'].number ? String(p['NIF | Contribuinte'].number) : '',
        morada: getText(p['Morada']),
        tel: p['Telemovel'] && p['Telemovel'].phone_number ? p['Telemovel'].phone_number : '',
        email: p['E-mail'] && p['E-mail'].email ? p['E-mail'].email : ''
      });
    });

    hasMore = data.has_more;
    startCursor = data.next_cursor;
  }

  return { ok: true, clientes: clientes, total: clientes.length };
}

function getTitle(prop) {
  if (!prop || !prop.title || !prop.title.length) return '';
  return prop.title.map(function(t) { return t.plain_text; }).join('');
}
function getText(prop) {
  if (!prop || !prop.rich_text || !prop.rich_text.length) return '';
  return prop.rich_text.map(function(t) { return t.plain_text; }).join('');
}

// ------------------------------------------------------------
// UTILITÁRIO — formata resposta HTTP
// ------------------------------------------------------------
function resposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
