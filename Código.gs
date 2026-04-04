// ============================================================
// AGÊNCIA CONTÁBIL IA — Código.gs v3
// Matheus Saldanha Garcia | OAB/SP 411.209
// ============================================================

const SHEETS = {
  dimatteo:    { id: '14DklR1iIdIdUE5HpheixbBBRUNvyPKsSHPX7BRKVANw', name: 'DiMatteo Açaí' },
  containers:  { id: '1PEZ_M3dFc4AVc9f4zHXlY1xKj360zbtXUZqOHPcaeiM', name: 'Containers Hamburgueria' },
  pizzaria:    { id: '14qDlPqYgi0mf7e1-iyE399NKVUun9oG_-mVDYfq5RVk', name: 'Santa Massa Pizzaria' },
  restaurante: { id: '1mC1l6-BB1oHdSoAOFfH2XI9s02KrOUOfoyb6bzakMIs', name: 'Santa Massa Restaurante' },
  riopreto:    { id: '1ybDbTOIjDNaYWjct-fFFouEPdV2FLdvRIMk04bhrfA8', name: 'Santa Massa Rio Preto' },
};

const CATEGORIAS = [
  'Receita Bruta', 'CMV', 'Custo com Pessoal', 'Embalagens',
  'Higiene e Limpeza', 'Consumo Interno', 'Marketing / Patrocínio',
  'Despesas Operacionais', 'Despesas Fixas', 'Despesas Diversas', 'Manutenção',
];

// IDs das pastas Drive por empresa (preenchidos via criarPastasDrive())
const DRIVE_FOLDERS = {
  containers:  '',
  dimatteo:    '',
  pizzaria:    '',
  restaurante: '',
  riopreto:    '',
};

// ============================================================
// ENTRY POINTS
// ============================================================
function doGet(e) {
  const p = e.parameter;
  try {
    if (p.action === 'ping')          return jsonResponse({ status: 'ok', message: 'API Agência Contábil v3 ativa' });
    if (p.action === 'getLancamentos') return jsonResponse({ status: 'ok', data: getLancamentos(p.empresa) });
    if (p.action === 'getDRE')         return jsonResponse({ status: 'ok', data: getDRE(p.empresa, p.mes, p.ano) });
    if (p.action === 'getConsolidado') return jsonResponse({ status: 'ok', data: getConsolidado() });
    if (p.action === 'getCategorias')  return jsonResponse({ status: 'ok', data: CATEGORIAS });
    if (p.action === 'getFornecedores') return jsonResponse({ status: 'ok', data: getFornecedores() });
    return jsonResponse({ status: 'error', message: 'Ação não reconhecida.' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    switch (body.action) {
      case 'ping':
        return jsonResponse({ status: 'ok' });

      case 'addLancamento': {
        const r = addLancamento(body.empresa, body.lancamento);
        atualizarDREFromLancamentos(body.empresa, parseInt(body.lancamento.mes), parseInt(body.lancamento.ano));
        return jsonResponse({ status: 'ok', result: r });
      }

      case 'editarLancamento': {
        const r = editarLancamento(body.empresa, body.id, body.lancamento);
        atualizarDREFromLancamentos(body.empresa, parseInt(body.lancamento.mes), parseInt(body.lancamento.ano));
        return jsonResponse({ status: 'ok', result: r });
      }

      case 'deletarLancamento': {
        const r = deletarLancamento(body.empresa, body.id);
        return jsonResponse({ status: 'ok', result: r });
      }

      case 'recalcularDRE': {
        atualizarDREFromLancamentos(body.empresa, parseInt(body.mes), parseInt(body.ano));
        return jsonResponse({ status: 'ok', message: 'DRE recalculada.' });
      }

      case 'auditarIds': {
        const r = auditarIds();
        return jsonResponse({ status: 'ok', totalCorrigidos: r });
      }

      case 'uploadNota': {
        const r = uploadNotaFiscal(body.empresa, body.fileName, body.base64Data, body.mimeType, body.descricao, body.fornecedor);
        return jsonResponse({ status: 'ok', result: r });
      }

      // --------------------------------------------------------
      // NOVO: Proxy Anthropic — chave nunca vai ao frontend
      // --------------------------------------------------------
      case 'proxyAnthropic': {
        const r = proxyAnthropic(body);
        return jsonResponse(r);
      }

      // --------------------------------------------------------
      // NOVO: Salvar chave Anthropic nas PropertiesService
      // --------------------------------------------------------
      case 'salvarAnthropicKey': {
        if (!body.key || body.key.length < 10) {
          return jsonResponse({ status: 'error', message: 'Chave inválida.' });
        }
        PropertiesService.getScriptProperties().setProperty('ANTHROPIC_KEY', body.key);
        return jsonResponse({ status: 'ok', message: 'Chave salva com sucesso.' });
      }

      case 'getFornecedores': {
        return jsonResponse({ status: 'ok', data: getFornecedores() });
      }

      default:
        return jsonResponse({ status: 'error', message: 'Ação não reconhecida: ' + body.action });
    }
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

// ============================================================
// PROXY ANTHROPIC
// ============================================================
function proxyAnthropic(body) {
  // — 1. VERIFICAÇÕES PRÉ-CHAMADA —
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_KEY');
  const keyOk  = !!key;
  const keyFmt = keyOk && key.startsWith('sk-ant-');
  Logger.log('[proxy] key existe=' + keyOk + ' formato_valido=' + keyFmt
             + ' prefixo=' + (keyOk ? key.substring(0, 14) + '...' : 'N/A'));

  if (!keyOk)  return { status: 'error', message: 'ANTHROPIC_KEY não encontrada nas PropertiesService.' };
  if (!keyFmt) return { status: 'error', message: 'ANTHROPIC_KEY com formato inválido (esperado: sk-ant-...). Prefixo recebido: ' + key.substring(0, 10) };

  const model    = body.model    || 'claude-haiku-4-5-20251001';
  const maxTok   = body.max_tokens || 800;
  const messages = body.messages;
  const system   = body.system;

  Logger.log('[proxy] model=' + model + ' max_tokens=' + maxTok);
  Logger.log('[proxy] messages existe=' + !!messages + ' qtd=' + (Array.isArray(messages) ? messages.length : 'N/A'));
  Logger.log('[proxy] system existe=' + !!system);

  if (!messages || !Array.isArray(messages) || messages.length === 0) {
    return { status: 'error', message: 'body.messages ausente ou vazio.' };
  }

  // — 2. MONTAR PAYLOAD —
  const payload = { model, max_tokens: maxTok, messages };
  if (system) payload.system = system;

  const payloadStr = JSON.stringify(payload);
  Logger.log('[proxy] payload bytes=' + payloadStr.length);

  // — 3. CHAMADA URLFETCHAPP —
  let res;
  try {
    res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type':      'application/json',
        'x-api-key':         key,
        'anthropic-version': '2023-06-01',
      },
      payload: payloadStr,
      muteHttpExceptions: true,
    });
  } catch (fetchErr) {
    Logger.log('[proxy] ERRO no UrlFetchApp.fetch: ' + fetchErr.message);
    return { status: 'error', message: 'UrlFetchApp falhou: ' + fetchErr.message };
  }

  // — 4. TRATAR RESPOSTA —
  const httpCode   = res.getResponseCode();
  const rawBody    = res.getContentText();
  Logger.log('[proxy] httpCode=' + httpCode + ' responseBody=' + rawBody.substring(0, 500));

  if (httpCode !== 200) {
    return {
      status:       'error',
      message:      'Anthropic retornou HTTP ' + httpCode,
      httpCode:     httpCode,
      responseBody: rawBody.substring(0, 500),
    };
  }

  try {
    const data = JSON.parse(rawBody);
    return { status: 'ok', data };
  } catch (parseErr) {
    Logger.log('[proxy] ERRO ao parsear JSON da resposta: ' + parseErr.message);
    return { status: 'error', message: 'Resposta da Anthropic não é JSON válido: ' + rawBody.substring(0, 200) };
  }
}

// ============================================================
// LEITURA
// ============================================================
function getLancamentos(empresaKey) {
  return sheetToJson(getSheet(empresaKey, 'Lançamentos'));
}

function getDRE(empresaKey, mes, ano) {
  const rows = sheetToJson(getSheet(empresaKey, 'DRE'));
  if (!mes && !ano) return rows;
  return rows.filter(r => (!mes || String(r['Mês']) === String(mes)) && (!ano || String(r['Ano']) === String(ano)));
}

function getConsolidado() {
  const result = {};
  for (const key in SHEETS) {
    try { result[key] = sheetToJson(getSheet(key, 'DRE')); }
    catch (e) { result[key] = []; }
  }
  return result;
}

// ============================================================
// HISTÓRICO DE FORNECEDORES (item 4)
// ============================================================
function getFornecedores() {
  const mapa = {}; // chave: nome normalizado → {nome, cnpj, frequencia, ultimaCategoria, empresas}

  for (const empKey in SHEETS) {
    let lancs = [];
    try { lancs = sheetToJson(getSheet(empKey, 'Lançamentos')); } catch (e) { continue; }

    for (const l of lancs) {
      const nome = (l['Fornecedor'] || '').toString().trim();
      if (!nome || nome === '—') continue;
      const chave = nome.toLowerCase().replace(/\s+/g, ' ');
      if (!mapa[chave]) {
        mapa[chave] = { nome, cnpj: '', frequencia: 0, ultimaCategoria: '', empresas: [] };
      }
      mapa[chave].frequencia++;
      mapa[chave].ultimaCategoria = l['Categoria'] || mapa[chave].ultimaCategoria;
      if (!mapa[chave].empresas.includes(empKey)) mapa[chave].empresas.push(empKey);
      // CNPJ armazenado na coluna Fornecedor como "Nome [XX.XXX.XXX/XXXX-XX]" — extrai se presente
      const cnpjMatch = nome.match(/\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}/);
      if (cnpjMatch && !mapa[chave].cnpj) mapa[chave].cnpj = cnpjMatch[0];
    }
  }

  return Object.values(mapa).sort((a, b) => b.frequencia - a.frequencia);
}

// ============================================================
// ESCRITA
// ============================================================
function addLancamento(empresaKey, lanc) {
  const sheet = getSheet(empresaKey, 'Lançamentos');
  const id = sheet.getLastRow(); // header=1, so id=rows after header
  const now = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');
  const row = [id, lanc.data, lanc.categoria, lanc.descricao, lanc.fornecedor || '—',
               parseFloat(lanc.valor) || 0, parseInt(lanc.mes), parseInt(lanc.ano), now];
  sheet.appendRow(row);
  return { id };
}

function editarLancamento(empresaKey, id, lanc) {
  const sheet = getSheet(empresaKey, 'Lançamentos');
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      const now = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');
      sheet.getRange(i + 1, 1, 1, 9).setValues([[
        data[i][0], lanc.data, lanc.categoria, lanc.descricao,
        lanc.fornecedor || '—', parseFloat(lanc.valor) || 0,
        parseInt(lanc.mes), parseInt(lanc.ano), now,
      ]]);
      return { updated: true };
    }
  }
  throw new Error('Lançamento #' + id + ' não encontrado.');
}

function deletarLancamento(empresaKey, id) {
  const sheet = getSheet(empresaKey, 'Lançamentos');
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { deleted: true };
    }
  }
  throw new Error('Lançamento #' + id + ' não encontrado.');
}

function auditarIds() {
  let totalCorrigidos = 0;
  for (const empKey in SHEETS) {
    try {
      const sheet = getSheet(empKey, 'Lançamentos');
      const data  = sheet.getDataRange().getValues();
      if (data.length < 2) continue;

      // Detecta IDs duplicados
      const seen = {};
      const dups = [];
      for (let i = 1; i < data.length; i++) {
        const id = String(data[i][0]);
        if (seen[id] !== undefined) { dups.push(i); }
        else { seen[id] = i; }
      }

      if (dups.length === 0) continue;

      // Renumera a partir do maior ID único + 1
      const maxId = Math.max(...Object.keys(seen).map(Number).filter(n => !isNaN(n)));
      let next = maxId + 1;
      for (const rowIdx of dups) {
        sheet.getRange(rowIdx + 1, 1).setValue(next++);
        totalCorrigidos++;
      }
    } catch (e) {
      Logger.log('Erro ao auditar ' + empKey + ': ' + e.message);
    }
  }
  return totalCorrigidos;
}

// ============================================================
// ATUALIZAR DRE A PARTIR DOS LANÇAMENTOS
// ============================================================
function atualizarDREFromLancamentos(empresaKey, mes, ano) {
  const ss      = SpreadsheetApp.openById(SHEETS[empresaKey].id);
  const abaLanc = ss.getSheetByName('Lançamentos');
  const abaDRE  = ss.getSheetByName('DRE');
  if (!abaLanc || !abaDRE) return;

  const dados = abaLanc.getDataRange().getValues();
  const totais = {};
  CATEGORIAS.forEach(c => (totais[c] = 0));

  for (let i = 1; i < dados.length; i++) {
    if (parseInt(dados[i][6]) === mes && parseInt(dados[i][7]) === ano) {
      const cat = dados[i][2];
      if (totais[cat] !== undefined) totais[cat] += parseFloat(dados[i][5]) || 0;
    }
  }

  const receita     = totais['Receita Bruta'];
  const cmv         = totais['CMV'];
  const lucroBruto  = receita - cmv;
  const cmvPct      = receita > 0 ? parseFloat((cmv / receita * 100).toFixed(2)) : 0;
  const lbPct       = receita > 0 ? parseFloat((lucroBruto / receita * 100).toFixed(2)) : 0;
  const pessoal     = totais['Custo com Pessoal'];
  const embalagens  = totais['Embalagens'];
  const higiene     = totais['Higiene e Limpeza'];
  const consumo     = totais['Consumo Interno'];
  const marketing   = totais['Marketing / Patrocínio'];
  const operacional = totais['Despesas Operacionais'];
  const fixas       = totais['Despesas Fixas'];
  const diversas    = totais['Despesas Diversas'];
  const manutencao  = totais['Manutenção'];
  const totalDesp   = pessoal + embalagens + higiene + consumo + marketing + operacional + fixas + diversas + manutencao;
  const tdPct       = receita > 0 ? parseFloat((totalDesp / receita * 100).toFixed(2)) : 0;
  const ebitda      = lucroBruto - totalDesp;
  const margemPct   = receita > 0 ? parseFloat((ebitda / receita * 100).toFixed(2)) : 0;

  const linha = [mes, ano, receita, cmv, cmvPct, lucroBruto, lbPct, pessoal, embalagens,
                 higiene, consumo, marketing, operacional, fixas, diversas, manutencao,
                 totalDesp, tdPct, ebitda, margemPct];

  const dreData = abaDRE.getDataRange().getValues();
  for (let i = 1; i < dreData.length; i++) {
    if (parseInt(dreData[i][0]) === mes && parseInt(dreData[i][1]) === ano) {
      abaDRE.getRange(i + 1, 1, 1, linha.length).setValues([linha]);
      return;
    }
  }
  abaDRE.appendRow(linha);
}

// ============================================================
// UPLOAD DRIVE (estrutura: Empresa / Fornecedor / arquivo)
// ============================================================
function uploadNotaFiscal(empresaKey, fileName, base64Data, mimeType, descricao, fornecedor) {
  let folderId = DRIVE_FOLDERS[empresaKey];
  if (!folderId) {
    const ids = criarPastasDrive();
    folderId = ids[empresaKey];
    if (!folderId) throw new Error('Pasta Drive não encontrada: ' + empresaKey);
  }

  const pastaEmpresa = DriveApp.getFolderById(folderId);
  const pastaForn    = getOuCriarPasta(pastaEmpresa, fornecedor || '_Sem Fornecedor');

  const ext     = mimeType === 'application/pdf' ? '.pdf' : mimeType === 'image/png' ? '.png' : '.jpg';
  const nomeArq = fileName.includes('.') ? fileName : fileName + ext;
  const blob    = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, nomeArq);
  const file    = pastaForn.createFile(blob);
  const now     = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');
  file.setDescription(descricao || 'NF ' + sanitizarNomeDrive(fornecedor) + ' — ' + now);

  return {
    fileId:    file.getId(),
    fileUrl:   file.getUrl(),
    folderUrl: pastaForn.getUrl(),
    pasta:     pastaEmpresa.getName() + ' / ' + sanitizarNomeDrive(fornecedor),
  };
}

function criarPastasDrive() {
  const root  = DriveApp.getRootFolder();
  const iter  = root.getFoldersByName('Agência Contábil IA — Notas Fiscais');
  const raiz  = iter.hasNext() ? iter.next() : root.createFolder('Agência Contábil IA — Notas Fiscais');
  const nomes = { containers: 'Containers Hamburgueria', dimatteo: 'DiMatteo Açaí',
                  pizzaria: 'Santa Massa Pizzaria', restaurante: 'Santa Massa Restaurante',
                  riopreto: 'Santa Massa Rio Preto' };
  const ids = {};
  for (const k in nomes) {
    const sub = raiz.getFoldersByName(nomes[k]);
    ids[k] = (sub.hasNext() ? sub.next() : raiz.createFolder(nomes[k])).getId();
  }
  return ids;
}

function getOuCriarPasta(pai, nome) {
  const n  = sanitizarNomeDrive(nome);
  const it = pai.getFoldersByName(n);
  return it.hasNext() ? it.next() : pai.createFolder(n);
}

function sanitizarNomeDrive(nome) {
  if (!nome || nome === '—') return '_Sem Fornecedor';
  return nome.replace(/[\/\\:*?"<>|]/g, '-').replace(/\s+/g, ' ').trim().substring(0, 80);
}

// ============================================================
// UTILITÁRIOS
// ============================================================
function getSheet(empresaKey, abaName) {
  const cfg = SHEETS[empresaKey];
  if (!cfg) throw new Error('Empresa não encontrada: ' + empresaKey);
  const ss    = SpreadsheetApp.openById(cfg.id);
  const sheet = ss.getSheetByName(abaName);
  if (!sheet) throw new Error('Aba "' + abaName + '" não encontrada em ' + cfg.name);
  return sheet;
}

function sheetToJson(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
