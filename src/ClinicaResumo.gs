/************************************************************
📊 DASHBOARD EPIDEMIOLÓGICO – Luky + GPT-5 (V12.2.3 – Hotfix estabilidade + métricas únicas)
• Funções em inglês nas fórmulas; separador de argumentos ";"
• Deduplicação por Prontuário (C) priorizando destino (Óbito > Residência/Outro Hospital > demais), permanência (R) e última Data Saída (Q)
• Abas requeridas:
  - 'Base Filtrada (Fórmula)' (A:Y)
  - 'LISTAS DE APOIO' (valores únicos por coluna com mesmo cabeçalho)
  - 'Municípios' (A:D municípios → Capital/RMF/Interior; G:K Procedência)
  - 'Cadastro CIDS' (B=Capítulo; G=Código CID10)
  - 'PERFIL EPIDEMIOLÓGICO' (I1 tipo período; J1 período; K1 ano)

⚠️ Regras:
  - Sem deduplicar: Clínica Entrada (U), Clínica Entrada (Setor N), Alta (N filtrando Destino O), Leito Equitópico (V)
  - Demais blocos contam via 'DadosÚnicos' (dedup)
*************************************************************/

/* ===== MENU RÁPIDO ===== */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Análises HUC')
    .addItem('Atualizar Dashboard', 'criarDashboardEpidemiologico')
    .addToUi();
}

/* ===== Helper: recriação segura de abas ===== */
function safeRecreateSheet_(ss, name, fallbackSheet) {
  let lastError = null;
  for (let attempt = 0; attempt < 3; attempt++) {
    const lock = LockService.getDocumentLock();
    if (!lock.tryLock(5000)) {
      Utilities.sleep(200 * (attempt + 1));
      continue;
    }
    try {
      let sheet = ss.getSheetByName(name);
      if (!sheet) {
        sheet = ss.insertSheet(name);
      } else {
        if (fallbackSheet && ss.getActiveSheet().getName() === name) {
          ss.setActiveSheet(fallbackSheet);
        }
        const fullRange = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
        fullRange.breakApart();
        fullRange.clear();
        const filter = sheet.getFilter();
        if (filter) filter.remove();
        sheet.setConditionalFormatRules([]);
        try {
          sheet.getCharts().forEach(chart => sheet.removeChart(chart));
        } catch (e) {}
      }
      Utilities.sleep(120);
      SpreadsheetApp.flush();
      return sheet;
    } catch (err) {
      lastError = err;
      Utilities.sleep(250 * (attempt + 1));
    } finally {
      lock.releaseLock();
    }
  }
  if (lastError) throw lastError;
  throw new Error('Falha ao recriar aba: ' + name);
}

function getColumnValues_(sheet, rangeA1) {
  return sheet
    .getRange(rangeA1)
    .getValues()
    .flat()
    .filter(value => value !== '' && value !== null);
}

function escapeFormulaString_(value) {
  return value.toString().replace(/"/g, '""');
}

function columnToLetter_(columnNumber) {
  let letter = '';
  let col = Math.max(1, columnNumber);
  while (col > 0) {
    const remainder = (col - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

const COL = {
  PRONTUARIO: 3,
  SEXO: 5,
  IDADE: 6,
  ESCOLARIDADE: 7,
  RACA_COR: 8,
  MUNICIPIO: 9,
  REGIAO_SAUDE: 10,
  ADS: 11,
  PROCEDENCIA: 12,
  CLINICA_ORIGEM: 13,
  SETOR: 14,
  DESTINO: 15,
  DATA_ENTRADA: 16,
  DATA_SAIDA: 17,
  PERMANENCIA: 18,
  CID: 19,
  ESPECIALIDADE: 21,
  LEITO: 22,
  OBITO_PRIORITARIO: 23,
  CLASSIFICACAO_OBITO: 24,
};

function normalizeText_(value) {
  if (value === null || value === undefined) return '';
  return value
    .toString()
    .normalize('NFD')
    .replace(/\p{Diacritic}/gu, '')
    .toUpperCase()
    .replace(/\s+/g, ' ')
    .trim();
}

function equalsText_(a, b) {
  return normalizeText_(a) === normalizeText_(b);
}

function toNumber_(value) {
  if (typeof value === 'number') return value;
  if (value instanceof Date && !isNaN(value)) return value.getTime();
  if (value === null || value === undefined) return NaN;
  const asString = value.toString().replace(',', '.').trim();
  if (asString === '') return NaN;
  const parsed = Number(asString);
  return isNaN(parsed) ? NaN : parsed;
}

function toDate_(value) {
  if (value instanceof Date && !isNaN(value)) return value;
  if (typeof value === 'number' && !isNaN(value)) {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    return new Date(excelEpoch.getTime() + value * 24 * 60 * 60 * 1000);
  }
  if (typeof value === 'string') {
    const parsed = new Date(value);
    if (!isNaN(parsed)) return parsed;
  }
  return null;
}

function toDateKey_(value) {
  const date = toDate_(value);
  if (!date) return null;
  return Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
}

function parseMonth_(raw) {
  if (!raw) return null;
  const value = raw.toString().trim();
  if (value === '') return null;
  const numeric = Number(value.replace(/[^0-9]/g, ''));
  if (!isNaN(numeric) && numeric >= 1 && numeric <= 12) return numeric;
  const abrevs = {
    JAN: 1,
    FEV: 2,
    MAR: 3,
    ABR: 4,
    MAI: 5,
    JUN: 6,
    JUL: 7,
    AGO: 8,
    SET: 9,
    OUT: 10,
    NOV: 11,
    DEZ: 12,
  };
  const key = normalizeText_(value).slice(0, 3);
  return abrevs[key] || null;
}

function median_(numbers) {
  const ordered = numbers.filter(n => !isNaN(n)).sort((a, b) => a - b);
  if (ordered.length === 0) return NaN;
  const mid = Math.floor(ordered.length / 2);
  if (ordered.length % 2 === 0) return (ordered[mid - 1] + ordered[mid]) / 2;
  return ordered[mid];
}

function durationToDays_(value) {
  if (typeof value === 'number') return value;
  if (value instanceof Date && !isNaN(value)) {
    const epoch = new Date(Date.UTC(1899, 11, 30));
    return (value.getTime() - epoch.getTime()) / (24 * 60 * 60 * 1000);
  }
  if (value === null || value === undefined) return NaN;
  const parsed = Number(value.toString().replace(',', '.'));
  return isNaN(parsed) ? NaN : parsed;
}

function destinoPriority_(value) {
  const normalized = normalizeText_(value);
  if (normalized.indexOf('OBITO') >= 0) return 0;
  if (normalized.indexOf('RESIDENCIA') >= 0 || normalized.indexOf('OUTRO HOSPITAL') >= 0) return 1;
  return 2;
}

function average_(numbers) {
  const valid = numbers.filter(n => !isNaN(n));
  if (valid.length === 0) return NaN;
  const sum = valid.reduce((acc, n) => acc + n, 0);
  return sum / valid.length;
}

/* ===== PRINCIPAL ===== */
function criarDashboardEpidemiologico() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shBase   = ss.getSheetByName('Base Filtrada (Fórmula)');
  const shApoio  = ss.getSheetByName('LISTAS DE APOIO');
  const shMuni   = ss.getSheetByName('Municípios');
  const shCIDS   = ss.getSheetByName('Cadastro CIDS');
  const shPerfil = ss.getSheetByName('PERFIL EPIDEMIOLÓGICO');

  if (!shBase || !shApoio || !shMuni || !shCIDS || !shPerfil) {
    SpreadsheetApp.getUi().alert('❌ Faltam abas obrigatórias: Base Filtrada (Fórmula), LISTAS DE APOIO, Municípios, Cadastro CIDS e PERFIL EPIDEMIOLÓGICO.');
    return;
  }

  const headerCols = shBase.getLastColumn();
  const headerRow = headerCols > 0 ? shBase.getRange(1, 1, 1, headerCols).getValues()[0] : [];
  const baseLastRow = shBase.getLastRow();
  const rawValues = headerCols > 0 && baseLastRow > 1
    ? shBase.getRange(2, 1, baseLastRow - 1, headerCols).getValues()
    : [];
  const rowsWithIndex = [];
  rawValues.forEach((row, idx) => {
    if (normalizeText_(row[COL.PRONTUARIO - 1]) !== '') {
      rowsWithIndex.push({ row, index: idx });
    }
  });

  const tipoRaw = shPerfil.getRange('I1').getDisplayValue();
  const mesRaw  = shPerfil.getRange('J1').getDisplayValue();
  const anoRaw  = shPerfil.getRange('K1').getDisplayValue();
  const setorA  = shPerfil.getRange('G1').getDisplayValue();
  const setorB  = shPerfil.getRange('H1').getDisplayValue();
  const obitoRaw = shPerfil.getRange('M1').getDisplayValue();

  const tipoKey  = normalizeText_(tipoRaw);
  const setorAKey = normalizeText_(setorA);
  const setorBKey = normalizeText_(setorB);
  const obitoKey  = normalizeText_(obitoRaw);

  const requireObito = /SIM$/.test(obitoKey);
  const useSaida = /SAIDA|ALTA/.test(tipoKey);
  const useAcum  = /ACUMUL/.test(tipoKey);
  const anoSel = (() => {
    if (!anoRaw) return null;
    const cleaned = Number(anoRaw.toString().replace(/[^0-9]/g, ''));
    return cleaned ? cleaned : null;
  })();
  const mesSel = parseMonth_(mesRaw);

  const setorIsGlobal = (!setorAKey && !setorBKey) || setorAKey === 'HUC (GERAL)' || setorBKey === 'HUC (GERAL)';
  function matchesSetor(value) {
    if (setorIsGlobal) return true;
    const target = normalizeText_(value);
    if (!setorAKey) return target === setorBKey;
    if (!setorBKey) return target === setorAKey;
    return target === setorAKey || target === setorBKey;
  }

  function matchesPeriodo(row) {
    if (useAcum) return true;
    const ref = useSaida ? row[COL.DATA_SAIDA - 1] : row[COL.DATA_ENTRADA - 1];
    const date = toDate_(ref);
    if (!date) return false;
    if (anoSel && date.getFullYear() !== anoSel) return false;
    if (mesSel && date.getMonth() + 1 !== mesSel) return false;
    return true;
  }

  function matchesProfile(row) {
    return matchesSetor(row[COL.SETOR - 1]) &&
      (!requireObito || destinoPriority_(row[COL.DESTINO - 1]) === 0) &&
      matchesPeriodo(row);
  }

  function deduplicateRows(rows) {
    const groups = new Map();
    rows.forEach(entry => {
      const key = normalizeText_(entry.row[COL.PRONTUARIO - 1]);
      if (!key) return;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(entry);
    });
    const selected = [];
    groups.forEach(list => {
      list.sort((a, b) => {
        const priA = destinoPriority_(a.row[COL.DESTINO - 1]);
        const priB = destinoPriority_(b.row[COL.DESTINO - 1]);
        if (priA !== priB) return priA - priB;
        const permA = durationToDays_(a.row[COL.PERMANENCIA - 1]);
        const permB = durationToDays_(b.row[COL.PERMANENCIA - 1]);
        if (!isNaN(permA) || !isNaN(permB)) {
          if (isNaN(permA)) return 1;
          if (isNaN(permB)) return -1;
          if (permA !== permB) return permB - permA;
        }
        const saidaA = toDate_(a.row[COL.DATA_SAIDA - 1]);
        const saidaB = toDate_(b.row[COL.DATA_SAIDA - 1]);
        const timeA = saidaA ? saidaA.getTime() : -Infinity;
        const timeB = saidaB ? saidaB.getTime() : -Infinity;
        if (timeA !== timeB) return timeB - timeA;
        return a.index - b.index;
      });
      selected.push(list[0]);
    });
    selected.sort((a, b) => {
      const prontA = normalizeText_(a.row[COL.PRONTUARIO - 1]);
      const prontB = normalizeText_(b.row[COL.PRONTUARIO - 1]);
      if (prontA < prontB) return -1;
      if (prontA > prontB) return 1;
      return a.index - b.index;
    });
    return selected.map(item => item.row);
  }

  const filteredRows = rowsWithIndex.filter(entry => matchesProfile(entry.row)).map(entry => entry.row);
  const dedupRows = deduplicateRows(rowsWithIndex);
  const dedupFilteredRows = dedupRows.filter(row => matchesProfile(row));

  const periodoTexto = `${tipoRaw || ''} – ${mesRaw || ''} / ${anoRaw || ''}`;

  function dateMin(rows, colIndex) {
    let min = null;
    rows.forEach(row => {
      const date = toDate_(row[colIndex - 1]);
      if (!date) return;
      if (!min || date < min) min = date;
    });
    return min;
  }

  function dateMax(rows, colIndex) {
    let max = null;
    rows.forEach(row => {
      const date = toDate_(row[colIndex - 1]);
      if (!date) return;
      if (!max || date > max) max = date;
    });
    return max;
  }

  const filteredSheetName = 'Base Filtrada (Filtro)';
  const shBaseFiltro = safeRecreateSheet_(ss, filteredSheetName, shBase);
  if (headerRow.length) {
    shBaseFiltro.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
  }
  if (filteredRows.length) {
    shBaseFiltro.getRange(2, 1, filteredRows.length, headerCols).setValues(filteredRows);
  }
  SpreadsheetApp.flush();
  Utilities.sleep(120);
  shBaseFiltro.hideSheet();

  /* ===== PALETA / UI ===== */
  const COLOR = {
    primary:      '#159382',
    primaryDark:  '#0E6D62',
    header:       '#E9F6F4',
    bandA:        '#FFFFFF',
    bandB:        '#F7F9F9',
    textMuted:    '#5F6B6B',
    border:       '#E3E7E7',
  };

  /* ===== 1) DEDUP "DadosÚnicos" ===== */
  const shUni = safeRecreateSheet_(ss, 'DadosÚnicos', shBase);
  shUni.getRange('A1')
    .setValue('⚙️ Base deduplicada por prontuário (prioridade destino > permanência > última saída)')
    .setFontWeight('bold')
    .setFontColor(COLOR.textMuted);
  if (dedupRows.length) {
    shUni.getRange(2, 1, dedupRows.length, headerCols).setValues(dedupRows);
  }
  SpreadsheetApp.flush();
  Utilities.sleep(120);
  shUni.hideSheet();

  const shUniFiltro = safeRecreateSheet_(ss, 'DadosÚnicos Filtrados', shUni);
  shUniFiltro.getRange('A1')
    .setValue('⚙️ Base deduplicada filtrada pelo perfil selecionado')
    .setFontWeight('bold')
    .setFontColor(COLOR.textMuted);
  if (dedupFilteredRows.length) {
    shUniFiltro.getRange(2, 1, dedupFilteredRows.length, headerCols).setValues(dedupFilteredRows);
  }
  SpreadsheetApp.flush();
  Utilities.sleep(120);
  shUniFiltro.hideSheet();

  /* ===== 2) ⚙️DATA (séries p/ gráficos e auxiliares) ===== */
  const shData = safeRecreateSheet_(ss, '⚙️DATA', shBase);

  shData.getRange('A1:C1').setValues([['Datas (período)','Entradas (dia)','Altas (dia)']]).setFontWeight('bold');
  const entradasPorDia = new Map();
  const altasPorDia = new Map();
  filteredRows.forEach(row => {
    const entradaKey = toDateKey_(row[COL.DATA_ENTRADA - 1]);
    if (entradaKey !== null) {
      entradasPorDia.set(entradaKey, (entradasPorDia.get(entradaKey) || 0) + 1);
    }
    const saidaKey = toDateKey_(row[COL.DATA_SAIDA - 1]);
    if (saidaKey !== null) {
      altasPorDia.set(saidaKey, (altasPorDia.get(saidaKey) || 0) + 1);
    }
  });
  const datasFluxo = Array.from(new Set([...entradasPorDia.keys(), ...altasPorDia.keys()])).sort((a, b) => a - b);
  if (datasFluxo.length) {
    const fluxoValores = datasFluxo.map(key => [new Date(key), entradasPorDia.get(key) || 0, altasPorDia.get(key) || 0]);
    shData.getRange(2, 1, fluxoValores.length, 3).setValues(fluxoValores);
  }

  shData.getRange('E1:F1').setValues([['Especialidade','Qtd (dedup)']]).setFontWeight('bold');
  const listaEspecialidades = getColumnValues_(shApoio, 'U2:U');
  if (listaEspecialidades.length) {
    const espTotais = listaEspecialidades.map(label => {
      const count = dedupRows.reduce((acc, row) => acc + (equalsText_(row[COL.ESPECIALIDADE - 1], label) ? 1 : 0), 0);
      return [label, count];
    });
    shData.getRange(2, 5, espTotais.length, 2).setValues(espTotais);
  }

  shData.getRange('L1:N1').setValues([['Capítulo CID10 (catálogo)','Código (catálogo)','—']]).setFontWeight('bold');
  const cidCatalog = [];
  const cidMap = new Map();
  const cidLast = shCIDS.getLastRow();
  if (cidLast > 1) {
    const capVals = shCIDS.getRange(2, 2, cidLast - 1, 1).getValues().flat();
    const codVals = shCIDS.getRange(2, 7, cidLast - 1, 1).getValues().flat();
    for (let i = 0; i < codVals.length; i++) {
      const codigo = normalizeText_(codVals[i]);
      const capitulo = (capVals[i] || '').toString().trim();
      if (!codigo || !capitulo) continue;
      cidCatalog.push([capitulo, codVals[i], '']);
      cidMap.set(codigo, capitulo);
      cidMap.set(codigo.replace('.', ''), capitulo);
    }
  }
  if (cidCatalog.length) {
    shData.getRange(2, 12, cidCatalog.length, 3).setValues(cidCatalog);
  }

  shData.getRange('O1:P1').setValues([['Capítulo (uso)','Qtd (uso)']]).setFontWeight('bold');
  const capCounts = new Map();
  dedupRows.forEach(row => {
    const raw = row[COL.CID - 1];
    if (!raw) return;
    const codigo = normalizeText_(raw).replace('.', '');
    if (!codigo) return;
    const capitulo = cidMap.get(codigo) || cidMap.get(codigo.slice(0, 3));
    if (!capitulo) return;
    capCounts.set(capitulo, (capCounts.get(capitulo) || 0) + 1);
  });
  const capEntries = Array.from(capCounts.entries()).sort((a, b) => b[1] - a[1]);
  if (capEntries.length) {
    shData.getRange(2, 15, capEntries.length, 2).setValues(capEntries);
  }

  SpreadsheetApp.flush();
  Utilities.sleep(150);
  shData.hideSheet();

  const totalInternacoes = filteredRows.length;
  const totalPacientes = dedupRows.length;
  const totalObitos = dedupRows.reduce(
    (acc, row) => acc + (destinoPriority_(row[COL.DESTINO - 1]) === 0 ? 1 : 0),
    0
  );
  const permanencias = dedupRows
    .map(row => durationToDays_(row[COL.PERMANENCIA - 1]))
    .filter(value => !isNaN(value));
  const mediaPermanencia = average_(permanencias);
  const somaPermanencia = permanencias.reduce((acc, value) => acc + value, 0);
  const primeiraInternacao = dateMin(dedupRows, COL.DATA_ENTRADA);
  const ultimaSaida = dateMax(dedupRows, COL.DATA_SAIDA);
  const idadeValores = dedupRows
    .map(row => toNumber_(row[COL.IDADE - 1]))
    .filter(value => !isNaN(value));
  const idadeMedia = average_(idadeValores);
  const idadeMediana = median_(idadeValores);
  const mediaPermanenciaValor = isNaN(mediaPermanencia) ? null : mediaPermanencia;
  const somaPermanenciaValor = isNaN(somaPermanencia) ? 0 : somaPermanencia;
  const idadeMediaValor = isNaN(idadeMedia) ? null : idadeMedia;
  const idadeMedianaValor = isNaN(idadeMediana) ? null : idadeMediana;
  const idadeAte19 = idadeValores.filter(v => v <= 19).length;
  const idade20a59 = idadeValores.filter(v => v >= 20 && v <= 59).length;
  const idade60Mais = idadeValores.filter(v => v >= 60).length;

  const taxaObito = totalPacientes ? totalObitos / totalPacientes : 0;

  /* ===== 3) DASHBOARD ===== */
  const sh = safeRecreateSheet_(ss, 'Dashboard', shBase);

  // Grid base e tipografia
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, 12, 120);
  sh.setColumnWidth(1, 220);
  sh.setRowHeights(1, 60, 28);
  sh.getRange('A1:L2000').setFontFamily('Roboto');

  // Helpers visuais
  function headerBlock(r, c, text, span=12) {
    sh.getRange(r, c, 1, span).merge()
      .setValue(text).setFontSize(14).setFontWeight('bold')
      .setBackground(COLOR.primary).setFontColor('#FFFFFF')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  }
  function subHeader(r, c, text, span=12) {
    sh.getRange(r, c, 1, span).merge()
      .setValue(text).setFontWeight('bold').setBackground(COLOR.header)
      .setFontColor('#0C3D3A').setHorizontalAlignment('left').setVerticalAlignment('middle');
  }
  function bandTable(rangeA1, pctCols=[]) {
    const rg = sh.getRange(rangeA1);
    const rows = rg.getNumRows(), cols = rg.getNumColumns();
    const backgrounds = Array.from({ length: rows }, (_, idx) =>
      Array(cols).fill(idx % 2 === 0 ? COLOR.bandA : COLOR.bandB)
    );
    rg.setBackgrounds(backgrounds)
      .setBorder(true,true,true,true,true,true)
      .setVerticalAlignment('middle');
    pctCols.forEach(idx => {
      if (rows > 1) sh.getRange(rg.getRow()+1, rg.getColumn()+idx-1, rows-1, 1).setNumberFormat('0.0%');
    });
  }
  function miniMuted(r, c, label, value, fmt) {
    sh.getRange(r, c).setValue(label).setFontColor(COLOR.textMuted);
    const cell = sh.getRange(r, c + 1);
    if (value === null || value === undefined) cell.setValue('');
    else cell.setValue(value);
    cell.setFontWeight('bold');
    if (fmt) cell.setNumberFormat(fmt);
  }
  function kpiCard(r, c, title, value, fmt, icon='') {
    const titleR = sh.getRange(r, c, 1, 3).merge();
    const valueR = sh.getRange(r + 1, c, 1, 3).merge();
    titleR.setValue(`${icon} ${title}`).setFontWeight('bold')
      .setBackground(COLOR.header).setFontColor('#0C3D3A')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    if (value === null || value === undefined) valueR.setValue('');
    else valueR.setValue(value);
    valueR.setFontSize(18).setFontWeight('bold')
      .setBackground('#FFFFFF').setHorizontalAlignment('left').setVerticalAlignment('middle');
    if (fmt) valueR.setNumberFormat(fmt);
    sh.getRange(r, c, 2, 3).setBorder(true, true, true, true, true, true);
  }
  function linkTo(r, c, text, anchorCellA1) {
    const gid = sh.getSheetId();
    const rich = SpreadsheetApp.newRichTextValue()
      .setText(text)
      .setLinkUrl(`#gid=${gid}&range=${anchorCellA1}`)
      .build();
    sh.getRange(r, c).setRichTextValue(rich)
      .setFontColor(COLOR.primaryDark)
      .setFontWeight('bold');
  }

  let row = 1;

  /* Header */
  headerBlock(row, 1, '📊 Dashboard Epidemiológico – HUC'); row += 2;

  // Período + metadata
  miniMuted(row, 1, 'Período selecionado', periodoTexto);
  row++;

  /* KPI cards – linha 1 */
  const kpiRow1 = row;
  kpiCard(kpiRow1, 1,  'Pacientes Únicos',            totalPacientes, '#,##0', '👤');
  kpiCard(kpiRow1, 4,  'Total de Internações',        totalInternacoes, '#,##0', '🏥');
  kpiCard(kpiRow1, 7,  'Taxa de Óbito',               taxaObito, '0.0%', '☠️');
  kpiCard(kpiRow1, 10, 'Média de Permanência (dias)', mediaPermanenciaValor, '0.00', '⏱️');
  row = kpiRow1 + 3;

  /* KPI cards – linha 2 */
  const kpiRow2 = row;
  kpiCard(kpiRow2, 1,  'Primeira Internação', primeiraInternacao, 'dd/mm/yyyy', '📅');
  kpiCard(kpiRow2, 4,  'Última Alta/Saída',   ultimaSaida, 'dd/mm/yyyy', '📅');
  kpiCard(kpiRow2, 7,  'Idade Média',         idadeMediaValor, '0.0', '👶');
  kpiCard(kpiRow2, 10, 'Dias-Paciente (soma)', somaPermanenciaValor, '#,##0', '📈');
  row = kpiRow2 + 4;

  /* Navegação rápida */
  subHeader(row, 1, '📌 Navegação'); row++;
  linkTo(row, 1, '1) Fluxo e Procedência',           'A12');
  linkTo(row, 3, '2) Perfil Sociodemográfico',       'A40');
  linkTo(row, 6, '3) Clínicas (Origem/Entradas/Alta)', 'A80');
  linkTo(row, 9, '4) CID-10 e Especialidades',       'A120');
  row += 2;

  /* 1) Fluxo e Procedência */
  subHeader(row, 1, '1) Fluxo do Paciente e Procedência (agrupada)'); row++;

  // Procedência agrupada (conta em DadosÚnicos!L:L)
  sh.getRange(row,1,1,3).setValues([['Categoria','Qtd','%']]).setBackground(COLOR.header).setFontWeight('bold');
  const gruposProc = [
    ['Hospital', getColumnValues_(shMuni, 'G2:G')],
    ['UPA', getColumnValues_(shMuni, 'H2:H')],
    ['Ambulatório', getColumnValues_(shMuni, 'I2:I')],
    ['Residência', getColumnValues_(shMuni, 'J2:J')],
    ['CRESUS', getColumnValues_(shMuni, 'K2:K')],
  ];
  const startProc = row + 1;
  if (gruposProc.length) {
    sh.getRange(startProc, 1, gruposProc.length, 1)
      .setValues(gruposProc.map(([nome]) => [nome]));
    const procFormulas = gruposProc.map(([, lista]) => {
      if (!lista.length) return ['=0'];
      const cond = lista
        .map(v => `(DadosÚnicos!L:L="${escapeFormulaString_(v)}")`)
        .join('+');
      return [`=ARRAYFORMULA(SUM(--(${cond})))`];
    });
    sh.getRange(startProc, 2, gruposProc.length, 1).setFormulas(procFormulas);
  }
  const totalProc = startProc + gruposProc.length;
  sh.getRange(totalProc, 1, 1, 3).setValues([['TOTAL','','']]).setBackground(COLOR.header).setFontWeight('bold');
  sh.getRange(totalProc, 2).setFormula(gruposProc.length ? `=SUM(B${startProc}:B${totalProc - 1})` : '=0');
  if (gruposProc.length) {
    const pctFormulas = Array.from({ length: gruposProc.length }, () => [`=IFERROR(RC[-1]/R${totalProc}C2;0)`]);
    sh.getRange(startProc, 3, gruposProc.length, 1).setFormulasR1C1(pctFormulas);
  }
  bandTable(`A${row}:C${totalProc}`, [3]);

  // Gráfico de rosquinha – Procedência
  try { sh.getCharts().forEach(c => sh.removeChart(c)); } catch(e){}
  const donut = sh.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sh.getRange(`A${startProc}:B${totalProc-1}`))
    .setPosition(row, 8, 0, 0)
    .setOption('pieHole', 0.5)
    .setOption('legend', {position: 'right'})
    .setOption('title', 'Procedência (agrupada)')
    .build();
  sh.insertChart(donut);

  // Gráfico Linha: Entradas × Altas por dia (⚙️DATA!A:C)
  const line = sh.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(shData.getRange('A1:C'))
    .setOption('title', 'Entradas × Altas por dia')
    .setOption('legend', {position: 'bottom'})
    .setPosition(row+16, 8, 0, 0)
    .build();
  sh.insertChart(line);

  row = totalProc + 3;

  /* 2) Perfil Sociodemográfico */
  subHeader(row, 1, '2) Perfil Sociodemográfico'); row++;

  // Município agrupado (Fortaleza / RMF / Interior)
  sh.getRange(row,1,1,3).setValues([['Município (agrupado)','Qtd','%']]).setBackground(COLOR.header).setFontWeight('bold');
  const muniData = shMuni.getRange(2, 1, Math.max(shMuni.getLastRow() - 1, 0), 3).getValues();
  const fortaleza = [];
  const metro = [];
  const interior = [];
  muniData.forEach(([nome, isCapital, isRmf]) => {
    if (!nome) return;
    if (isCapital === 'Sim') fortaleza.push(nome);
    else if (isRmf === 'Sim') metro.push(nome);
    else interior.push(nome);
  });
  const mStart = row+1;
  const gruposM = [['Fortaleza', fortaleza], ['RMF', metro], ['Interior', interior]];
  sh.getRange(mStart, 1, gruposM.length, 1).setValues(gruposM.map(([label]) => [label]));
  const gruposFormulas = gruposM.map(([, lista]) => {
    if (!lista.length) return ['=0'];
    const cond = lista
      .map(v => `(DadosÚnicos!I:I="${escapeFormulaString_(v)}")`)
      .join('+');
    return [`=ARRAYFORMULA(SUM(--(${cond})))`];
  });
  sh.getRange(mStart, 2, gruposFormulas.length, 1).setFormulas(gruposFormulas);
  const rOutros = mStart + gruposM.length;
  sh.getRange(rOutros, 1).setValue('Outros');
  sh.getRange(rOutros, 2).setFormula(`=MAX(0;COUNTA(DadosÚnicos!I2:I)-SUM(B${mStart}:B${rOutros - 1}))`);
  const mTot = rOutros + 1;
  sh.getRange(mTot, 1, 1, 3).setValues([['TOTAL','','']]).setBackground(COLOR.header).setFontWeight('bold');
  sh.getRange(mTot, 2).setFormula(`=SUM(B${mStart}:B${rOutros})`);
  sh.getRange(mStart, 3, rOutros - mStart + 1, 1)
    .setFormulasR1C1(Array.from({ length: rOutros - mStart + 1 }, () => [`=IFERROR(RC[-1]/R${mTot}C2;0)`]));
  bandTable(`A${row}:C${mTot}`, [3]);
  row = mTot + 2;

    // Blocos simples (dedup)
  function blocoSimples(tituloBloco, colBase, apoioCol, startRow) {
    sh.getRange(startRow, 1, 1, 3)
      .setValues([[tituloBloco, 'Qtd', '%']])
      .setBackground(COLOR.header)
      .setFontWeight('bold');
    const labels = getColumnValues_(shApoio, `${apoioCol}2:${apoioCol}`);
    if (labels.length === 0) return startRow + 2;
    sh.getRange(startRow + 1, 1, labels.length, 1).setValues(labels.map(v => [v]));
    const end = startRow + labels.length;
    const total = end + 1;
    const countFormulas = Array.from({ length: labels.length }, () => [`=COUNTIFS(DadosÚnicos!${colBase}:${colBase};RC1)`]);
    sh.getRange(startRow + 1, 2, labels.length, 1).setFormulasR1C1(countFormulas);
    sh.getRange(total, 1, 1, 3)
      .setValues([['TOTAL', '', '']])
      .setBackground(COLOR.header)
      .setFontWeight('bold');
    sh.getRange(total, 2).setFormula(`=SUM(B${startRow + 1}:B${end})`);
    sh.getRange(startRow + 1, 3, labels.length, 1)
      .setFormulasR1C1(Array.from({ length: labels.length }, () => [`=IFERROR(RC[-1]/R${total}C2;0)`]));
    bandTable(`A${startRow}:C${total}`, [3]);
    return total + 2;
  }
  let rDemog = row;
  rDemog = blocoSimples('Sexo',                         'E', 'E', rDemog);
  rDemog = blocoSimples('Raça/Cor',                     'H', 'H', rDemog);
  rDemog = blocoSimples('Escolaridade',                 'G', 'G', rDemog);
  rDemog = blocoSimples('Região de Saúde',              'J', 'J', rDemog);
  rDemog = blocoSimples('Área Descentralizada de Saúde (ADS)', 'K', 'K', rDemog);

  // Idade (faixas)
  subHeader(rDemog, 1, 'Idade (faixas etárias)'); rDemog++;
  sh.getRange(rDemog,1,1,3).setValues([['Faixa','Qtd','%']]).setBackground(COLOR.header).setFontWeight('bold');
  const fStart = rDemog+1;
  const idadeFaixas = [
    ['≤ 19 anos', idadeAte19],
    ['20 a 59 anos', idade20a59],
    ['≥ 60 anos', idade60Mais],
  ];
  const idadeTotal = idadeFaixas.reduce((acc, [, count]) => acc + count, 0);
  const idadePerc = idadeFaixas.map(([, count]) => (idadeTotal ? count / idadeTotal : 0));
  sh.getRange(fStart,1,3,1).setValues(idadeFaixas.map(row => [row[0]]));
  sh.getRange(fStart,2,3,1).setValues(idadeFaixas.map(row => [row[1]]));
  sh.getRange(fStart,3,3,1).setValues(idadePerc.map(value => [value]));
  const fTot = fStart+3;
  sh.getRange(fTot,1,1,3).setValues([['TOTAL', idadeTotal, '']]).setBackground(COLOR.header).setFontWeight('bold');
  bandTable(`A${rDemog}:C${fTot}`, [3]);

  rDemog = fTot + 2;
  miniMuted(rDemog,1,'Idade média',   idadeMediaValor,"0.0"); rDemog++;
  miniMuted(rDemog,1,'Idade mediana', idadeMedianaValor,"0.0"); rDemog += 2;

  row = rDemog;

  /* 3) Clínicas – Origem (dedup) | Entradas (U e Setor N) | Alta (Destino O) | Leito (V) */
  subHeader(row,1,'3) Clínicas – Origem (dedup) | Entradas | Alta (destino) | Leito'); row++;

  function blocoBaseCompletaSimples(tituloBloco, colBase, apoioCol, startRow) {
    sh.getRange(startRow, 1, 1, 3)
      .setValues([[tituloBloco, 'Qtd', '%']])
      .setBackground(COLOR.header)
      .setFontWeight('bold');
    const labels = getColumnValues_(shApoio, `${apoioCol}2:${apoioCol}`);
    if (labels.length === 0) return startRow + 2;
    sh.getRange(startRow + 1, 1, labels.length, 1).setValues(labels.map(v => [v]));
    const end = startRow + labels.length;
    const total = end + 1;
    const countFormulas = Array.from({ length: labels.length }, () => [`=COUNTIFS('${filteredSheetName}'!${colBase}:${colBase};RC1)`]);
    sh.getRange(startRow + 1, 2, labels.length, 1).setFormulasR1C1(countFormulas);
    sh.getRange(total, 1, 1, 3)
      .setValues([['TOTAL', '', '']])
      .setBackground(COLOR.header)
      .setFontWeight('bold');
    sh.getRange(total, 2).setFormula(`=SUM(B${startRow + 1}:B${end})`);
    sh.getRange(startRow + 1, 3, labels.length, 1)
      .setFormulasR1C1(Array.from({ length: labels.length }, () => [`=IFERROR(RC[-1]/R${total}C2;0)`]));
    bandTable(`A${startRow}:C${total}`, [3]);
    return total + 2;
  }
  function blocoEntradaSetorCompleta(startRow) {
    const tituloBloco = 'Clínica Entrada (Setor) – base completa';
    sh.getRange(startRow, 1, 1, 5)
      .setValues([[tituloBloco, 'Qtd total', '%', 'Qtd únicos', '% únicos']])
      .setBackground(COLOR.header).setFontWeight('bold');
    const labels = getColumnValues_(shApoio, 'N2:N');
    if (labels.length === 0) return startRow + 2;
    sh.getRange(startRow + 1, 1, labels.length, 1).setValues(labels.map(v => [v]));
    const end = startRow + labels.length;
    const total = end + 1;
    const countFormulas = Array.from({ length: labels.length }, () => [`=COUNTIFS('${filteredSheetName}'!N:N;RC1)`]);
    const uniqueFormulas = Array.from(
      { length: labels.length },
      () => [`=COUNTIFS(DadosÚnicos!N:N;RC1)`]
    );
    sh.getRange(startRow + 1, 2, labels.length, 1).setFormulasR1C1(countFormulas);
    sh.getRange(startRow + 1, 4, labels.length, 1).setFormulasR1C1(uniqueFormulas);
    sh.getRange(total, 1, 1, 5)
      .setValues([['TOTAL', '', '', '', '']])
      .setBackground(COLOR.header).setFontWeight('bold');
    sh.getRange(total, 2).setFormula(`=SUM(B${startRow + 1}:B${end})`);
    sh.getRange(total, 4).setFormula(`=SUM(D${startRow + 1}:D${end})`);
    const pctTotal = Array.from({ length: labels.length }, () => [`=IFERROR(RC[-1]/R${total}C2;0)`]);
    const pctUnique = Array.from({ length: labels.length }, () => [`=IFERROR(RC[-1]/R${total}C4;0)`]);
    sh.getRange(startRow + 1, 3, labels.length, 1).setFormulasR1C1(pctTotal);
    sh.getRange(startRow + 1, 5, labels.length, 1).setFormulasR1C1(pctUnique);
    sh.getRange(total, 3, 1, 1).setValue('');
    sh.getRange(total, 5, 1, 1).setValue('');
    bandTable(`A${startRow}:E${total}`, [3,5]);
    return total + 2;
  }
  function blocoSaidaSetorCompleta(startRow) {
    const tituloBloco = 'Clínica Saídas (Setor) – base completa + únicos';
    sh.getRange(startRow, 1, 1, 5)
      .setValues([[tituloBloco, 'Qtd total', '%', 'Qtd únicos', '% únicos']])
      .setBackground(COLOR.header)
      .setFontWeight('bold');
    const labels = getColumnValues_(shApoio, 'N2:N');
    if (labels.length === 0) return startRow + 2;
    sh.getRange(startRow + 1, 1, labels.length, 1).setValues(labels.map(v => [v]));
    const end = startRow + labels.length;
    const total = end + 1;
    const countFormulas = Array.from(
      { length: labels.length },
      () => [
        `=COUNTIFS('${filteredSheetName}'!N:N;RC1;'${filteredSheetName}'!O:O;"Óbito")` +
        `+COUNTIFS('${filteredSheetName}'!N:N;RC1;'${filteredSheetName}'!O:O;"Residência")` +
        `+COUNTIFS('${filteredSheetName}'!N:N;RC1;'${filteredSheetName}'!O:O;"Outro hospital")`
      ]
    );
    const uniqueFormulas = Array.from(
      { length: labels.length },
      () => [
        "=COUNTIFS(DadosÚnicos!N:N;RC1;DadosÚnicos!O:O;\"Óbito\")" +
        "+COUNTIFS(DadosÚnicos!N:N;RC1;DadosÚnicos!O:O;\"Residência\")" +
        "+COUNTIFS(DadosÚnicos!N:N;RC1;DadosÚnicos!O:O;\"Outro hospital\")"
      ]
    );
    sh.getRange(startRow + 1, 2, labels.length, 1).setFormulasR1C1(countFormulas);
    sh.getRange(startRow + 1, 4, labels.length, 1).setFormulasR1C1(uniqueFormulas);
    sh.getRange(total, 1, 1, 5)
      .setValues([['TOTAL', '', '', '', '']])
      .setBackground(COLOR.header)
      .setFontWeight('bold');
    sh.getRange(total, 2).setFormula(`=SUM(B${startRow + 1}:B${end})`);
    sh.getRange(total, 4).setFormula(`=SUM(D${startRow + 1}:D${end})`);
    const pctTotal = Array.from({ length: labels.length }, () => [`=IFERROR(RC[-1]/R${total}C2;0)`]);
    const pctUnique = Array.from({ length: labels.length }, () => [`=IFERROR(RC[-1]/R${total}C4;0)`]);
    sh.getRange(startRow + 1, 3, labels.length, 1).setFormulasR1C1(pctTotal);
    sh.getRange(startRow + 1, 5, labels.length, 1).setFormulasR1C1(pctUnique);
    sh.getRange(total, 3, 1, 1).setValue('');
    sh.getRange(total, 5, 1, 1).setValue('');
    bandTable(`A${startRow}:E${total}`, [3,5]);
    return total + 2;
  }
  function blocoDedupSimples(tituloBloco, colBase, apoioCol, startRow) {
    sh.getRange(startRow, 1, 1, 3)
      .setValues([[tituloBloco, 'Qtd', '%']])
      .setBackground(COLOR.header)
      .setFontWeight('bold');
    const labels = getColumnValues_(shApoio, `${apoioCol}2:${apoioCol}`);
    if (labels.length === 0) return startRow + 2;
    sh.getRange(startRow + 1, 1, labels.length, 1).setValues(labels.map(v => [v]));
    const end = startRow + labels.length;
    const total = end + 1;
    const countFormulas = Array.from({ length: labels.length }, () => [`=COUNTIFS(DadosÚnicos!${colBase}:${colBase};RC1)`]);
    sh.getRange(startRow + 1, 2, labels.length, 1).setFormulasR1C1(countFormulas);
    sh.getRange(total, 1, 1, 3)
      .setValues([['TOTAL', '', '']])
      .setBackground(COLOR.header)
      .setFontWeight('bold');
    sh.getRange(total, 2).setFormula(`=SUM(B${startRow + 1}:B${end})`);
    sh.getRange(startRow + 1, 3, labels.length, 1)
      .setFormulasR1C1(Array.from({ length: labels.length }, () => [`=IFERROR(RC[-1]/R${total}C2;0)`]));
    bandTable(`A${startRow}:C${total}`, [3]);
    return total + 2;
  }
  function blocoAltaPorDestino(tituloBloco, startRow) {
    const destinos = ['Óbito', 'Residência', 'Outro hospital'];
    sh.getRange(startRow, 1, 1, 5)
      .setValues([[tituloBloco, 'Qtd total', '%', 'Qtd únicos', '% únicos']])
      .setBackground(COLOR.header).setFontWeight('bold');
    sh.getRange(startRow + 1, 1, destinos.length, 1).setValues(destinos.map(v => [v]));
    const end = startRow + destinos.length;
    const total = end + 1;
    const countFormulas = destinos.map(() => [`=COUNTIFS('${filteredSheetName}'!O:O;RC1)`]);
    const uniqueFormulas = destinos.map(() => [`=COUNTIFS(DadosÚnicos!O:O;RC1)`]);
    sh.getRange(startRow + 1, 2, destinos.length, 1).setFormulasR1C1(countFormulas);
    sh.getRange(startRow + 1, 4, destinos.length, 1).setFormulasR1C1(uniqueFormulas);
    sh.getRange(total, 1, 1, 5)
      .setValues([['TOTAL', '', '', '', '']])
      .setBackground(COLOR.header).setFontWeight('bold');
    sh.getRange(total, 2).setFormula(`=SUM(B${startRow + 1}:B${end})`);
    sh.getRange(total, 4).setFormula(`=SUM(D${startRow + 1}:D${end})`);
    const pctTotal = Array.from({ length: destinos.length }, () => [`=IFERROR(RC[-1]/R${total}C2;0)`]);
    const pctUnique = Array.from({ length: destinos.length }, () => [`=IFERROR(RC[-1]/R${total}C4;0)`]);
    sh.getRange(startRow + 1, 3, destinos.length, 1).setFormulasR1C1(pctTotal);
    sh.getRange(startRow + 1, 5, destinos.length, 1).setFormulasR1C1(pctUnique);
    sh.getRange(total, 3, 1, 1).setValue('');
    sh.getRange(total, 5, 1, 1).setValue('');
    bandTable(`A${startRow}:E${total}`, [3,5]);
    return total + 2;
  }

  let rClin = row;
  // Origem (Emergência) – deduplicada
  rClin = blocoDedupSimples('Clínica Origem (Emergência) – dedup', 'M', 'M', rClin);

  // Entradas – Especialidade (U) (base completa, já existia)
  rClin = blocoBaseCompletaSimples('Clínica Entrada (Especialidade) – base completa', 'U', 'U', rClin);

  // NOVO: Entradas – Setor (N) com totais e únicos
  rClin = blocoEntradaSetorCompleta(rClin);

  // NOVO: Saídas – Setor (N) filtrando destinos de saída prioritários
  rClin = blocoSaidaSetorCompleta(rClin);

  // Alta (Saída) – Destino (O) com totais e únicos
  rClin = blocoAltaPorDestino('Clínica Alta (Saída) – Destino (Óbito, Residência, Outro hospital)', rClin);

  // Leito Equitópico – base completa
  rClin = blocoBaseCompletaSimples('Leito Equitópico – base completa', 'V', 'V', rClin);

  // Outros blocos dedup
  rClin = blocoDedupSimples('Destino do Paciente', 'O', 'O', rClin);
  rClin = blocoDedupSimples('Óbito Prioritário',  'W', 'W', rClin);
  rClin = blocoDedupSimples('Classificação do Óbito', 'X', 'X', rClin);

  row = rClin + 1;

  /* 4) Capítulos do CID-10 e Especialidades */
  subHeader(row,1,'4) Capítulos do CID-10 e Especialidades'); row++;

  // Capítulos do CID-10 (⚙️DATA!O:P; com fallback)
  sh.getRange(row,1,1,3).setValues([['Capítulo CID10','Qtd','%']]).setBackground(COLOR.header).setFontWeight('bold');
  SpreadsheetApp.flush();

  const capStart = row+1;
  let capVals = shData.getRange('O2:P').getValues().filter(r => r[0]); // [[Cap, Qtd],...]

  if (capVals.length === 0) {
    const cLast = shCIDS.getLastRow();
    const cCap = shCIDS.getRange('B2:B' + cLast).getValues().flat();
    const cCod = shCIDS.getRange('G2:G' + cLast).getValues().flat();
    const mapCidToCap = {};
    for (let i = 0; i < cCod.length; i++) {
      const code = (cCod[i] || '').toString().trim().toUpperCase();
      const cap  = (cCap[i] || '').toString().trim();
      if (code && cap) mapCidToCap[code] = cap;
    }
    const uLast = shUni.getLastRow();
    const duCids = shUni.getRange('S2:S' + uLast).getValues().flat();
    const capCounts = {};
    duCids.forEach(raw => {
      let k = (raw || '').toString().trim().toUpperCase();
      if (!k) return;
      const kNoDot = k.replace('.', '');
      const cap = mapCidToCap[k] || mapCidToCap[kNoDot] || mapCidToCap[kNoDot.slice(0,3)];
      if (!cap) return;
      capCounts[cap] = (capCounts[cap] || 0) + 1;
    });
    capVals = Object.entries(capCounts).sort((a,b)=>b[1]-a[1]);
  }

  if (capVals.length > 0) {
    sh.getRange(capStart,1,capVals.length,1).setValues(capVals.map(r=>[r[0]]));
    sh.getRange(capStart,2,capVals.length,1).setValues(capVals.map(r=>[r[1]]));
    const capTot = capStart + capVals.length;
    sh.getRange(capTot,1,1,3).setValues([['TOTAL','','']]).setBackground(COLOR.header).setFontWeight('bold');
    sh.getRange(capTot,2).setFormula(`=SUM(B${capStart}:B${capTot-1})`);
    for (let r=capStart; r<=capTot-1; r++) sh.getRange(r,3).setFormula(`=IFERROR(B${r}/$B$${capTot};0)`);
    bandTable(`A${row}:C${capTot}`, [3]);

    const barsCID = sh.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(sh.getRange(`A${capStart}:B${Math.min(capStart+14, capTot-1)}`))
      .setOption('legend', {position: 'none'})
      .setOption('title', 'Capítulos CID-10 (Top)')
      .setPosition(row, 8, 0, 0)
      .build();
    sh.insertChart(barsCID);
    row = capTot + 3;
  } else {
    const capTot = capStart;
    sh.getRange(capTot,1,1,3).setValues([['TOTAL',0,1]]).setBackground(COLOR.header).setFontWeight('bold');
    bandTable(`A${row}:C${capTot}`, [3]);
    row = capTot + 3;
  }

  // Especialidades – tabela dedup + gráfico
  sh.getRange(row,1,1,4).setValues([['Especialidade','Qtd','Média Permanência (dias)','% Óbito']])
    .setBackground(COLOR.header).setFontWeight('bold');

  const especialidades = shApoio.getRange('U2:U').getValues().flat().filter(v=>v);
  const espStart = row+1;
  if (especialidades.length > 0) {
    sh.getRange(espStart,1,especialidades.length,1).setValues(especialidades.map(v=>[v]));
    const countFormulas = Array.from({ length: especialidades.length }, () => [`=COUNTIFS(DadosÚnicos!U:U;RC1)`]);
    const avgFormulas = Array.from(
      { length: especialidades.length },
      () => [`=IFERROR(AVERAGE(FILTER(DadosÚnicos!R:R;DadosÚnicos!U:U=RC1));0)`]
    );
    const deathFormulas = Array.from(
      { length: especialidades.length },
      () => [`=IFERROR(COUNTIFS(DadosÚnicos!U:U;RC1;DadosÚnicos!O:O;"Óbito")/COUNTIFS(DadosÚnicos!U:U;RC1);0)`]
    );
    sh.getRange(espStart, 2, especialidades.length, 1).setFormulasR1C1(countFormulas);
    sh.getRange(espStart, 3, especialidades.length, 1).setFormulasR1C1(avgFormulas);
    sh.getRange(espStart, 4, especialidades.length, 1).setFormulasR1C1(deathFormulas);
    const espEnd = espStart + especialidades.length - 1;
    sh.getRange(espStart,3,especialidades.length,1).setNumberFormat('0.00');
    sh.getRange(espStart,4,especialidades.length,1).setNumberFormat('0.0%');
    sh.getRange(espStart,1,especialidades.length,4).setBorder(true,true,true,true,true,true);

    const barsEsp = sh.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(shData.getRange('E1:F'))
      .setPosition(row, 8, 0, 180)
      .setOption('title', 'Especialidades (Top)')
      .setOption('legend', {position: 'none'})
      .build();
    sh.insertChart(barsEsp);
    row = espEnd + 3;
  } else {
    row += 3;
  }

  /* Ajustes finais */
  sh.getRange(1,1,row,12).setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange('A1:L1').setFontSize(14);
  sh.setFrozenRows(1);

  SpreadsheetApp.flush();
  sh.getRange(1, 1, row, 12).copyTo(sh.getRange(1, 1, row, 12), { contentsOnly: true });

  SpreadsheetApp.flush();
  Utilities.sleep(120);

  SpreadsheetApp.getUi().alert('✅ Dashboard (V12.2.3) criado com sucesso! Lock de estabilidade + métricas únicas em Setor/Destino.');
}
