/************************************************************
📊 DASHBOARD EPIDEMIOLÓGICO – Luky + GPT-5 (V12.2.3 – Hotfix estabilidade + métricas únicas)
• Funções em inglês nas fórmulas; separador de argumentos ";"
• Deduplicação por Prontuário (C) usando a última Data Saída (Q)
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
  shUni.getRange('A1').setValue('⚙️ Base deduplicada por prontuário (última ocorrência pela Data Saída)')
       .setFontWeight('bold').setFontColor(COLOR.textMuted);
  shUni.getRange('A2').setFormula(
    "=UNIQUE(SORTN('Base Filtrada (Fórmula)'!A2:Y;9^9;2;'Base Filtrada (Fórmula)'!C2:C;TRUE;'Base Filtrada (Fórmula)'!Q2:Q;FALSE))"
  );
  SpreadsheetApp.flush();
  Utilities.sleep(120);
  shUni.hideSheet();

  /* ===== 2) ⚙️DATA (séries p/ gráficos e auxiliares) ===== */
  const shData = safeRecreateSheet_(ss, '⚙️DATA', shBase);

  // Fluxo Entradas × Altas – robusto
  shData.getRange('A1:C1').setValues([['Datas (período)','Entradas (dia)','Altas (dia)']]).setFontWeight('bold');
  shData.getRange('A2').setFormula(
    "=UNIQUE(SORT({" +
      "FILTER('Base Filtrada (Fórmula)'!P2:P;'Base Filtrada (Fórmula)'!P2:P<>\"\");" +
      "FILTER('Base Filtrada (Fórmula)'!Q2:Q;'Base Filtrada (Fórmula)'!Q2:Q<>\"\")" +
    "}))"
  );
  shData.getRange('B2').setFormula("=ARRAYFORMULA(IF(A2:A=\"\";;COUNTIF('Base Filtrada (Fórmula)'!P:P;A2:A)))");
  shData.getRange('C2').setFormula("=ARRAYFORMULA(IF(A2:A=\"\";;COUNTIF('Base Filtrada (Fórmula)'!Q:Q;A2:A)))");

  // Especialidades (dedup)
  shData.getRange('E1:F1').setValues([['Especialidade','Qtd (dedup)']]).setFontWeight('bold');
  shData.getRange('E2').setFormula("=FILTER('LISTAS DE APOIO'!U2:U;'LISTAS DE APOIO'!U2:U<>\"\")");
  shData.getRange('F2').setFormula("=ARRAYFORMULA(IF(E2:E=\"\";;COUNTIFS(DadosÚnicos!U:U;E2:E)))");

  // Catálogo de CIDs (referência)
  shData.getRange('L1:N1').setValues([['Capítulo CID10 (catálogo)','Código (catálogo)','—']]).setFontWeight('bold');
  shData.getRange('L2').setFormula("=FILTER('Cadastro CIDS'!B2:B;'Cadastro CIDS'!B2:B<>\"\")");
  shData.getRange('M2').setFormula("=FILTER('Cadastro CIDS'!G2:G;'Cadastro CIDS'!G2:G<>\"\")");

  // Capítulos do CID-10 (via VLOOKUP+QUERY sobre dedup)
  shData.getRange('O1:P1').setValues([['Capítulo (uso)','Qtd (uso)']]).setFontWeight('bold');
  shData.getRange('O2').setFormula(
    "=QUERY(" +
      "ARRAYFORMULA(" +
        "IFNA(VLOOKUP(" +
          "FILTER(DadosÚnicos!S2:S;DadosÚnicos!S2:S<>\"\");" +
          "HSTACK('Cadastro CIDS'!G2:G;'Cadastro CIDS'!B2:B);" +
          "2;FALSE" +
        "))" +
      ");" +
      "\"select Col1,count(Col1) where Col1 is not null group by Col1 order by count(Col1) desc label count(Col1) ''\";" +
      "0)"
  );

  SpreadsheetApp.flush();
  Utilities.sleep(150);
  shData.hideSheet();

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
  function miniMuted(r, c, label, formulaOrValue, fmt) {
    sh.getRange(r, c).setValue(label).setFontColor(COLOR.textMuted);
    const cell = sh.getRange(r, c+1);
    if (typeof formulaOrValue === 'string' && formulaOrValue.startsWith('=')) cell.setFormula(formulaOrValue);
    else cell.setValue(formulaOrValue);
    cell.setFontWeight('bold'); if (fmt) cell.setNumberFormat(fmt);
  }
  function kpiCard(r, c, title, formula, fmt, icon='') {
    const titleR = sh.getRange(r, c, 1, 3).merge();
    const valueR = sh.getRange(r+1, c, 1, 3).merge();
    titleR.setValue(`${icon} ${title}`).setFontWeight('bold')
      .setBackground(COLOR.header).setFontColor('#0C3D3A')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    valueR.setFormula(formula).setFontSize(18).setFontWeight('bold')
      .setBackground('#FFFFFF').setHorizontalAlignment('left').setVerticalAlignment('middle');
    if (fmt) valueR.setNumberFormat(fmt);
    sh.getRange(r, c, 2, 3).setBorder(true,true,true,true,true,true);
  }
  function linkTo(r, c, text, anchorCellA1) {
    const gid = sh.getSheetId();
    sh.getRange(r, c).setFormula(`=HYPERLINK("#gid=${gid}&range=${anchorCellA1}"; "${text}")`)
      .setFontColor(COLOR.primaryDark).setFontWeight('bold');
  }

  let row = 1;

  /* Header */
  headerBlock(row, 1, '📊 Dashboard Epidemiológico – HUC'); row += 2;

  // Período + metadata
  miniMuted(row, 1, 'Período selecionado',
    "='PERFIL EPIDEMIOLÓGICO'!I1 & \" – \" & 'PERFIL EPIDEMIOLÓGICO'!J1 & \" / \" & 'PERFIL EPIDEMIOLÓGICO'!K1");
  sh.getRange('B4').setFormula("='PERFIL EPIDEMIOLÓGICO'!I1 & \" – \" & 'PERFIL EPIDEMIOLÓGICO'!J1 & \" / \" & 'PERFIL EPIDEMIOLÓGICO'!K1");
  row++;

  /* KPI cards – linha 1 */
  const kpiRow1 = row;
  kpiCard(kpiRow1, 1,  'Pacientes Únicos',            "=COUNTA(DadosÚnicos!C2:C)", '#,##0', '👤');
  kpiCard(kpiRow1, 4,  'Total de Internações',        "=COUNTA('Base Filtrada (Fórmula)'!C2:C)", '#,##0', '🏥');
  kpiCard(kpiRow1, 7,  'Taxa de Óbito',               "=IFERROR(COUNTIFS(DadosÚnicos!O:O;\"Óbito\")/COUNTA(DadosÚnicos!C2:C);0)", '0.0%', '☠️');
  kpiCard(kpiRow1, 10, 'Média de Permanência (dias)', "=AVERAGE(DadosÚnicos!R:R)", '0.00', '⏱️');
  row = kpiRow1 + 3;

  /* KPI cards – linha 2 */
  const kpiRow2 = row;
  kpiCard(kpiRow2, 1,  'Primeira Internação', "=MIN(DadosÚnicos!P:P)", 'dd/mm/yyyy', '📅');
  kpiCard(kpiRow2, 4,  'Última Alta/Saída',   "=MAX(DadosÚnicos!Q:Q)", 'dd/mm/yyyy', '📅');
  kpiCard(kpiRow2, 7,  'Idade Média',         "=AVERAGE(DadosÚnicos!F:F)", '0.0', '👶');
  kpiCard(kpiRow2, 10, 'Dias-Paciente (soma)',"=SUM(DadosÚnicos!R:R)", '#,##0', '📈');
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
  sh.getRange(fStart,1,3,1).setValues([['≤ 19 anos'],['20 a 59 anos'],['≥ 60 anos']]);
  sh.getRange(fStart, 2, 3, 1).setFormulas([
    ['=COUNTIFS(DadosÚnicos!F:F;"<=19")'],
    ['=COUNTIFS(DadosÚnicos!F:F;">=20";DadosÚnicos!F:F;"<=59")'],
    ['=COUNTIFS(DadosÚnicos!F:F;">=60")'],
  ]);
  const fTot = fStart+3;
  sh.getRange(fTot,1,1,3).setValues([['TOTAL','','']]).setBackground(COLOR.header).setFontWeight('bold');
  sh.getRange(fTot,2).setFormula(`=SUM(B${fStart}:B${fStart+2})`);
  sh.getRange(fStart, 3, 3, 1)
    .setFormulasR1C1(Array.from({ length: 3 }, () => [`=IFERROR(RC[-1]/R${fTot}C2;0)`]));
  bandTable(`A${rDemog}:C${fTot}`, [3]);

  rDemog = fTot + 2;
  miniMuted(rDemog,1,'Idade média',   "=AVERAGE(DadosÚnicos!F:F)","0.0"); rDemog++;
  miniMuted(rDemog,1,'Idade mediana', "=MEDIAN(DadosÚnicos!F:F)","0.0"); rDemog += 2;

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
    const countFormulas = Array.from({ length: labels.length }, () => [`=COUNTIFS('Base Filtrada (Fórmula)'!${colBase}:${colBase};RC1)`]);
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
    const countFormulas = Array.from({ length: labels.length }, () => [`=COUNTIFS('Base Filtrada (Fórmula)'!N:N;RC1)`]);
    const uniqueFormulas = Array.from(
      { length: labels.length },
      () => [`=IFERROR(COUNTUNIQUE(FILTER('Base Filtrada (Fórmula)'!C:C;'Base Filtrada (Fórmula)'!N:N=RC1));0)`]
    );
    sh.getRange(startRow + 1, 2, labels.length, 1).setFormulasR1C1(countFormulas);
    sh.getRange(startRow + 1, 4, labels.length, 1).setFormulasR1C1(uniqueFormulas);
    sh.getRange(total, 1, 1, 5)
      .setValues([['TOTAL', '', '', '', '']])
      .setBackground(COLOR.header).setFontWeight('bold');
    sh.getRange(total, 2).setFormula(`=SUM(B${startRow + 1}:B${end})`);
    sh.getRange(total, 4).setFormula("=IFERROR(COUNTUNIQUE(FILTER('Base Filtrada (Fórmula)'!C:C;'Base Filtrada (Fórmula)'!N:N<>\"\"));0)");
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
    const countFormulas = destinos.map(() => [`=COUNTIFS('Base Filtrada (Fórmula)'!O:O;RC1)`]);
    const uniqueFormulas = destinos.map(() => [`=IFERROR(COUNTUNIQUE(FILTER('Base Filtrada (Fórmula)'!C:C;'Base Filtrada (Fórmula)'!O:O=RC1));0)`]);
    sh.getRange(startRow + 1, 2, destinos.length, 1).setFormulasR1C1(countFormulas);
    sh.getRange(startRow + 1, 4, destinos.length, 1).setFormulasR1C1(uniqueFormulas);
    sh.getRange(total, 1, 1, 5)
      .setValues([['TOTAL', '', '', '', '']])
      .setBackground(COLOR.header).setFontWeight('bold');
    sh.getRange(total, 2).setFormula(`=SUM(B${startRow + 1}:B${end})`);
    sh.getRange(total, 4).setFormula(
      "=IFERROR(COUNTUNIQUE(FILTER('Base Filtrada (Fórmula)'!C:C;REGEXMATCH('Base Filtrada (Fórmula)'!O:O;\"^(Óbito|Residência|Outro hospital)$\")));0)"
    );
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
  Utilities.sleep(120);

  SpreadsheetApp.getUi().alert('✅ Dashboard (V12.2.3) criado com sucesso! Lock de estabilidade + métricas únicas em Setor/Destino.');
}
