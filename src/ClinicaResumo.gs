/**
 * Utilitários para gerar os painéis de Entrada e Alta da Clínica.
 *
 * O script lê os dados completos da aba "Base Filtrada (Fórmula)" e os
 * correspondentes deduplicados da aba "DadosÚnicos" para montar duas tabelas:
 *
 * - Entrada: contagem por Setor (coluna N).
 * - Alta: contagem por Destino (coluna O) apenas para "Óbito", "residência" e
 *   "outro hospital".
 *
 * As tabelas possuem cinco colunas: Descrição, Qtd (total), %, Qtd (únicos) e
 * % (únicos). As porcentagens são fornecidas como valores decimais (0 a 1) para
 * que sejam exibidas como porcentagem na planilha.
 */

const CLINICA_SHEETS = {
  baseCompleta: 'Base Filtrada (Fórmula)',
  baseUnica: 'DadosÚnicos',
};

const CLINICA_TABELAS = {
  entrada: {
    descricao: 'Clínica Entrada',
    coluna: 'N',
    destino: {
      sheet: 'Clínica Entrada',
      startRow: 1,
      startCol: 1,
      clearHeight: 50,
    },
  },
  alta: {
    descricao: 'Clínica Alta',
    coluna: 'O',
    destinosValidos: [
      'Óbito',
      'residência',
      'outro hospital',
    ],
    destino: {
      sheet: 'Clínica Alta',
      startRow: 1,
      startCol: 1,
      clearHeight: 20,
    },
  },
};

const TABELA_HEADER = ['Descrição', 'Qtd (total)', '%', 'Qtd (únicos)', '% (únicos)'];

/**
 * Atualiza os resumos de Entrada e Alta na planilha configurada acima.
 */
function atualizarTabelasClinica() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseCompleta = obterAbaObrigatoria_(ss, CLINICA_SHEETS.baseCompleta);
  const baseUnica = obterAbaObrigatoria_(ss, CLINICA_SHEETS.baseUnica);

  const entradaTabela = montarTabelaEntrada_(baseCompleta, baseUnica);
  const altaTabela = montarTabelaAlta_(baseCompleta, baseUnica);

  escreverTabela_(ss, CLINICA_TABELAS.entrada.destino, entradaTabela);
  escreverTabela_(ss, CLINICA_TABELAS.alta.destino, altaTabela);
}

/**
 * Monta a tabela de entrada (por setor / coluna N).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} baseCompleta Aba com a base
 *     completa.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} baseUnica Aba com os dados
 *     deduplicados.
 * @return {!Array<!Array<*>>} Matriz representando a tabela.
 */
function montarTabelaEntrada_(baseCompleta, baseUnica) {
  const coluna = colunaParaIndice_(CLINICA_TABELAS.entrada.coluna);
  const valoresTotais = extrairColuna_(baseCompleta, coluna);
  const valoresUnicos = extrairColuna_(baseUnica, coluna);

  const frequenciaTotal = contarValores_(valoresTotais);
  const frequenciaUnica = contarValores_(valoresUnicos);

  const chaves = new Set([...frequenciaTotal.keys(), ...frequenciaUnica.keys()]);
  const linhas = Array.from(chaves)
    .map((chave) => {
      const totalInfo = frequenciaTotal.get(chave);
      const unicoInfo = frequenciaUnica.get(chave);
      const descricao = (totalInfo || unicoInfo).label;
      const total = totalInfo ? totalInfo.count : 0;
      const unicos = unicoInfo ? unicoInfo.count : 0;
      return { chave, descricao, total, unicos };
    })
    .sort((a, b) => {
      if (b.total !== a.total) {
        return b.total - a.total;
      }
      return a.descricao.localeCompare(b.descricao);
    });

  const totalGeral = linhas.reduce((acc, item) => acc + item.total, 0);
  const totalUnicos = linhas.reduce((acc, item) => acc + item.unicos, 0);

  const corpo = linhas.map((item) => [
    item.descricao,
    item.total,
    calcularPercentual_(item.total, totalGeral),
    item.unicos,
    calcularPercentual_(item.unicos, totalUnicos),
  ]);

  const totalRow = [
    'Total',
    totalGeral,
    totalGeral > 0 ? 1 : 0,
    totalUnicos,
    totalUnicos > 0 ? 1 : 0,
  ];

  return [TABELA_HEADER, ...corpo, totalRow];
}

/**
 * Monta a tabela de alta (por destino / coluna O).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} baseCompleta Aba com a base
 *     completa.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} baseUnica Aba com os dados
 *     deduplicados.
 * @return {!Array<!Array<*>>} Matriz representando a tabela.
 */
function montarTabelaAlta_(baseCompleta, baseUnica) {
  const coluna = colunaParaIndice_(CLINICA_TABELAS.alta.coluna);
  const valoresTotais = extrairColuna_(baseCompleta, coluna);
  const valoresUnicos = extrairColuna_(baseUnica, coluna);

  const destinos = CLINICA_TABELAS.alta.destinosValidos.map((label) => ({
    label,
    chave: normalizarTexto_(label),
  }));

  const totalCounts = contarValoresFiltrados_(valoresTotais, destinos);
  const unicoCounts = contarValoresFiltrados_(valoresUnicos, destinos);

  const totalGeral = destinos.reduce((acc, destino) => acc + (totalCounts.get(destino.chave) || 0), 0);
  const totalUnicos = destinos.reduce((acc, destino) => acc + (unicoCounts.get(destino.chave) || 0), 0);

  const corpo = destinos.map((destino) => {
    const total = totalCounts.get(destino.chave) || 0;
    const unicos = unicoCounts.get(destino.chave) || 0;
    return [
      destino.label,
      total,
      calcularPercentual_(total, totalGeral),
      unicos,
      calcularPercentual_(unicos, totalUnicos),
    ];
  });

  const totalRow = [
    'Total',
    totalGeral,
    totalGeral > 0 ? 1 : 0,
    totalUnicos,
    totalUnicos > 0 ? 1 : 0,
  ];

  return [TABELA_HEADER, ...corpo, totalRow];
}

/**
 * Extrai todos os valores de uma coluna a partir da segunda linha.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Aba de onde extrair.
 * @param {number} columnIndex Índice da coluna (1 baseado).
 * @return {!Array<string>} Valores não nulos extraídos da coluna.
 */
function extrairColuna_(sheet, columnIndex) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  const valores = sheet
    .getRange(2, columnIndex, lastRow - 1, 1)
    .getValues()
    .map((row) => row[0])
    .filter((value) => value !== '' && value !== null && value !== undefined);

  return valores;
}

/**
 * Conta as ocorrências de cada valor (normalizado) preservando o primeiro rótulo.
 *
 * @param {!Array<string>} valores Lista de valores.
 * @return {!Map<string, {label: string, count: number}>}
 */
function contarValores_(valores) {
  const mapa = new Map();

  valores.forEach((valor) => {
    const chave = normalizarTexto_(valor);
    if (!chave) {
      return;
    }

    if (!mapa.has(chave)) {
      mapa.set(chave, { label: valor.toString().trim(), count: 0 });
    }

    const info = mapa.get(chave);
    info.count += 1;
  });

  return mapa;
}

/**
 * Conta apenas valores presentes na lista permitida.
 *
 * @param {!Array<string>} valores Lista de valores a serem contados.
 * @param {!Array<{label: string, chave: string}>} destinos Lista de destinos
 *     aceitos.
 * @return {!Map<string, number>} Contagem por chave normalizada.
 */
function contarValoresFiltrados_(valores, destinos) {
  const permitido = new Map(destinos.map((destino) => [destino.chave, 0]));

  valores.forEach((valor) => {
    const chave = normalizarTexto_(valor);
    if (!chave || !permitido.has(chave)) {
      return;
    }

    permitido.set(chave, permitido.get(chave) + 1);
  });

  return permitido;
}

/**
 * Converte uma letra de coluna em índice numérico (1 baseado).
 *
 * @param {string} letra Letra da coluna.
 * @return {number} Índice numérico correspondente.
 */
function colunaParaIndice_(letra) {
  let indice = 0;
  const texto = letra.toUpperCase();
  for (let i = 0; i < texto.length; i += 1) {
    indice = indice * 26 + (texto.charCodeAt(i) - 64);
  }
  return indice;
}

/**
 * Normaliza textos removendo espaços extras, acentuação e usando minúsculas.
 *
 * @param {*} valor Valor a normalizar.
 * @return {string} Texto normalizado.
 */
function normalizarTexto_(valor) {
  if (valor === null || valor === undefined) {
    return '';
  }

  return valor
    .toString()
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

/**
 * Calcula um percentual baseado em total.
 *
 * @param {number} parte Parte do total.
 * @param {number} total Valor total.
 * @return {number} Percentual em formato decimal (0 a 1).
 */
function calcularPercentual_(parte, total) {
  if (!total) {
    return 0;
  }
  return parte / total;
}

/**
 * Recupera uma aba obrigatória, lançando erro caso não exista.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss Planilha ativa.
 * @param {string} nome Nome da aba desejada.
 * @return {GoogleAppsScript.Spreadsheet.Sheet} Aba encontrada.
 */
function obterAbaObrigatoria_(ss, nome) {
  const sheet = ss.getSheetByName(nome);
  if (!sheet) {
    throw new Error('A aba "' + nome + '" não foi encontrada.');
  }
  return sheet;
}

/**
 * Escreve a tabela na aba de destino configurada.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss Planilha ativa.
 * @param {{sheet: string, startRow: number, startCol: number, clearHeight: (number|undefined)}} destino
 *     Configuração da área de escrita.
 * @param {!Array<!Array<*>>} tabela Dados a gravar.
 */
function escreverTabela_(ss, destino, tabela) {
  const sheet = obterAbaObrigatoria_(ss, destino.sheet);
  const largura = TABELA_HEADER.length;
  const altura = tabela.length;
  const clearHeight = destino.clearHeight || Math.max(altura + 5, 10);

  sheet
    .getRange(destino.startRow, destino.startCol, clearHeight, largura)
    .clearContent();

  sheet
    .getRange(destino.startRow, destino.startCol, altura, largura)
    .setValues(tabela);

  // Cabeçalho em negrito.
  sheet
    .getRange(destino.startRow, destino.startCol, 1, largura)
    .setFontWeight('bold');

  if (altura > 1) {
    const primeiraLinhaDados = destino.startRow + 1;
    const linhasDados = altura - 1;
    [destino.startCol + 2, destino.startCol + 4].forEach((colunaPercentual) => {
      sheet
        .getRange(primeiraLinhaDados, colunaPercentual, linhasDados, 1)
        .setNumberFormat('0.0%');
    });
  }
}
