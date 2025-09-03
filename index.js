// ID da planilha do Google Sheets
const SPREADSHEET_ID = '1BmLf6vM5FsvMOUbuV8VxtS7drcnRkR4HMZXiriDhcGM';

/**
 * Função doGet - necessária para web apps do Google Apps Script
 * Esta função é chamada quando a web app é acessada via GET
 */
function doGet(e) {
  console.log('doGet chamado com parâmetros:', e.parameter);
  
  // Inicializa as abas da planilha se necessário
  try {
    console.log('Inicializando abas da planilha...');
    initializeAllSheets();
    console.log('Abas inicializadas com sucesso!');
  } catch (error) {
    console.error('Erro ao inicializar abas:', error);
  }
  
  // Retorna o HTML da aplicação
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('AgroStock Manager')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Função para lidar com requisições OPTIONS (CORS preflight)
 */
function doOptions(e) {
  console.log('doOptions chamado:', e);
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT)
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type')
    .setHeader('Access-Control-Max-Age', '86400');
}

/**
 * Função doPost - para requisições POST
 */
function doPost(e) {
  try {
    // Log de debug para verificar se a requisição está chegando
    console.log('doPost chamado:', e);
    console.log('Parâmetros recebidos:', e.parameter);
    console.log('Dados POST:', e.postData);
    
    // Verifica se há dados POST
    if (!e.postData || !e.postData.contents) {
      console.error('Nenhum dado POST recebido');
      return ContentService.createTextOutput(JSON.stringify({error: 'Nenhum dado POST recebido'}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    console.log('Ação recebida:', action, 'Dados:', data);
    
    switch (action) {
      case 'getAllData':
        const result = getAllData();
        return ContentService.createTextOutput(JSON.stringify(result))
          .setMimeType(ContentService.MimeType.JSON)
          .setHeader('Access-Control-Allow-Origin', '*')
          .setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
          .setHeader('Access-Control-Allow-Headers', 'Content-Type');
      case 'saveInsumo':
        return ContentService.createTextOutput(JSON.stringify(saveInsumo(data.insumo)))
          .setMimeType(ContentService.MimeType.JSON);
      case 'saveFornecedor':
        return ContentService.createTextOutput(JSON.stringify(saveFornecedor(data.fornecedor)))
          .setMimeType(ContentService.MimeType.JSON);
      case 'saveTalhao':
        return ContentService.createTextOutput(JSON.stringify(saveTalhao(data.talhao)))
          .setMimeType(ContentService.MimeType.JSON);
      case 'saveEntrada':
        return ContentService.createTextOutput(JSON.stringify(saveEntrada(data.entrada)))
          .setMimeType(ContentService.MimeType.JSON);
      case 'saveSaida':
        return ContentService.createTextOutput(JSON.stringify(saveSaida(data.saida)))
          .setMimeType(ContentService.MimeType.JSON);
      case 'deleteInsumo':
        return ContentService.createTextOutput(JSON.stringify(deleteInsumo(data.id)))
          .setMimeType(ContentService.MimeType.JSON);
      case 'deleteFornecedor':
        return ContentService.createTextOutput(JSON.stringify(deleteFornecedor(data.id)))
          .setMimeType(ContentService.MimeType.JSON);
      case 'deleteTalhao':
        return ContentService.createTextOutput(JSON.stringify(deleteTalhao(data.id)))
          .setMimeType(ContentService.MimeType.JSON);
      case 'deleteEntrada':
        return ContentService.createTextOutput(JSON.stringify(deleteEntrada(data.id)))
          .setMimeType(ContentService.MimeType.JSON);
      case 'deleteSaida':
        return ContentService.createTextOutput(JSON.stringify(deleteSaida(data.id)))
          .setMimeType(ContentService.MimeType.JSON);
      default:
        throw new Error('Ação não reconhecida: ' + action);
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({error: error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Nomes das abas da planilha
const SHEETS = {
  INSUMOS: 'Insumos',
  FORNECEDORES: 'Fornecedores',
  TALHOES: 'Talhoes',
  ENTRADAS: 'Entradas',
  SAIDAS: 'Saidas',
  ESTOQUE: 'Estoque'
};

/**
 * Função para obter a planilha
 */
function getSpreadsheet() {
  try {
    console.log('Tentando abrir planilha com ID:', SPREADSHEET_ID);
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    console.log('Planilha aberta com sucesso:', spreadsheet.getName());
    return spreadsheet;
  } catch (error) {
    console.error('Erro ao abrir planilha:', error);
    console.error('ID da planilha:', SPREADSHEET_ID);
    throw new Error('Não foi possível acessar a planilha. Verifique o ID: ' + error.toString());
  }
}

/**
 * Função para obter uma aba específica
 */
function getSheet(sheetName) {
  const spreadsheet = getSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  // Se a aba não existir, criar
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    initializeSheet(sheet, sheetName);
  }
  
  return sheet;
}

/**
 * Função para inicializar as abas com cabeçalhos
 */
function initializeSheet(sheet, sheetName) {
  let headers = [];
  
  switch (sheetName) {
    case SHEETS.INSUMOS:
      headers = ['ID', 'Nome', 'Categoria', 'Unidade', 'Estoque Mínimo', 'Data Criação'];
      break;
    case SHEETS.FORNECEDORES:
      headers = ['ID', 'Nome', 'Contato', 'Email', 'Endereço', 'Data Criação'];
      break;
    case SHEETS.TALHOES:
      headers = ['ID', 'Nome', 'Área (ha)', 'Cultura', 'Localização', 'Data Criação'];
      break;
    case SHEETS.ENTRADAS:
      headers = ['ID', 'Data', 'Insumo ID', 'Quantidade', 'Valor Unitário', 'Valor Total', 'Tipo', 'Fornecedor ID', 'Observações'];
      break;
    case SHEETS.SAIDAS:
      headers = ['ID', 'Data', 'Insumo ID', 'Quantidade', 'Tipo', 'Talhão ID', 'Responsável', 'Observações'];
      break;
    case SHEETS.ESTOQUE:
      headers = ['Insumo ID', 'Quantidade Atual', 'Valor Unitário Médio', 'Última Atualização'];
      break;
  }
  
  if (headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.getRange(1, 1, 1, headers.length).setBackground('#4a7c59');
    sheet.getRange(1, 1, 1, headers.length).setFontColor('white');
  }
}

/**
 * Função para gerar próximo ID
 */
function getNextId(sheetName) {
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    return 1;
  }
  
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  return Math.max(...ids) + 1;
}

/**
 * Função principal para obter todos os dados
 */
function getAllData() {
  try {
    console.log('getAllData: Iniciando carregamento de dados...');
    
    const result = {
      insumos: getInsumos(),
      fornecedores: getFornecedores(),
      talhoes: getTalhoes(),
      entradas: getEntradas(),
      saidas: getSaidas(),
      estoque: getEstoque()
    };
    
    console.log('getAllData: Dados carregados com sucesso:', {
      insumos: result.insumos.length,
      fornecedores: result.fornecedores.length,
      talhoes: result.talhoes.length,
      entradas: result.entradas.length,
      saidas: result.saidas.length,
      estoque: result.estoque.length
    });
    
    return result;
  } catch (error) {
    console.error('Erro ao carregar dados:', error);
    throw error;
  }
}

// ==================== CRUD INSUMOS ====================

/**
 * Obter todos os insumos
 */
function getInsumos() {
  const sheet = getSheet(SHEETS.INSUMOS);
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    return [];
  }
  
  const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  return data.map(row => ({
    id: row[0],
    nome: row[1],
    categoria: row[2],
    unidade: row[3],
    estoqueMinimo: row[4],
    dataCriacao: row[5]
  }));
}

/**
 * Criar novo insumo
 */
function createInsumo(insumo) {
  const sheet = getSheet(SHEETS.INSUMOS);
  const id = getNextId(SHEETS.INSUMOS);
  const dataCriacao = new Date();
  
  const newRow = [
    id,
    insumo.nome,
    insumo.categoria,
    insumo.unidade,
    insumo.estoqueMinimo,
    dataCriacao
  ];
  
  sheet.appendRow(newRow);
  
  // Inicializar estoque zerado
  initializeEstoque(id);
  
  return { id, ...insumo, dataCriacao };
}

/**
 * Atualizar insumo
 */
function updateInsumo(id, insumo) {
  const sheet = getSheet(SHEETS.INSUMOS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.getRange(i + 1, 2, 1, 4).setValues([[
        insumo.nome,
        insumo.categoria,
        insumo.unidade,
        insumo.estoqueMinimo
      ]]);
      return true;
    }
  }
  
  throw new Error('Insumo não encontrado');
}

/**
 * Deletar insumo
 */
function deleteInsumo(id) {
  const sheet = getSheet(SHEETS.INSUMOS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      // Remover do estoque também
      removeFromEstoque(id);
      return true;
    }
  }
  
  throw new Error('Insumo não encontrado');
}

// ==================== CRUD FORNECEDORES ====================

/**
 * Obter todos os fornecedores
 */
function getFornecedores() {
  const sheet = getSheet(SHEETS.FORNECEDORES);
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    return [];
  }
  
  const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  return data.map(row => ({
    id: row[0],
    nome: row[1],
    contato: row[2],
    email: row[3],
    endereco: row[4],
    dataCriacao: row[5]
  }));
}

/**
 * Criar novo fornecedor
 */
function createFornecedor(fornecedor) {
  const sheet = getSheet(SHEETS.FORNECEDORES);
  const id = getNextId(SHEETS.FORNECEDORES);
  const dataCriacao = new Date();
  
  const newRow = [
    id,
    fornecedor.nome,
    fornecedor.contato,
    fornecedor.email,
    fornecedor.endereco || '',
    dataCriacao
  ];
  
  sheet.appendRow(newRow);
  return { id, ...fornecedor, dataCriacao };
}

/**
 * Atualizar fornecedor
 */
function updateFornecedor(id, fornecedor) {
  const sheet = getSheet(SHEETS.FORNECEDORES);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.getRange(i + 1, 2, 1, 4).setValues([[
        fornecedor.nome,
        fornecedor.contato,
        fornecedor.email,
        fornecedor.endereco || ''
      ]]);
      return true;
    }
  }
  
  throw new Error('Fornecedor não encontrado');
}

/**
 * Deletar fornecedor
 */
function deleteFornecedor(id) {
  const sheet = getSheet(SHEETS.FORNECEDORES);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  
  throw new Error('Fornecedor não encontrado');
}

// ==================== CRUD TALHÕES ====================

/**
 * Obter todos os talhões
 */
function getTalhoes() {
  const sheet = getSheet(SHEETS.TALHOES);
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    return [];
  }
  
  const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  return data.map(row => ({
    id: row[0],
    nome: row[1],
    area: row[2],
    cultura: row[3],
    localizacao: row[4],
    dataCriacao: row[5]
  }));
}

/**
 * Criar novo talhão
 */
function createTalhao(talhao) {
  const sheet = getSheet(SHEETS.TALHOES);
  const id = getNextId(SHEETS.TALHOES);
  const dataCriacao = new Date();
  
  const newRow = [
    id,
    talhao.nome,
    talhao.area,
    talhao.cultura,
    talhao.localizacao || '',
    dataCriacao
  ];
  
  sheet.appendRow(newRow);
  return { id, ...talhao, dataCriacao };
}

/**
 * Atualizar talhão
 */
function updateTalhao(id, talhao) {
  const sheet = getSheet(SHEETS.TALHOES);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.getRange(i + 1, 2, 1, 4).setValues([[
        talhao.nome,
        talhao.area,
        talhao.cultura,
        talhao.localizacao || ''
      ]]);
      return true;
    }
  }
  
  throw new Error('Talhão não encontrado');
}

/**
 * Deletar talhão
 */
function deleteTalhao(id) {
  const sheet = getSheet(SHEETS.TALHOES);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  
  throw new Error('Talhão não encontrado');
}

// ==================== ENTRADAS ====================

/**
 * Obter todas as entradas
 */
function getEntradas() {
  const sheet = getSheet(SHEETS.ENTRADAS);
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    return [];
  }
  
  const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  return data.map(row => ({
    id: row[0],
    data: row[1],
    insumoId: row[2],
    quantidade: row[3],
    valorUnitario: row[4],
    valorTotal: row[5],
    tipo: row[6],
    fornecedorId: row[7],
    observacoes: row[8]
  }));
}

/**
 * Criar nova entrada
 */
function createEntrada(entrada) {
  const sheet = getSheet(SHEETS.ENTRADAS);
  const id = getNextId(SHEETS.ENTRADAS);
  const valorTotal = entrada.quantidade * entrada.valorUnitario;
  
  const newRow = [
    id,
    entrada.data,
    entrada.insumoId,
    entrada.quantidade,
    entrada.valorUnitario,
    valorTotal,
    entrada.tipo,
    entrada.fornecedorId,
    entrada.observacoes || ''
  ];
  
  sheet.appendRow(newRow);
  
  // Atualizar estoque
  updateEstoque(entrada.insumoId, entrada.quantidade, entrada.valorUnitario, 'entrada');
  
  return { id, ...entrada, valorTotal };
}

/**
 * Deletar entrada
 */
function deleteEntrada(id) {
  const sheet = getSheet(SHEETS.ENTRADAS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      const entrada = data[i];
      sheet.deleteRow(i + 1);
      
      // Reverter estoque
      updateEstoque(entrada[2], -entrada[3], entrada[4], 'entrada');
      return true;
    }
  }
  
  throw new Error('Entrada não encontrada');
}

// ==================== SAÍDAS ====================

/**
 * Obter todas as saídas
 */
function getSaidas() {
  const sheet = getSheet(SHEETS.SAIDAS);
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    return [];
  }
  
  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  return data.map(row => ({
    id: row[0],
    data: row[1],
    insumoId: row[2],
    quantidade: row[3],
    tipo: row[4],
    talhaoId: row[5],
    responsavel: row[6],
    observacoes: row[7]
  }));
}

/**
 * Criar nova saída
 */
function createSaida(saida) {
  const sheet = getSheet(SHEETS.SAIDAS);
  const id = getNextId(SHEETS.SAIDAS);
  
  // Verificar se há estoque suficiente
  const estoqueAtual = getEstoqueByInsumo(saida.insumoId);
  if (!estoqueAtual || estoqueAtual.quantidade < saida.quantidade) {
    throw new Error('Estoque insuficiente para esta operação');
  }
  
  const newRow = [
    id,
    saida.data,
    saida.insumoId,
    saida.quantidade,
    saida.tipo,
    saida.talhaoId,
    saida.responsavel || '',
    saida.observacoes || ''
  ];
  
  sheet.appendRow(newRow);
  
  // Atualizar estoque
  updateEstoque(saida.insumoId, -saida.quantidade, 0, 'saida');
  
  return { id, ...saida };
}

/**
 * Deletar saída
 */
function deleteSaida(id) {
  const sheet = getSheet(SHEETS.SAIDAS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      const saida = data[i];
      sheet.deleteRow(i + 1);
      
      // Reverter estoque
      updateEstoque(saida[2], saida[3], 0, 'saida');
      return true;
    }
  }
  
  throw new Error('Saída não encontrada');
}

// ==================== CONTROLE DE ESTOQUE ====================

/**
 * Obter estoque completo
 */
function getEstoque() {
  const sheet = getSheet(SHEETS.ESTOQUE);
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    return [];
  }
  
  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  return data.map(row => ({
    insumoId: row[0],
    quantidade: row[1],
    valorUnitario: row[2],
    ultimaAtualizacao: row[3]
  }));
}

/**
 * Obter estoque de um insumo específico
 */
function getEstoqueByInsumo(insumoId) {
  const sheet = getSheet(SHEETS.ESTOQUE);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == insumoId) {
      return {
        insumoId: data[i][0],
        quantidade: data[i][1],
        valorUnitario: data[i][2],
        ultimaAtualizacao: data[i][3]
      };
    }
  }
  
  return null;
}

/**
 * Inicializar estoque para novo insumo
 */
function initializeEstoque(insumoId) {
  const sheet = getSheet(SHEETS.ESTOQUE);
  const newRow = [
    insumoId,
    0,
    0,
    new Date()
  ];
  
  sheet.appendRow(newRow);
}

/**
 * Atualizar estoque
 */
function updateEstoque(insumoId, quantidade, valorUnitario, tipo) {
  const sheet = getSheet(SHEETS.ESTOQUE);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == insumoId) {
      const quantidadeAtual = data[i][1];
      const valorAtual = data[i][2];
      
      let novaQuantidade = quantidadeAtual + quantidade;
      let novoValor = valorAtual;
      
      // Calcular novo valor médio apenas para entradas
      if (tipo === 'entrada' && quantidade > 0) {
        const valorTotalAtual = quantidadeAtual * valorAtual;
        const valorTotalNovo = quantidade * valorUnitario;
        novoValor = novaQuantidade > 0 ? (valorTotalAtual + valorTotalNovo) / novaQuantidade : 0;
      }
      
      sheet.getRange(i + 1, 2, 1, 3).setValues([[
        novaQuantidade,
        novoValor,
        new Date()
      ]]);
      
      return true;
    }
  }
  
  // Se não encontrou, criar novo registro
  if (tipo === 'entrada') {
    initializeEstoque(insumoId);
    updateEstoque(insumoId, quantidade, valorUnitario, tipo);
  }
  
  return false;
}

/**
 * Remover insumo do estoque
 */
function removeFromEstoque(insumoId) {
  const sheet = getSheet(SHEETS.ESTOQUE);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == insumoId) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  
  return false;
}

// ==================== FUNÇÕES AUXILIARES ====================

/**
 * Função para testar a conexão
 */
function testConnection() {
  try {
    const spreadsheet = getSpreadsheet();
    return {
      success: true,
      message: 'Conexão estabelecida com sucesso!',
      spreadsheetName: spreadsheet.getName()
    };
  } catch (error) {
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * Função para inicializar todas as abas
 */
function initializeAllSheets() {
  Object.values(SHEETS).forEach(sheetName => {
    getSheet(sheetName);
  });
  
  return 'Todas as abas foram inicializadas com sucesso!';
}

/**
 * Função para obter relatório de estoque baixo
 */
function getEstoqueBaixo() {
  const insumos = getInsumos();
  const estoque = getEstoque();
  
  const estoqueBaixo = [];
  
  estoque.forEach(item => {
    const insumo = insumos.find(i => i.id === item.insumoId);
    if (insumo && item.quantidade <= insumo.estoqueMinimo) {
      estoqueBaixo.push({
        insumo: insumo.nome,
        quantidadeAtual: item.quantidade,
        estoqueMinimo: insumo.estoqueMinimo,
        unidade: insumo.unidade
      });
    }
  });
  
  return estoqueBaixo;
}

/**
 * Função para obter estatísticas gerais
 */
function getEstatisticas() {
  return {
    totalInsumos: getInsumos().length,
    totalFornecedores: getFornecedores().length,
    totalTalhoes: getTalhoes().length,
    totalItensEstoque: getEstoque().length,
    estoqueBaixo: getEstoqueBaixo().length
  };
}

/**
 * Função para excluir insumo
 */
function deleteInsumo(id) {
  const sheet = getSheet(SHEETS.INSUMOS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Insumo excluído com sucesso' };
    }
  }
  
  return { success: false, message: 'Insumo não encontrado' };
}

/**
 * Função para excluir fornecedor
 */
function deleteFornecedor(id) {
  const sheet = getSheet(SHEETS.FORNECEDORES);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Fornecedor excluído com sucesso' };
    }
  }
  
  return { success: false, message: 'Fornecedor não encontrado' };
}

/**
 * Função para excluir talhão
 */
function deleteTalhao(id) {
  const sheet = getSheet(SHEETS.TALHOES);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Talhão excluído com sucesso' };
    }
  }
  
  return { success: false, message: 'Talhão não encontrado' };
}

/**
 * Função para excluir entrada
 */
function deleteEntrada(id) {
  const sheet = getSheet(SHEETS.ENTRADAS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Entrada excluída com sucesso' };
    }
  }
  
  return { success: false, message: 'Entrada não encontrada' };
}

/**
 * Função para excluir saída
 */
function deleteSaida(id) {
  const sheet = getSheet(SHEETS.SAIDAS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Saída excluída com sucesso' };
    }
  }
  
  return { success: false, message: 'Saída não encontrada' };
}

/**
 * Função auxiliar para incluir arquivos HTML
 * Necessária para o Google Apps Script servir arquivos HTML
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}