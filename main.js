// Desenvolvido por Paulo J C Piva - cremascopaulo@gmail.com

// Nome do projeto
const projectName = "Pipedrive + RD Station"

// Lista de e-mails
const recipient = "cremascopaulo@gmail.com , pivapaulo@outlook.com";

// Nomes das sheets
const pipedriveDeals = "pipedrive_deals";
const pipedriveDealsPersonIdEmail = "pipedrive_deals_person_id_email";
const rdStationLeads = "rdstation_leads";
const dealsLeads = "deals_leads";

// Conexão das abas das planilhas
function conectToSheets(){
    // Acessa a planilha
    function getSheet(idSpreadsheet , sheetName) {
    return SpreadsheetApp.openById(idSpreadsheet).getSheetByName(sheetName)
  }

  // Declaração aos id's das planilhas
  const idPipedrive = "13CWhdq6ZxqkQ_QVY8etKoA6hNnAN1U8hfu2OR_mie7w"; // "1gjvt6AKzFDTEA86YY7RAjiqLjKTzauWxbKmRyON81V0";
  const idRdStation = "1FwSauFMqOzWPPIID1jPC3QiikzrG5gIJ9Yl138KhpYQ"; //"1pS82Wi6SW3OhsG0f3qfteWdnsbiZjXgzv8S8TLyA51I";
  const idPipeRd = "12rjFnmKqxvYu7-Y7juo5BLfA354wVW0SimYgalLcOto";

  // Conexão as abas das planilhas
  const sheetPipedriveDeals = getSheet(idPipedrive , pipedriveDeals);
  const sheetPipedriveIdEmails = getSheet(idPipedrive , pipedriveDealsPersonIdEmail);
  const sheetRdStationLeads =  getSheet(idRdStation , rdStationLeads);
  const sheetDealsLeadsPipeRd = getSheet(idPipeRd , dealsLeads);
  return {sheetPipedriveDeals , sheetPipedriveIdEmails , sheetRdStationLeads , sheetDealsLeadsPipeRd}
}


// Extrai cabeçalhos e dados de uma planilha
function getHeadersAndData(sheet) {
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  return {
    headers: values[0],
    data: values.slice(1),
  };
}

// Extrai IDs e e-mails da planilha Pipedrive
function getIdsEmails(sheetPipedriveIdEmails) {
  const { headers, data } = getHeadersAndData(sheetPipedriveIdEmails);

  const idIndex = headers.indexOf("id");
  const emailIndex = headers.indexOf("person_id_email_value");
  const checkedIndex = headers.indexOf("checked");

  if (idIndex === -1) {
    emailNotification()
    throw new Error("Colunas necessárias não encontradas na planilha.");
  }

  if (emailIndex === -1) {
    emailNotification()
    throw new Error("Colunas necessárias não encontradas na planilha.");
  }

  if (checkedIndex === -1) {
    emailNotification()
    throw new Error("Colunas necessárias não encontradas na planilha.");
  }

  return data
    .filter(row => row[emailIndex].trim() !== "" && row[checkedIndex] !== true)
    .map(row => ({ id: row[idIndex], email: row[emailIndex].trim() }));
}


// Filtra leads da RD Station com base nos e-mails fornecidos
function getFilteredLeadsByEmail(sheetRdStationLeads, filteredIdsEmails) {
  const { headers, data } = getHeadersAndData(sheetRdStationLeads);

  const emailIndex = headers.indexOf("email");
  const createdAtIndex = headers.indexOf("created_at");

  if (emailIndex === -1 || createdAtIndex === -1) {
    throw new Error("Colunas necessárias (email, created_at) não encontradas na planilha.");
  }

  const filteredEmailsSet = new Set(filteredIdsEmails.map(item => item.email));
  const latestLeadsMap = new Map();

  data.forEach(row => {
    const email = row[emailIndex].trim();
    const createdAt = new Date(row[createdAtIndex]);

    if (filteredEmailsSet.has(email)) {
      if (!latestLeadsMap.has(email) || createdAt > latestLeadsMap.get(email).createdAt) {
        latestLeadsMap.set(email, { createdAt, row });
      }
    }
  });

  return [headers, ...Array.from(latestLeadsMap.values()).map(entry => entry.row)];
}


// Filtra IDs e e-mails com base nos leads existentes na RD Station
function filterIdsEmailsByExistingLeads(filteredIdsEmails, filteredRdStationLeads) {
  if (filteredRdStationLeads.length < 2) return [];

  const headers = filteredRdStationLeads[0];
  const emailIndex = headers.indexOf("email");

  if (emailIndex === -1) {
    throw new Error("Coluna de e-mail não encontrada em filteredRdStationLeads.");
  }

  const existingEmailsSet = new Set(filteredRdStationLeads.slice(1).map(row => row[emailIndex].trim()));
  return filteredIdsEmails.filter(item => existingEmailsSet.has(item.email));
}


// Filtra deals da Pipedrive com base nos IDs fornecidos
function getFilteredDeals(sheetPipedriveDeals, filteredIdsEmails) {
  const { headers, data } = getHeadersAndData(sheetPipedriveDeals);

  const idIndex = headers.indexOf("id");
  if (idIndex === -1) {
    throw new Error("Coluna 'id' não encontrada na planilha.");
  }

  const filteredIdsSet = new Set(filteredIdsEmails.map(item => item.id.toString()));
  const filteredDeals = data.filter(row => filteredIdsSet.has(row[idIndex].toString()));

  return [headers, ...filteredDeals];
}


// Consolida dados da Pipedrive e RD Station
function consolidateData(filteredPipedriveDeals, filteredRdStationLeads, filteredIdsEmails) {
  const idToEmailMap = new Map(filteredIdsEmails.map(item => [item.id.toString(), item.email]));
  const emailToIdMap = new Map(filteredIdsEmails.map(item => [item.email, item.id.toString()]));

  const rdStationMap = new Map();
  const rdStationHeaders = filteredRdStationLeads[0];
  filteredRdStationLeads.slice(1).forEach(row => {
    const email = row[rdStationHeaders.indexOf("email")].trim();
    rdStationMap.set(email, row);
  });

  const pipedriveMap = new Map();
  const pipedriveHeaders = filteredPipedriveDeals[0];
  filteredPipedriveDeals.slice(1).forEach(row => {
    const id = row[pipedriveHeaders.indexOf("id")].toString();
    pipedriveMap.set(id, row);
  });

  const consolidatedData = [];
  const consolidatedHeaders = [...pipedriveHeaders, ...rdStationHeaders];
  consolidatedData.push(consolidatedHeaders);

  filteredIdsEmails.forEach(item => {
    const id = item.id.toString();
    const email = item.email;

    const pipedriveRow = pipedriveMap.get(id);
    const rdStationRow = rdStationMap.get(email);

    if (pipedriveRow && rdStationRow) {
      consolidatedData.push([...pipedriveRow, ...rdStationRow]);
    }
  });

  return consolidatedData;
}


// Insere dados consolidados na planilha de destino
function insertNewData(consolidatedData, sheetDealsLeadsPipeRd) {
  const lastRow = sheetDealsLeadsPipeRd.getLastRow();
  const startRow = lastRow === 1 ? 2 : lastRow + 1;
  const dataWithoutHeaders = consolidatedData.slice(1);

  if (dataWithoutHeaders.length > 0) {
    // Obtém a data atual
    const currentDate = new Date();

    // Adiciona a data atual em cada linha de dados
    const dataWithDate = dataWithoutHeaders.map(row => {
      return [...row, currentDate];
    });

    // Insere os dados na planilha, incluindo a data na última coluna
    sheetDealsLeadsPipeRd.getRange(startRow, 1, dataWithDate.length, dataWithDate[0].length).setValues(dataWithDate);
  } else {
    console.log("Nenhum dado para inserir.");
  }
}


// Função para marcar o campo "checked" como true
function markCheckedPipedriveIdEmails(sheetPipedriveIdEmails) {
    // Obtém os cabeçalhos da planilha
  const headers = sheetPipedriveIdEmails.getDataRange().getValues()[0];

  // Encontra o índice da coluna "checked"
  const checkedIndex = headers.indexOf("checked");

  // Verifica se a coluna "checked" existe
  if (checkedIndex === -1) {
    throw new Error("Coluna 'checked' não encontrada na planilha.");
  }

  // Obtém o número da última linha com dados
  const lastRow = sheetPipedriveIdEmails.getLastRow();

  // Se não houver dados além do cabeçalho, retorna
  if (lastRow < 2) {
    console.log("Nenhum dado para marcar como 'checked'.");
    return;
  }

  // Obtém todos os valores da coluna "checked" (excluindo o cabeçalho)
  const checkedValues = sheetPipedriveIdEmails.getRange(2, checkedIndex + 1, lastRow - 1, 1).getValues();

  // Itera sobre os valores e marca como true apenas os que não são true
  let count = 0; // Contador para registrar quantos campos foram atualizados
  for (let i = 0; i < checkedValues.length; i++) {
    if (checkedValues[i][0] !== true) { // Verifica se o valor não é true
      checkedValues[i][0] = true; // Atualiza o valor para true
      count++; // Incrementa o contador
    }
  }

  // Se pelo menos um campo foi atualizado, escreve os valores de volta na planilha
  if (count > 0) {
    sheetPipedriveIdEmails.getRange(2, checkedIndex + 1, checkedValues.length, 1).setValues(checkedValues);
    console.log(`Campo 'checked' marcado como true para ${count} linhas.`);
  } else {
    console.log("Nenhum campo 'checked' diferente de true encontrado.");
  }
}

// Atualiza os status dos deals na planilha final
function updateStatus(sheetPipedriveDeals, sheetDealsLeadsPipeRd) {
  const pipedriveData = extractPipedriveDealsData(sheetPipedriveDeals);
  const dealsLeadsData = extractDealsLeadsData(sheetDealsLeadsPipeRd);

  const filteredPipedriveData = filterPipedriveDeals(pipedriveData);
  updateDealsLeadsStatus(sheetDealsLeadsPipeRd, dealsLeadsData, filteredPipedriveData);
  markUpdatedInPipedriveDeals(sheetPipedriveDeals, filteredPipedriveData);

  // Extrai os dados da planilha Pipedrive Deals
  function extractPipedriveDealsData(sheetPipedriveDeals) {
    const { headers, data } = getHeadersAndData(sheetPipedriveDeals);
    const idIndex = headers.indexOf("id");
    const statusIndex = headers.indexOf("status");
    const closeTimeIndex = headers.indexOf("close_time");
    const updatedIndex = headers.indexOf("updated");

    return data.map((row, index) => ({
      rowIndex: index + 2, // +2 porque a primeira linha é o cabeçalho e o índice começa em 0
      id: row[idIndex].toString(),
      status: row[statusIndex],
      close_time: row[closeTimeIndex],
      updated: row[updatedIndex]
    }));
  }

  // Extrai os dados da planilha Deals Leads PipeRd
  function extractDealsLeadsData(sheetDealsLeadsPipeRd) {
    const { headers, data } = getHeadersAndData(sheetDealsLeadsPipeRd);
    
    const idPipedriveIndex = headers.indexOf("id_pipedrive");
    const statusPipedriveIndex = headers.indexOf("status_pipedrive");

    return data.map((row, index) => ({
      rowIndex: index + 2, // +2 porque a primeira linha é o cabeçalho e o índice começa em 0
      id_pipedrive: row[idPipedriveIndex].toString(),
      status_pipedrive: row[statusPipedriveIndex]
    }));
  }

  // Filtra os dados da Pipedrive Deals com base nas condições
  function filterPipedriveDeals(pipedriveData) {
    return pipedriveData.filter(row => row.close_time !== "" && row.updated !== true);
  }

  // Atualiza o campo status_pipedrive na planilha Deals Leads PipeRd
  function updateDealsLeadsStatus(sheetDealsLeadsPipeRd, dealsLeadsData, filteredPipedriveData) {
    const [headers] = sheetDealsLeadsPipeRd.getDataRange().getValues();
    const statusPipedriveIndex = headers.indexOf("status_pipedrive");

    const pipedriveStatusMap = new Map(filteredPipedriveData.map(row => [row.id, row.status]));

    dealsLeadsData.forEach(row => {
      const statusPipedrive = pipedriveStatusMap.get(row.id_pipedrive);
      if (statusPipedrive && statusPipedrive !== row.status_pipedrive) {
        sheetDealsLeadsPipeRd.getRange(row.rowIndex, statusPipedriveIndex + 1).setValue(statusPipedrive);
      }
    });
  }

  // Marca o campo "updated" como true na planilha Pipedrive Deals
  function markUpdatedInPipedriveDeals(sheetPipedriveDeals, filteredPipedriveData) {
    const [headers] = sheetPipedriveDeals.getDataRange().getValues();
    const updatedIndex = headers.indexOf("updated");

    if (updatedIndex === -1) {
      throw new Error("Coluna 'updated' não encontrada na planilha Pipedrive Deals.");
    }

    filteredPipedriveData.forEach(row => {
      sheetPipedriveDeals.getRange(row.rowIndex, updatedIndex + 1).setValue(true);
    });

    console.log(`Campo 'updated' marcado como true para ${filteredPipedriveData.length} linhas na planilha Pipedrive Deals.`);
  }
}

// Funções de e-mail
function errorEmail() {
  GmailApp.sendEmail(recipients, subject, body);
}

function reportEmail() {
  var subject = `${projectName} - Relatório de volumes`
  GmailApp.sendEmail(recipients, subject, body);
}


function emailNotification() {
  var recipient = "cremascopaulo@gmail.com , pivapaulo@outlook.com"; // Substitua pelo e-mail do destinatário
  var subject = "s"; // Substitua pelo assunto do e-mail
  var body = "Olá,\n\nEste é um e-mail de teste enviado usando Google Apps Script.\n\nAtenciosamente,\nSeu Nome"; // Substitua pelo corpo do e-mail

  // Envia o e-mail
}


// Função principal
function main() {
  const {sheetPipedriveDeals , sheetPipedriveIdEmails , sheetRdStationLeads , sheetDealsLeadsPipeRd} = conectToSheets();
  var idsEmails = getIdsEmails(sheetPipedriveIdEmails);

  //updateStatus(sheetPipedriveDeals , sheetDealsLeadsPipeRd);

  // Atribui os dados
  const filteredRdStationLeads = getFilteredLeadsByEmail(sheetRdStationLeads,idsEmails);

  if (filteredRdStationLeads.length === 1){
    markCheckedPipedriveIdEmails(sheetPipedriveIdEmails);
    return
  }

  // Filtra a lista de Ids e Emails apenas para os que existem na RD Station
  const filteredIdsEmails = filterIdsEmailsByExistingLeads(idsEmails, filteredRdStationLeads);

  // Atribui os dados filtrados da planilha do Pipedrive
  const filteredPipedriveDeals = getFilteredDeals(sheetPipedriveDeals, filteredIdsEmails);

  // Consolida os novos dados
  const consolidatedData = consolidateData(filteredPipedriveDeals, filteredRdStationLeads, filteredIdsEmails);

  // Insere os novos dados
  insertNewData(consolidatedData , sheetDealsLeadsPipeRd)

  // Mark checkek all e-mails extracted
  markCheckedPipedriveIdEmails(sheetPipedriveIdEmails , idsEmails);
}