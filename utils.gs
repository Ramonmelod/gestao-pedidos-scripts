function adicionarItensUltimaExped() {

  try{

    const dadosExped = sheetExped.getDataRange().getValues();
    const dadosPedItens = sheetPedItens.getDataRange().getValues();
     const dadosExpedItens = sheetExpedItens.getDataRange().getValues();

    const colExped_ID = dadosExped[0].indexOf("ID_Exp");
    const colExped_NumPedido = dadosExped[0].indexOf("Num_Pedido");

    const colPed_NumPedido = dadosPedItens[0].indexOf("Num_Pedido");
    const colPed_CodItem = dadosPedItens[0].indexOf("Cod_Item");
    const colPed_QtdPed = dadosPedItens[0].indexOf("Qtd_Ped");
    const colExpedItemCodItem = dadosExpedItens[0].indexOf("Cod_Item");
    const colExpedItemIdExp = dadosExpedItens[0].indexOf("ID_Exp");
    const colExpedItemQtdExp = dadosExpedItens[0].indexOf("Qtd_Exp");

  

    // última linha da aba exped
    const ultimaLinhaExped = dadosExped[dadosExped.length - 1];
    const idExp = ultimaLinhaExped[colExped_ID];
    const numPedido = ultimaLinhaExped[colExped_NumPedido];

    // filtra itens do pedido correspondente
    const itensPedido = dadosPedItens.slice(1).filter(linha => linha[colPed_NumPedido] === numPedido);
    
    
    // filter ID_Exp related to Num_Pedido
     const idExpRelativoAoPedido = dadosExped.slice(1).filter(linha => linha[colExped_NumPedido] === numPedido); //stores all the rows where Id_Exp matches the Num_Pedido from the last line
     const idExpFiltrados = idExpRelativoAoPedido.map(linha => linha[0]);
     //console.log(idExpFiltrados)
    

    //Add the remaining itens from the order to the table expedItens
    itensPedido.forEach(linha => {
      const codItem = linha[colPed_CodItem];
      Logger.log(`Verificando quantitativo do item de código: ${codItem}`)
      const expedItensByItemByIdExp = checkItens(codItem) //receives rows related to num_pedido and to cod_item
      Logger.log('Linhas relacionadas ao num_pedido e ao cod_item abaixo:')
      Logger.log(expedItensByItemByIdExp)
      const arrayQtd = expedItensByItemByIdExp.map(linha => linha[colExpedItemQtdExp])
      const qtdJaExped = arrayQtd.reduce((acumulador, qtd)=>{
        return acumulador + qtd
      },0)   
      const qtdAExpedir = linha[colPed_QtdPed] - qtdJaExped;
      Logger.log(`Quantidade disponível do item de código ${codItem} a expedir: ${qtdAExpedir}`)
      if(qtdAExpedir <= 0){
        console.log(`Não há mais quantitativo para expedir para o item de código: ${codItem} `)
        return //makes the the thread come out from the callback funtion of the forEach
      }

      sheetExpedItens.appendRow([idExp, codItem, qtdAExpedir]); // here the item is added in new row if there is itens remaining
    });

    Logger.log(`Adicionados ${itensPedido.length} itens para o ID_Exp ${idExp}`);

    function checkItens(codItem){
      const expedItensByItem = dadosExpedItens.slice(1).filter(row => { // the filter method iterates through the rows and if true is returned in its signature the row is selected, could be filter(true)
        const filteredRowByCodItem = row[colExpedItemCodItem] === codItem

        return filteredRowByCodItem
        }); // take all the lines retated to the item. Need to take only the related to do num_pedidos    

      const expedItensByItemByIdExp = expedItensByItem.filter(row =>{
      filteredRowByIdExp = false
      for(let i = 0; i < idExpFiltrados.length; i ++)
      {
        if(row[colExpedItemIdExp] === idExpFiltrados[i]) // here is verified if Id_Exp from the row is related to num_pedido
        {
          filteredRowByIdExp = true
      }
      }
      return filteredRowByIdExp //true of false is returned
      })
    
      return expedItensByItemByIdExp // returns the rows related to num_pedido and cod_item
      }
    
  }catch(error){
    Logger.log(`Na função adicionarItensUltimaExped ocorreu o erro: ${error}`)
    throw error
  }
}

function verificarStatusPedido(numPedido) {
  try{  
    const dadosPedidos = sheetPedidos.getDataRange().getValues();
    //console.log(dadosPedidos)
    const colNumPedido = dadosPedidos[0].indexOf("Num_Pedido"); // returns the index of the Num_Pedido column
    //console.log(colNumPedido)
    const colStatus = dadosPedidos[0].indexOf("Status");
    //console.log(colStatus)

    for (let i = 1; i < dadosPedidos.length; i++) {
      console.log(dadosPedidos[i][colNumPedido])
      if (dadosPedidos[i][colNumPedido] === numPedido) {
        const status = dadosPedidos[i][colStatus];
        return status === "Entregue";
      }
    }
    return false;
  }catch(error){
    Logger.log(`Na função verificarStatusPedido ocorreu o erro: ${error}`)
    throw error
  }
}

function mudaStatusEntregue(numPedido) {

 try{ 
  const dadosPedidos = sheetPedidos.getDataRange().getValues();

  const colNumPedido = dadosPedidos[0].indexOf("Num_Pedido");
  const colStatus = dadosPedidos[0].indexOf("Status");

  for (let i = 1; i < dadosPedidos.length; i++) {
    if (dadosPedidos[i][colNumPedido] === numPedido) {
      // atualiza a célula com "Entregue"
      sheetPedidos.getRange(i + 1, colStatus + 1).setValue("Entregue");
      Logger.log(`Status do pedido ${numPedido} alterado para Entregue.`);
      return;
    }
  }
  Logger.log(`Pedido ${numPedido} não encontrado.`);
 }catch(error){
    Logger.log(`Na função mudaStatusEntregue ocorreu o erro: ${error}`)
    throw error
 }
}

function atualizarStatusPedido(numPedido) {

 try{ 
  const target = String(numPedido).trim();
  const ss = SpreadsheetApp.openById(sheetId);


  const dadosPedidos = sheetPedidos.getDataRange().getValues();
  const dadosPedItens = sheetPedItens.getDataRange().getValues();
  const dadosExped = sheetExped.getDataRange().getValues();
  const dadosExpedItens = sheetExpedItens.getDataRange().getValues();

  // --- índices (valide nomes de coluna com exatidão) ---
  const iPed_NumPedido = dadosPedItens[0].indexOf("Num_Pedido");
  const iPed_CodItem = dadosPedItens[0].indexOf("Cod_Item");
  const iPed_QtdPed = dadosPedItens[0].indexOf("Qtd_Ped");

  const iExped_ID_Exp = dadosExped[0].indexOf("ID_Exp");
  const iExped_NumPedido = dadosExped[0].indexOf("Num_Pedido");

  const iExpedItens_ID_Exp = dadosExpedItens[0].indexOf("ID_Exp");
  const iExpedItens_CodItem = dadosExpedItens[0].indexOf("Cod_Item");
  const iExpedItens_QtdExp = dadosExpedItens[0].indexOf("Qtd_Exp");

  const iPedidos_NumPedido = dadosPedidos[0].indexOf("Num_Pedido");
  const iPedidos_Status = dadosPedidos[0].indexOf("Status");

  // helper
  const toNum = v => {
    const n = Number(v);
    return isNaN(n) ? 0 : n;
  };

  // --- 1) construir mapa dos itens pedidos para numPedido ---
  const pedMap = {}; // codItem -> qtdPed
  for (let r = 1; r < dadosPedItens.length; r++) {
    const row = dadosPedItens[r];
    if (String(row[iPed_NumPedido]).trim() === target) {
      const cod = String(row[iPed_CodItem]).trim();
      const qtd = toNum(row[iPed_QtdPed]);
      if (!cod) continue;
      pedMap[cod] = (pedMap[cod] || 0) + qtd;
    }
  }
  console.log(pedMap) //aqui é montado o objeto com relação cod_item:qtd_item

  if (Object.keys(pedMap).length === 0) {
    Logger.log(`Pedido ${target} não possui itens na tabela pedItens. Nada a fazer.`);
    return;
  }

  // --- 2) pegar todos ID_Exp da tabela exped que pertencem a esse numPedido ---
  const idExpList = [];
  for (let r = 1; r < dadosExped.length; r++) {
    const row = dadosExped[r];
    if (String(row[iExped_NumPedido]).trim() === target) {
      idExpList.push(String(row[iExped_ID_Exp]).trim());
    }
  }

  if (idExpList.length === 0) {
    Logger.log(`Nenhuma expedição (entrada em 'exped') encontrada para o pedido ${target}. Status definido como "Confirmado".`);

    // --- atualizar a planilha pedidos ---
    for (let r = 1; r < dadosPedidos.length; r++) {
      const row = dadosPedidos[r];
      if (String(row[iPedidos_NumPedido]).trim() === target) {
        const statusAtual = String(row[iPedidos_Status]).trim();
        if (statusAtual !== "Confirmado") {
          sheetPedidos.getRange(r + 1, iPedidos_Status + 1).setValue("Confirmado");
          Logger.log(`Pedido ${target}: status alterado de "${statusAtual}" para "Confirmado".`);
        } else {
          Logger.log(`Pedido ${target}: status já é "Confirmado". Nenhuma alteração.`);
        }
        return; // encerra após atualizar
      }
    }

    Logger.log(`Pedido ${target} não encontrado na aba 'pedidos' para atualizar o status.`);
    return;
  }

  // --- 3) somar Qtd_Exp em expedItens apenas para os ID_Exp acima ---
  const expedMap = {}; // codItem -> somaQtdExp
  const idSet = {};
  idExpList.forEach(id => idSet[id] = true);

  for (let r = 1; r < dadosExpedItens.length; r++) {
    const row = dadosExpedItens[r];
    const idExp = String(row[iExpedItens_ID_Exp]).trim();
    if (!idSet[idExp]) continue;
    const cod = String(row[iExpedItens_CodItem]).trim();
    const qtd = toNum(row[iExpedItens_QtdExp]);
    if (!cod) continue;
    expedMap[cod] = (expedMap[cod] || 0) + qtd;
  }

  // --- 4) comparar e decidir status ---
  let algumExpedido = false;
  let todosCompletos = true;

  for (const cod in pedMap) {
    const pedQtd = pedMap[cod] || 0;
    const expedQtd = expedMap[cod] || 0;
    if (expedQtd > 0) algumExpedido = true;
    if (expedQtd < pedQtd) todosCompletos = false;
  }

  let novoStatus = null;
  if (todosCompletos) novoStatus = "Entregue";
  else if (algumExpedido) novoStatus = "Parcial";

  if (!novoStatus) {
    Logger.log(`Pedido ${target}: nenhum item expedido ainda. Status mantido.`);
    return;
  }

  // --- 5) atualizar a planilha pedidos somente se mudou ---
  for (let r = 1; r < dadosPedidos.length; r++) {
    const row = dadosPedidos[r];
    if (String(row[iPedidos_NumPedido]).trim() === target) {
      const statusAtual = String(row[iPedidos_Status]).trim();
      if (statusAtual === novoStatus) {
        Logger.log(`Pedido ${target}: status já é "${novoStatus}". Nenhuma alteração.`);
        return;
      }
      sheetPedidos.getRange(r + 1, iPedidos_Status + 1).setValue(novoStatus);
      Logger.log(`Pedido ${target}: status alterado de "${statusAtual}" para "${novoStatus}".`);
      return;
    }
  }

  Logger.log(`Pedido ${target} não encontrado na aba 'pedidos' para atualizar o status.`);
 }catch(error){
    Logger.log(`Na função atualizarStatusPedido ocorreu o erro: ${error}`)
    throw error
 }
}

