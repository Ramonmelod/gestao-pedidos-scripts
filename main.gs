const sheetId = "1DWUoOVmhPryGbKqOpiSDtOPTQywzwgDboSshGLLFQSA"; 
const sheetExped = SpreadsheetApp.openById(sheetId).getSheetByName("exped");
const sheetPedItens = SpreadsheetApp.openById(sheetId).getSheetByName("pedItens");
const sheetExpedItens = SpreadsheetApp.openById(sheetId).getSheetByName("expedItens");
const sheetPedidos = SpreadsheetApp.openById(sheetId).getSheetByName("pedidos");

function main(numPedido) {
  try{
  Logger.log("Pedido recebido: " + numPedido);

  const statusEntregue = verificarStatusPedido(numPedido);//alterar nome para verificarSeStatusEntregue
  Logger.log("Já entregue? " + statusEntregue);

  if (statusEntregue) { 
    // se já estiver entregue, encerra

    return;
  }

  adicionarItensUltimaExped();

  // muda o status do pedido para entregue
  mudaStatusEntregue(numPedido);
  }catch(error){
    Logger.log(error)
  }
}

function fixExpedItens(){
    const dadosPedidos = sheetPedidos.getDataRange().getValues();
    const colNumPedido = dadosPedidos[0].indexOf("Num_Pedido"); // returns the index of the Num_Pedido column
    const colStatus = dadosPedidos[0].indexOf("Status");

    for (let i = 1; i < dadosPedidos.length; i++) {
      const numPedido = dadosPedidos[i][colNumPedido]
      console.log(numPedido)
      atualizarStatusPedido(numPedido)
  
    }

  //
}
/*
function showIdExped(){
    const dadosPedidos = sheetPedidos.getDataRange().getValues();
    const colNumPedido = dadosPedidos[0].indexOf("Num_Pedido"); // returns the index of the Num_Pedido column
    const colStatus = dadosPedidos[0].indexOf("Status");

    const dadosExped = sheetExped.getDataRange().getValues();
    const colExped_NumPedido = dadosExped[0].indexOf("Num_Pedido");




    for (let i = 1; i < dadosPedidos.length; i++) {
      const numPedido = dadosPedidos[i][colNumPedido]
      //console.log(numPedido)
      
   for(let j = 1; j < dadosExped.length; j++){
    const numPedidoExp = dadosExped[j][colExped_NumPedido]
    if(numPedido===numPedidoExp){
  
    console.log(numPedidoExp)
    }
    }



      //atualizarStatusPedido(numPedido)
  
    }


}*/
