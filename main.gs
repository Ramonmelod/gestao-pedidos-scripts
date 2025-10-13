const sheetId = "1vbbF1Bf52On0BQ85m_jGGWwiHUQPxJbH9K77EbaJkUM"; 
const sheetExped = SpreadsheetApp.openById(sheetId).getSheetByName("exped");
const sheetPedItens = SpreadsheetApp.openById(sheetId).getSheetByName("pedItens");
const sheetExpedItens = SpreadsheetApp.openById(sheetId).getSheetByName("expedItens");
const sheetPedidos = SpreadsheetApp.openById(sheetId).getSheetByName("pedidos");

function main(numPedido) {
  try{
  Logger.log("Pedido recebido: " + numPedido);

  const statusEntregue = verificarStatusPedido(numPedido);
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
