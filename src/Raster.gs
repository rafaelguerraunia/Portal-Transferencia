function realizarRequestRaster() {
  var url = "https://integra.rastergr.com.br:8443/datasnap/rest/TWebService/getPosicoes";
  
  // 1. O Objeto JSON EXATAMENTE como a documentação pede
  var payload = {
    "Ambiente": "Producao", 
    "Login": "61102778000",
    "Senha": "Ffz6110@",
    "TipoRetorno": "JSON",
    "TipoConsulta": "Proximas",
    "CodUltPosicao": 0 // Se der erro interno, mude para "0" (com aspas)
  };
  
  // 2. A montagem do Request com o truque para o servidor DataSnap
  var options = {
    "method": "post", // Enviamos como POST para conseguir embutir o JSON no corpo da requisição
    "contentType": "application/json",
    "payload": JSON.stringify(payload), // Transforma o objeto em texto JSON
    "headers": {
      "Accept": "application/json",
      "X-HTTP-Method-Override": "GET" // Engana o servidor Delphi: viaja como POST, mas é processado como GET lá dentro
    },
    "muteHttpExceptions": true // OBRIGATÓRIO: impede que o script trave se o servidor devolver erro 500, permitindo ler a mensagem real
  };

  Logger.log("=== INICIANDO REQUEST ===");
  Logger.log("Enviando JSON: " + JSON.stringify(payload));

  try {
    // 3. O disparo do Request e captura da resposta
    var response = UrlFetchApp.fetch(url, options);
    
    var statusCode = response.getResponseCode();
    var content = response.getContentText();
    
    Logger.log("Status Code : " + statusCode);
    Logger.log("Body/Retorno: " + content);
    
  } catch (error) {
    Logger.log("Falha crítica ao montar o Request: " + error.message);
  }
  
  Logger.log("=== FIM DO REQUEST ===");
}