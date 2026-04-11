function onFormwhatsApp(e) {
  Logger.log("Triggered ✅");

  var name = e.values[1];
  var joiningId = e.values[2];
  var groupId = e.values[3];

  var url = "https://graph.facebook.com/v18.0/1055276884338757/messages";

  var payload = {
    messaging_product: "whatsapp",
    to: "919110163886",   
    type: "template",
    template: {
      name: "hello_world",
      language: {
        code: "en_US"
      }
    }
  };

  var options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer EAARZCztlG0y0BRCJgLOMfPAZAPEEuXyHtEDUMt0Ri9ENljC2ErjlOzBazWcYNLhRI2CO9DWSzGAXn5hCJhUQC3KsKvGdG547Pq83F5K0MbEEaOdcwKXXWWgF55yN7GVPOyJ7fJuWZAKXKGlumWrQKYzBnMhRV1zCmgM4NMj4ZCIQr6dO2dbihosAaz3r0Cd0v2U5EsMhhWsBS9BE3aC7k18jnNGrgqjWHxZA5Xc7yRzJs5xbaZCPPtXWgdH4kBSB3TXLSDgbg0sZChVGcvr7zZAbRf73"
    },
    payload: JSON.stringify(payload)
  };

  var response = UrlFetchApp.fetch(url, options);
  Logger.log(response.getContentText());
}

