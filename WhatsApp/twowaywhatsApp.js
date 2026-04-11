function onWhatsAppFormSubmit(e) {
  var name = e.values[1];
  var joiningId = e.values[2];
  var groupId = e.values[3];
  var phone = e.values[4];   

  var url = "https://graph.facebook.com/v18.0/1055276884338757/messages";

  // 🔹 Message to USER
  var userPayload = {
    messaging_product: "whatsapp",
    to: "91" + phone,
    type: "template",
    template: {
      name: "form_notification",
      language: { code: "en_US" },
      components: [
        {
          type: "body",
          parameters: [
            { type: "text", text: name },
            { type: "text", text: joiningId },
            { type: "text", text: groupId }
          ]
        }
      ]
    }
  };

  // 🔹 Message to YOU
  var adminPayload = {
    messaging_product: "whatsapp",
    to: "919110163886",
    type: "template",
    template: {
      name: "form_notification",
      language: { code: "en_US" },
      components: [
        {
          type: "body",
          parameters: [
            { type: "text", text: name },
            { type: "text", text: joiningId },
            { type: "text", text: groupId }
          ]
        }
      ]
    }
  };

  var options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer EAARZCztlG0y0BRIyr4y70ZCdtRq6mYKplUt57tbsyCtOkhuLMZAaoPRjZAxTDm2IsOmej8yD5ZBrbAjhTS9mIJZCQqU8QdNJhPOZAHAfXkoFQ6kqZCBYpnul3njVLDmSPTjWx9mj48AjjbZAkcxQQEdckoDZCMWltVjd4LPnnD1mHIz6EAMFZBZBQzGjoQZBjZA4DcciT1CZBTmK11kLrchSjfdzsI3JZCGlXhPLp0BKCZAZC5Dd2sBLPU8zqx86aZBSY3ZC5rxfKmR7N1ROmxuoQAu2GcT8WPPbFvK7"
    },
  };

  // Send to USER
  options.payload = JSON.stringify(userPayload);
  UrlFetchApp.fetch(url, options);

  // Send to YOU
  options.payload = JSON.stringify(adminPayload);
  UrlFetchApp.fetch(url, options);
}
