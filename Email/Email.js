
function onEmailFormSubmit(e) {
  var responses = e.values;

  var name = responses[1];   
  var email = responses[2];  

  MailApp.sendEmail({
    to: email,
    subject: "Form Submission Received",
    body: "Hello " + name + ",\n\nThank you for submitting the form!"
  });
}