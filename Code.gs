function buildAddOnCard() {
  var card = CardService.newCardBuilder();
  var section = CardService.newCardSection();
  
  section.addWidget(CardService.newTextInput()
      .setFieldName("topic")
      .setTitle("Enter topic")
      .setHint("e.g., AI, Machine Learning"));

  section.addWidget(CardService.newTextButton()
      .setText("Rank Emails")
      .setOnClickAction(CardService.newAction()
        .setFunctionName("onRankEmails")));

  card.addSection(section);
  
  return card.build();
}

function rankEmails(topic, accessToken) {
  //Logger.log("Starting to rank emails for topic: " + topic);
  
  // Fetch a reasonable number of threads (e.g., the first 100 threads)
  var threads = GmailApp.getInboxThreads(0, 5); 
  //Logger.log("Fetched " + threads.length + " threads.");

  var rankedEmails = [];

  // Process each thread
  threads.forEach(function(thread) {
    var messages = thread.getMessages();  
    //Logger.log("Processing thread with " + messages.length + " messages.");

    messages.forEach(function(message) {
      //Logger.log("Processing message with subject: " + message.getSubject());

      var snippet = message.getPlainBody().substring(0, 5); // Use the first 100 characters as a snippet
      var emailContent = message.getPlainBody();

      // Call the LLM to get a relevance score
      var relevanceScore = getLLMResponse(emailContent, topic, accessToken);
      //Logger.log(emailContent);
      rankedEmails.push({
        subject: message.getSubject(),
        snippet: snippet,  
        relevance: parseFloat(relevanceScore),  // Assuming the LLM returns a numeric score
        date: message.getDate()
      });
    });
  });

  // Sort emails by relevance score in descending order
  rankedEmails.sort(function(a, b) {
    return b.relevance - a.relevance;
  });

  //Logger.log("Finished ranking emails.");
  return rankedEmails;
}


function onRankEmails(e) {

  var mytoken="eyJraWQiOiIyMDI0MDgwMzA4NDEiLCJhbGciOiJSUzI1NiJ9.eyJpYW1faWQiOiJJQk1pZC02OTcwMDBJTVNRIiwiaWQiOiJJQk1pZC02OTcwMDBJTVNRIiwicmVhbG1pZCI6IklCTWlkIiwianRpIjoiYThiYWIyZjEtZDQ0ZC00ZjVkLWJjNjEtN2NkNjMwYzQ4MGViIiwiaWRlbnRpZmllciI6IjY5NzAwMElNU1EiLCJnaXZlbl9uYW1lIjoiVGlhbnl1YW4iLCJmYW1pbHlfbmFtZSI6IlBhbmciLCJuYW1lIjoiVGlhbnl1YW4gUGFuZyIsImVtYWlsIjoib3NjYXJwYW5ndHlAZ21haWwuY29tIiwic3ViIjoib3NjYXJwYW5ndHlAZ21haWwuY29tIiwiYXV0aG4iOnsic3ViIjoib3NjYXJwYW5ndHlAZ21haWwuY29tIiwiaWFtX2lkIjoiSUJNaWQtNjk3MDAwSU1TUSIsIm5hbWUiOiJUaWFueXVhbiBQYW5nIiwiZ2l2ZW5fbmFtZSI6IlRpYW55dWFuIiwiZmFtaWx5X25hbWUiOiJQYW5nIiwiZW1haWwiOiJvc2NhcnBhbmd0eUBnbWFpbC5jb20ifSwiYWNjb3VudCI6eyJ2YWxpZCI6dHJ1ZSwiYnNzIjoiOTU1ODM5ZTM0NWJmNDU2OWE2NTZmNGJiMDY2NjNlYzkiLCJpbXNfdXNlcl9pZCI6IjEyNjMyNjQxIiwiZnJvemVuIjp0cnVlLCJpbXMiOiIyNzQ5OTg2In0sImlhdCI6MTcyNDY4NTU1NiwiZXhwIjoxNzI0Njg5MTU2LCJpc3MiOiJodHRwczovL2lhbS5jbG91ZC5pYm0uY29tL2lkZW50aXR5IiwiZ3JhbnRfdHlwZSI6InVybjppYm06cGFyYW1zOm9hdXRoOmdyYW50LXR5cGU6YXBpa2V5Iiwic2NvcGUiOiJpYm0gb3BlbmlkIiwiY2xpZW50X2lkIjoiZGVmYXVsdCIsImFjciI6MSwiYW1yIjpbInB3ZCJdfQ.el93pPjGl5sm36GsAe8vBBqNzw0mzSj7CCwod4NQxi77d3EK4_aEgn-OMqUgzqrTUoWXVEqFGrqKbFQTaNAYRBhhWXvOFr5nfcr8wfZyuncHzvReYR4Di2h_-N25VAFt419qN7EcWC_Cv2vFJpTVQ0lONA2XfTH27kIl_hMlJybcO2njsJcV86l2RvbvqSiTCDOftWDdOYPgDIBIqhwdAmH0fpS4_mgFCICrAZNA-ZTRUU_RcrqlQYVWIAYKPCB0Eu_3VUq9NsaOOolUYyeDsq99t2a9OPxI3Q1tUU5FFyxS6Nn7hO7St1bPULFzjG2TvvFeRi5D-fz6M9nrgXoFrQ";
  // Retrieve the topic from the form input
  var topic = "";
  if (e && e.commonEventObject && e.commonEventObject.formInputs && e.commonEventObject.formInputs.topic) {
    topic = e.commonEventObject.formInputs.topic.stringValue || "";
  } else {
    topic = "default topic"; // or handle the error appropriately
  }

  // Get ranked emails from the inbox
  var rankedEmails = rankEmails(topic, mytoken);
  
  // Build the card to display ranked emails
  var card = CardService.newCardBuilder();
  var section = CardService.newCardSection();
  
  rankedEmails.forEach(function(email) {
    section.addWidget(CardService.newTextParagraph().setText(
      '<b>' + email.subject + '</b><br>' +
      '<i>' + email.snippet + '</i><br>' +
      'Relevance Score: ' + email.relevance +
      '<br> Date: ' + email.date));
  });
  
  card.addSection(section);
  
  return card.build();
}

function getLLMResponse(emailContent, topic, accessToken) {
  var url = "https://us-south.ml.cloud.ibm.com/ml/v1/text/generation?version=2023-05-29";
  
  var headers = {
    "Content-Type": "application/json",
    "Accept": "application/json",
    "Authorization": "Bearer " + accessToken
  };

  var payload = {
    "input": "Please rate the relevance of the following email content to the topic: " + topic + " on a scale of 0 out of 10. Give a number only.\n\nEmail content: \n" + emailContent+ "\n\nRelevance Score:",
    "parameters": {
      "decoding_method": "greedy",
      "max_new_tokens": 200,
      "min_new_tokens": 0,
      "stop_sequences": [],
      "repetition_penalty": 1.05
    },
    "model_id": "ibm/granite-13b-chat-v2",
    "project_id": "97c1ad78-327e-4237-96a8-307043771cd2"
  };
  
  var options = {
    "method": "post",
    "headers": headers,
    "payload": JSON.stringify(payload)
  };
  
  var response = UrlFetchApp.fetch(url, options);
  var responseData = JSON.parse(response.getContentText());
  Logger.log(payload); 
  Logger.log(responseData); 
  // Assume the response contains a relevance score or some text indicating relevance
  //return responseData.generated_text || "0";  // Adjust based on actual response structure
  //Logger.log("before: "+responseData.results[0].generated_text);
  var generatedText = responseData.results[0].generated_text || "0";  // Fallback to "0" if undefined
  //Logger.log("after: "+generatedText);
  // Get the first character safely
  var firstCharacter = generatedText.trim().charAt(0);

  // Ensure the first character is a digit
  if (!isNaN(firstCharacter)) {
    Logger.log("The first character is a digit: " + firstCharacter);
    return firstCharacter;
  } else {
    Logger.log("The first character is not a digit.");
    return generatedText.trim()|| "0";
  }
}
