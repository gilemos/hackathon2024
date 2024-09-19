/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

const { AzureOpenAI } = require("openai");

const endpoint = "ENDPOINT";
const apiKey = "KEY"; // DevSkim: ignore DS173237
const apiVersion = "2023-03-15-preview";
const deployment = "giogpt"; //This must match your deployment name.

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Assign event handlers and other initialization logic.
    document.getElementById("sendQuestionButton").onclick = () => tryCatch(answerQuestion);
    document.getElementById("sideload-msg").style.display = "none";
  }
});

async function insertParagraph() {
  await Word.run(async (context) => {

    const docBody = context.document.body;
    docBody.insertParagraph("Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.",
                            Word.InsertLocation.start);

    await context.sync();

  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
      await callback();
  } catch (error) {
      // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
      console.error(error);
  }
}

async function answerQuestion() {
  await Word.run(async (context) => {

    // Getting the text from the body
    var documentBody = context.document.body;
    context.load(documentBody);

    // Getting the question from the input field
    var questionInput = document.getElementById('questionInput').value;

    // Posting the question into chat area
    var newUserBubble = document.createElement("div");
    newUserBubble.innerText = questionInput;
    newUserBubble.classList.add("right");
    newUserBubble.classList.add("bubble");
    document.getElementById('chatArea').appendChild(newUserBubble);  

    return context.sync()
    .then(async function(){
        const client = new AzureOpenAI({ endpoint: endpoint, apiKey:apiKey, apiVersion:apiVersion, deployment:deployment, dangerouslyAllowBrowser: true});
        const result = await client.chat.completions.create({
          messages: [
          { role: "system", content: "You are a helpful assistant. This story pertains to a DnD character." },
          { role: "user", content: "Hollix has purple hair. What is Hollix's hair color." },
          { role: "assistant", content: "The character Hollix is described as having purple hair" },
          { role: "user", content: documentBody.text + ". " + questionInput },
          ],
          model: "",
        });

        // Creating the response bubble
        for (const choice of result.choices) {
          console.log(choice.message);

          var newResponseBubble = document.createElement("div");
          newResponseBubble.innerText = choice.message.content;
          newResponseBubble.classList.add("left");
          newResponseBubble.classList.add("bubble");
          document.getElementById('chatArea').appendChild(newResponseBubble); 
        }

        await context.sync();
    })
  });
}

async function getSummary() {
  await Word.run(async (context) => {

    var documentBody = context.document.body;
    context.load(documentBody);
    return context.sync()
    .then(async function(){
        const client = new AzureOpenAI({ endpoint: endpoint, apiKey:apiKey, apiVersion:apiVersion, deployment:deployment, dangerouslyAllowBrowser: true});
        const result = await client.chat.completions.create({
          messages: [
          { role: "system", content: "You are a helpful assistant. This story pertains to a DnD character." },
          { role: "user", content: "Hollix has purple hair. What is Hollix's hair color." },
          { role: "assistant", content: "The character Hollix is described as having purple hair" },
          { role: "user", content: documentBody.text + ". Where was she raised?" },
          ],
          model: "",
        });
        for (const choice of result.choices) {
          console.log(choice.message);
        }
        console.log(documentBody.text);
    })
  });
}
