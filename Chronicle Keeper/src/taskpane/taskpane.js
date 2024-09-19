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
    /*document.getElementById("insert-paragraph").onclick = () => tryCatch(insertParagraph);*/
    document.getElementById("apply-style").onclick = () => tryCatch(getSummary);
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    /*document.getElementById("run").onclick = run;*/
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

    const client = new AzureOpenAI({ endpoint: endpoint, apiKey:apiKey, apiVersion:apiVersion, deployment:deployment, dangerouslyAllowBrowser: true});
        const result = await client.chat.completions.create({
          messages: [
          { role: "system", content: "You are a helpful assistant." },
          { role: "user", content: "Does Azure OpenAI support customer managed keys?" },
          { role: "assistant", content: "Yes, customer managed keys are supported by Azure OpenAI?" },
          { role: "user", content: "Do other Azure AI services support this too?" },
          ],
          model: "",
        });

    const docBody = context.document.body;

    for (const choice of result.choices) {
      docBody.insertParagraph(choice.message.content,
        Word.InsertLocation.start);
    }

    await context.sync();
  });
}
