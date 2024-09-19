# Microsoft Hackathon 2024: Storytellers Guild

This project contains a Microsoft word plug-in that adds a customized ChatGPT bot that can answer questions about your Word document. The AI model is tuned to answer DnD/Lore-related questions

Here you will see two folders:
- dndOneNote: an incomplete OneNote plugin (abandoned)
- Chronicle Keeper: the more complete Word plugin (current)

## TODO items
- Improve the AI prompts so the model gives us better solutions
- Improve the design of the send button
- Test the model with longer word documents/more complicated questions
- Whatever feature you find interesting!

## How to run the code
1. Create an Azure OpenAI resource
2. In the file `src/taskpane/taskpane.js`, add your models's endpoint and apiKey to the variables of the same name
3. Create a Microsoft Word document in the browser
4. In your Microsoft Word document, select Share and then Copy Link
6. Go to terminal and, inside the Chronicle Keeper folder, run npm run start:web -- --document {url you copied}
7. If prompted, enable developer mode
8. If prompted to add a manifest, select yes and add the manifest.xml file under the Chronicle Keeper folder
