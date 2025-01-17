/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */
let selectedChat = null;
var listChats = []

let answerLoadingContainer = document.getElementById("answer-loading-container")
let appBody =  document.getElementById("app-body")
let rfpBody =  document.getElementById("rfp-body")
let promptTextArea = document.getElementById("prompt-txt")
let answerCard = document.getElementById("answer-card")
let answerContainer = document.getElementById("answer-container")
let insertBtn = document.getElementById("insert-btn")
let answerCmdBar =  document.getElementById("answer-cmd-bar")

let questionCard =  document.getElementById("question-card")
let questionContainer =  document.getElementById("question-container")

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("back-btn").onclick = () => {
      // Reset UI to list RFP
      appBody.style.display = "flex";
      rfpBody.style.display = "none";
      selectedChat = null;
      promptTextArea.value = ""
      answerContainer.innerHTML = "I can help answer RFP technical questions. Have a try here !"
      answerCmdBar.style.display = "none"
      questionCard.style.display = "none"
    }
    // PROMPT Click
    document.getElementById("prompt-send-btn").onclick = async () => {
      let promptTextArea = document.getElementById("prompt-txt")
      let promptText = promptTextArea.value
      // display question
      questionCard.style.display = "block"
      answerCard.style.display = "none"
      questionContainer.innerHTML = promptText
      answerLoadingContainer.style.display = "block"
      answerLoadingContainer.innerHTML = "Working on a response for you..."
      promptTextArea.value = ""
      let promptPayload = {
        "input": promptText,
        "variables": [
          {
            "key": "chatId",
            "value": selectedChat.id
          },
          {
            "key": "messageType",
            "value": "0"
          }
        ]
      }
      let res = await fetch(`https://r4trfi-copilot.azurewebsites.net/chats/${selectedChat.id}/messages`, {
        method: 'POST',
        headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(promptPayload)
      })
      answerLoadingContainer.innerHTML = "Getting answer ready ..."
      const content = await res.json();
      if(content.value){
        answerContainer.innerHTML = content.value
        insertBtn.setAttribute("data-text",content.value)
        answerCard.style.display = "block"
        answerCmdBar.style.display = "flex"
      }
      answerLoadingContainer.style.display = "none"
    }

    insertBtn.onclick = async (e) => {
      let data = e.currentTarget.attributes["data-text"].value
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        // Read the range address
        range.load(["values"]);
        await context.sync();
        
        range.values = [[data]]
      });
    }
    if (!selectedChat) {
      document.getElementById("app-body").style.display = "flex";
      //document.getElementById("run").onclick = run;
      // Load list of RFP
      try {
        let res = await fetch(`https://r4trfi-copilot.azurewebsites.net/chats`)
        let payload = await res.json()
        listChats = payload;
        let listAnswerUl = document.getElementById("list-rfps")

        listAnswerUl.innerHTML = ""

        for (let chat of payload) {
          let itemDiv = document.createElement("div")
          itemDiv.className = "ms-card"
          let itemHtml =
            `<span>RFP: ${chat.title}</span>
            <span style="font-size: 10px;color:#707070;font-style:normal;"></span>
          `
          itemDiv.innerHTML = itemHtml
          itemDiv.setAttribute("data-chat-id", chat.id)
          itemDiv.onclick = async (e) => {
            // display RFP view
            let data = e.currentTarget.attributes["data-chat-id"].value
            selectedChat = listChats.filter((c) => c.id === data)[0];

            document.getElementById("app-body").style.display = "none";
            document.getElementById("rfp-body").style.display = "flex";
            // bind data
            document.getElementById("rfp-id").innerHTML = selectedChat.title
          }
          listAnswerUl.appendChild(itemDiv)
        }
      } catch (error) {
        console.error(error);
      }
    } else {
      // display RFP details UI
      document.getElementById("app-body").style.display = "none";
      document.getElementById("rfp-body").style.display = "flex";

      // Display data
      document.getElementById("rfp-id").innerHTML = selectedChat.title


    }

    await Excel.run(async (context) => {
      let onSelectionChanged = () => {
        try {
          Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load(["address", "values"]);
            await context.sync();
            let promptText = range.values[0][0]
            if (promptText !== "") {
              document.getElementById("prompt-txt").value = promptText
            }
          })
        } catch (error) {

        }
      }

      let workbook = context.workbook;
      let handler = workbook.onSelectionChanged.add(onSelectionChanged);
      //eventHandlers.push({ workbook, handler });
      await context.sync();
    });
  }
});
