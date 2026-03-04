/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("askai").onclick = analyzeEmail;
  }
}); 

// Office.context.mailbox.item.getSharedPropertiesAsync((result) => {
//   if (result.status === Office.AsyncResultStatus.Failed) {
//     console.error("The current folder or mailbox isn't shared.");
//     return;
//   }
//   const sharedProperties = result.value;
//   console.log(`Owner: ${sharedProperties.owner}`);
//   console.log(`Permissions: ${sharedProperties.delegatePermissions} `);
// });

export async function run() {
  /**
   * Insert your Outlook code here
   */

  const item = Office.context.mailbox.item;
  let insertAt = document.getElementById("item-subject");
  let label = document.createElement("b").appendChild(document.createTextNode("Subject: "));
  insertAt.appendChild(label);
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createTextNode(item.subject));
  insertAt.appendChild(document.createElement("br"));
}

export async function analyzeEmail() {
  // Get email content from Outlook
  const item = Office.context.mailbox.item;

  const email_body = await new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (bodyResult) => {
      if (bodyResult.status === Office.AsyncResultStatus.Failed) {
        reject(bodyResult.error.message);
      } else {
        resolve(bodyResult.value);
      }
    });
  });

  const body_message = {
            subject: item.subject,
            body: email_body
        }

  // console.log(body_message)

  // Send to your backend
  const response = await fetch('https://localhost:5000/api/analyze', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(body_message)
  });
  
  const aiResponse = await response.json();

  let insertAt = document.getElementById("item-subject");

  // Category & Subcategory
  let category = document.createElement("p");
  category.innerHTML = `<b>Category:</b> ${aiResponse.category} — ${aiResponse.subcategory}`;
  insertAt.appendChild(category);

  // Summary
  let summary = document.createElement("p");
  summary.innerHTML = `<b>Summary:</b> ${aiResponse.summary}`;
  insertAt.appendChild(summary);

  // Intent
  let intent = document.createElement("p");
  intent.innerHTML = `<b>Intent:</b> ${aiResponse.intent}`;
  insertAt.appendChild(intent);

  // Suggested Action
  let action = document.createElement("p");
  action.innerHTML = `<b>Suggested Action:</b> ${aiResponse.copilot_action}`;
  insertAt.appendChild(action);

  // Draft Reply
  let draftLabel = document.createElement("b");
  draftLabel.textContent = "Draft Reply:";
  insertAt.appendChild(draftLabel);

  let draft = document.createElement("p");
  draft.style.whiteSpace = "pre-wrap"; // preserves line breaks in the draft
  draft.textContent = aiResponse.draft_reply;
  insertAt.appendChild(draft);
}
