/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

import { getUserProfile } from "../helpers/sso-helper";
import { filterUserProfileInfo } from "./../helpers/documentHelper";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("getProfileButton").onclick = run;

    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
      console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("insert-paragraph").onclick = insertParagraph;
    document.getElementById("apply-style").onclick = applyStyle;
    document.getElementById("hit-api").onclick = hitApi;
  }
});

export async function run() {
  getUserProfile(writeDataToOfficeDocument);
}

function writeDataToOfficeDocument(result) {
  return Word.run(function (context) {
    let data = [];
    let userProfileInfo = filterUserProfileInfo(result);

    for (let i = 0; i < userProfileInfo.length; i++) {
      if (userProfileInfo[i] !== null) {
        data.push(userProfileInfo[i]);
      }
    }

    const documentBody = context.document.body;
    for (let i = 0; i < data.length; i++) {
      if (data[i] !== null) {
        documentBody.insertParagraph(data[i], "End");
      }
    }
    return context.sync();
  });
}

async function insertParagraph() {
  await Word.run(async (context) => {

    const docBody = context.document.body;
    docBody.insertParagraph("Lorem ipsum dolor sit amet, consectetur adipiscing elit.",
      "Start");

    await context.sync();
  })
  .catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function applyStyle() {
  await Word.run(async (context) => {

    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;

    await context.sync();
  })
  .catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function hitApi() {
  console.log('hit api called')
  await Word.run(async (context) => {
    const docBody = context.document.body;

    const response = await fetch('https://rickandmortyapi.com/api/character/1,2,3')
    const planets = await response.json()

    console.log(planets)

    planets.map(planet => docBody.insertParagraph(planet.name, 'End'))

    // const response = await fetch('https://jsonplaceholder.typicode.com/posts')
    // const posts = await response.json()

    // console.log(posts)

    // posts.map(post => docBody.insertParagraph(post.title, 'End'))

    await context.sync();
  })
  .catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function test() {
  console.log('test called')
  await Word.run(async (context) => {
    const docBody = context.document.body;
    docBody.insertParagraph("test",
      "Start");

    await context.sync();
  })
  .catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}



