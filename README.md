# Word Plugin Example Using SSO

## Getting Started

Install Dependencies
`npm install`

Build it
`npm run build:dev`

Run it
`npm start`

## Configuring SSO
This requires installing azure cli which might require installing a few other tools. It took me a while to get it all installed. You can kick this off by running `npm run configure-sso`

My login fails due to my lack of admin permissions, but it does attempt to sign in and displays a Microsoft Sign In window.

## Gist of things
These buttons `<button class="ms-Button" id="insert-paragraph">Insert Paragraph</button>` in taskpane.html call functions in taskpane.js that interact with the Word JS API which perform operations on the Word document.