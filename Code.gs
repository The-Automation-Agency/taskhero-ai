// Replace 'YOUR_OPENAI_API_KEY' with your actual OpenAI API key
const OPENAI_API_URL = 'https://api.openai.com/v1/chat/completions';

/**
 * Adds a custom menu to the active document, whether it's a Sheet or a Doc.
 */
function onOpen(e) {
  var ui;
  try {
    ui = SpreadsheetApp.getUi();
    ui.createMenu('GPT-4 Add-on')
      .addItem('Open Sidebar', 'showSidebarSheets')
      .addToUi();
  } catch (e) {
    try {
      ui = DocumentApp.getUi();
      ui.createMenu('GPT-4 Add-on')
        .addItem('Open Sidebar', 'showSidebarDocs')
        .addToUi();
    } catch (e) {
      Logger.log("Not a Google Sheets or Google Docs environment.");
    }
  }
}

/**
 * Shows the sidebar for Google Sheets.
 */
function showSidebarSheets() {
  const html = HtmlService.createHtmlOutputFromFile('Page')
      .setTitle('TaskHero')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Shows the sidebar for Google Docs.
 */
function showSidebarDocs() {
  const html = HtmlService.createHtmlOutputFromFile('Page')
      .setTitle('TaskHero - Your Friendly AI Docs Assistant')
      .setWidth(300);
  DocumentApp.getUi().showSidebar(html);
}

// Example function to insert text into a Google Doc
function insertTextIntoDoc(text) {
  const body = DocumentApp.getActiveDocument().getBody();
  body.appendParagraph(text);
}

// Example function to insert text into a Google Sheet
function insertTextIntoSheet(text) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const cell = sheet.getActiveCell();
  cell.setValue(text);
}
/**
 * Calls the OpenAI API with the provided user prompt and system prompt for content generation for Google Docs , incorporating additional parameters for content writing.
 * @param {string} contentType The type of content requested.
 * @param {string} subheadings Subheadings to include in the content.
 * @param {string} writingTone The desired tone of the writing.
 * @param {string} description A brief description of the content desired.
 * @param {string} keywords Keywords to include in the content.
 * @return {Object} The response from the OpenAI API.
 */
// Corrected and completed function
function callOpenAI(contentType, subheadings, writingTone, description, keywords, posts) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const OPENAI_API_KEY = scriptProperties.getProperty('OPENAI_API_KEY');
  var userEmail = Session.getActiveUser().getEmail();
  var userData = JSON.parse(scriptProperties.getProperty(userEmail) || '{"count": 0, "timestamp": 0}');
  var currentTime = new Date().getTime();
  var twelveHours = 12 * 60 * 60 * 1000; // 12 hours in milliseconds

  // Reset the count if more than 12 hours have passed
  if (currentTime - userData.timestamp > twelveHours) {
    userData = {"count": 0, "timestamp": currentTime};
  }

  // Check and enforce the rate limit
  if (userData.count >= 10) {
    throw new Error("Rate limit exceeded. Please try again later.");
  } else {
    userData.count++;
    scriptProperties.setProperty(userEmail, JSON.stringify(userData));

    let systemPrompt, userPrompt;

    if (contentType === "General") {
      systemPrompt = `You are a very helpful professional creative copywrite expert assistant. THE USER WILL PROVIDE THE DESCRIPTION FOR THE COPY YOU NEED TO WRITE FOR THE USER. YOU ALWAYS RESPOND IN MARKDOWN LANGUAGE AND ONLY OUTPUT THE CONTENT YOU GENERATE ONLY.`;
      userPrompt = description;
    } else if (["Facebook Post", "Linkedin Post", "TikTok Post"].includes(contentType)) {
      systemPrompt = `You are a very helpful creative copywrite expert assistant tasked with generating ${contentType}. PRODUCE THE FOLLOWING AMOUNT OF POSTS: ${posts}. YOU WILL BE PROVIDED THE TONE, KEYWORDS, AND DESCRIPTION. YOU MUST ALWAYS FOLLOW THE TONE GIVEN AND ALWAYS WORK FROM THE DESCRIPTION AND INCLUDE THE KEYWORDS GIVEN IN THE CONTENT YOU GENERATE. YOU ALWAYS RESPOND IN MARKDOWN AND ONLY OUTPUT THE CONTENT ONLY.`;
      userPrompt = `Tone: ${writingTone}. Description: ${description}. Keywords: ${keywords}.`;
    } else {
      systemPrompt = `You are a very helpful creative copywrite expert assistant tasked with generating ${contentType}. Use the following amount of subheadings: ${subheadings}. YOU WILL BE PROVIDED THE TONE, KEYWORDS, AND DESCRIPTION. YOU MUST ALWAYS FOLLOW THE TONE GIVEN AND ALWAYS WORK FROM THE DESCRIPTION AND INCLUDE THE KEYWORDS GIVEN IN THE CONTENT YOU GENERATE. YOU ALWAYS RESPOND IN MARKDOWN AND ONLY OUTPUT THE CONTENT ONLY.`;
      userPrompt = `Tone: ${writingTone}. Description: ${description}. Keywords: ${keywords}.`;
    }

    const payload = JSON.stringify({
      model: "gpt-4-turbo-preview",
      messages: [
        {"role": "system", "content": systemPrompt},
        {"role": "user", "content": userPrompt}
      ]
    });

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: payload,
      headers: {'Authorization': `Bearer ${OPENAI_API_KEY}`},
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(OPENAI_API_URL, options);
    Logger.log(response.getContentText()); // Log the API response for debugging
    const result = JSON.parse(response.getContentText());
    return result.choices[0].message.content; // Return only the assistant's response
  }
}
