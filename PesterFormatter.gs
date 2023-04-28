// Change these to your characters' nicknames. You can add any number of elements to the main array or any of the sub-arrays.
// In my case, I added the character's abbreviated chumhandle and real name.
const characters = [
  ["RR", "Alex"],
  ["AA", "Eric"],
  ["PA", "Kate"],
  ["AH", "Sara"],
  ["CH", "Abby"],
  ["TM", "Liam"]
]

const ui = DocumentApp.getUi();

function onOpen (e) {
  ui
    .createMenu("PesterFormatter")
    .addItem("Format everyone", "formatEveryone")
    .addItem("Format character", "formatCharacter")
    .addToUi();
}

function getDocumentText () {
  return DocumentApp.getActiveDocument().getBody().editAsText();
}


function getCharacterAttributes (character) {
  let text = getDocumentText();

  return text.findText(character).getElement().asText().getAttributes();
}

function findNickname (nickname) {
  return getDocumentText().findText(nickname).getElement().asText();
}

function findCharacterIndex (str) {
  for (char in characters) {
    for (nickname of characters[char]) {
      if (nickname == str) return char;
    }
  }

  return null;
}

// Formats all instances of a certain nickname
function formatNickname (nickname, attributes, color) {
  let text = getDocumentText();

  nickname = nickname + ":"; // Dialogue is in a 'chat log' format, so it's displayed as "NICKNAME: Blah blah blah"

  let foundRange = text.findText(nickname);
  let foundText;

  while (foundRange != null) {
    foundText = foundRange.getElement().asText();
    foundText.setAttributes(attributes);
    foundText.setForegroundColor(color); // setAttributes is really unpredictable with color for some reason, so we set it separately

    foundRange = text.findText(nickname, foundRange); // Finds the next range that begins with the nickname, starting from the current range
  }
}

// Formats all instances of a certain character
function formatCharacter (character = null) {
  // Prompts the user for a nickname if character parameter doesn't exist
  if (character == null) {
    let input = ui.prompt("Character nickname: ", "Character nickname", ui.ButtonSet.OK_CANCEL);
    let selectedNickname = input.getResponseText().toString();

    character = findCharacterIndex(selectedNickname); // Finds the character with the corresponding nickname in the characters array

    if (character == null) { // If character is still null, the user asked for a nickname that doesn't exist
      ui.alert("ERROR: Character with nickname \"" + selectedNickname + "\" not found!");
      return;
    }
  }

  let formatTemplate = findNickname(characters[character][0]); // It doesn't matter which element of the character array we use, since they all have the same formatting.

  for (nickname of characters[character]) {
    formatNickname (nickname, formatTemplate.getAttributes(), formatTemplate.getForegroundColor());
  }
}

// Formats every character
function formatEveryone () {
  for (char in characters) {
    formatCharacter (char);
  }
}
