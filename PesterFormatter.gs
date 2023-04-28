/*
  ===== PESTERFORMATTER =====
  By EduardoBirchal (https://github.com/EduardoBirchal)

  HOW TO USE:

  For this script to work, add a list of the character's nicknames at the beginning of your document.

  Format each nickname the way you want the corresponding character's dialogue to be formatted. For
  example, John would have "EB" and "John" as nicknames, and they would be formatted with the
  Courier New font and blue color.

  A single character can have any number of nicknames and you can put them in the same paragraph or in
  different paragraphs, whichever you like. You can write whatever you like on the paragraph,
  as long as it has all the nicknames you want.

  Do not put multiple formattings in the same paragraph!

  EXAMPLE:
    [At the beginning of the document]

    "Nicknames for my guys:
    EB - John (ectoBiologist)
    TT - Rose (tentacleTherapist)
    TG: Dave (turntechGodhead)
    bluh bluh this line doesnt matter since there's no nicknames here
    GG - this character's name is Jade"

    [Note that the contents of the paragraphs don't matter, as long as they contain all the nicknames you
    want and have the formatting you want]

  ===========================
*/

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
  if (character == null) {
    let input = ui.prompt("Character nickname: ", "Character nickname", ui.ButtonSet.OK_CANCEL);
    let selectedNickname = input.getResponseText().toString();

    character = findCharacterIndex(selectedNickname);

    if (character == null) {
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

function formatPA () {
  formatCharacter(1);
}

function test () {
  let aa = findNickname("AA");
  formatNickname("BAB", aa.getAttributes(), aa.getForegroundColor());
}
