const TRIGGER_FUNCTION = 'registerChanges';

// Function to make the menu.
function onOpen() {
  let menu = SpreadsheetApp.getUi().createMenu('Vocab');
  menu.addItem('Authorize/Install Script', 'setupTrigger');
  menu.addItem('Download Audio', 'downloadAudio')
  menu.addToUi();
}

function triggerExists() {
  return ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction()).includes(TRIGGER_FUNCTION);
}

// Function to setup the onChange trigger.
function setupTrigger() {
  let ui = SpreadsheetApp.getUi();
  if (!triggerExists())
  {
    ScriptApp.newTrigger(TRIGGER_FUNCTION)
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onChange()
    .create();
    ui.alert("The script has been successfully setup! 🥳🎉\nYou can now use the spreadsheet.\nYou do not need to do this step again.")
  }
  else
  {
    ui.alert("The script has already been setup.")
  }
}

// This is a long function which fills in data for new rows.
function registerChanges(e) {

  const sheet = e.source.getActiveSheet();
  const range = sheet.getActiveRange();

  // This is why you can rearrange tcolumns without any problems:
  let headerValues = sheet.getRange("A1:1").getValues()[0];
  const TRADITIONAL_COL = headerValues.indexOf("traditional")+1;
  const SIMPLIFIED_COL = headerValues.indexOf("simplified")+1;
  const ENGLISH_COL = headerValues.indexOf("english")+1;
  const PINYIN_COL = headerValues.indexOf("pinyin")+1;
  const DATE_ADDED_COL = headerValues.indexOf("date added")+1;
  const LOADING_COL = headerValues.indexOf("loading")+1;
  const FIRST_ROW = 2;
  const LAST_ROW = sheet.getLastRow();

  // Query the whole spreadsheet to look for any rows that are still "loading," aka need to be filled in.
  let columnValues = sheet.getRange(FIRST_ROW, LOADING_COL, LAST_ROW).getValues();
  // I forget what the next line does...
  let searchResult = columnValues.map((elm, idx) => elm[0] == true ? idx+2 : '').filter(String);

  searchResult.forEach(function(rangeRow) {

  console.log("Update row " + rangeRow + " (" + e.changeType + ")");

  let date_added = sheet.getRange(rangeRow, DATE_ADDED_COL);
  let simplified = sheet.getRange(rangeRow, SIMPLIFIED_COL);
  let traditional = sheet.getRange(rangeRow, TRADITIONAL_COL);
  let english = sheet.getRange(rangeRow, ENGLISH_COL);
  let pinyin = sheet.getRange(rangeRow, PINYIN_COL);

  let dataRange = sheet.getRange(rangeRow, 1, 1, DATE_ADDED_COL-1);
  let isEmpty = dataRange.getValues()[0].filter(String).length == 0;
  if (rangeRow > 1 && !isEmpty)
  {
    // Fill in timestamp:
    if (date_added.getValue() === "") {
      date_added.setValue(new Date());
    }

    // Fill in translations:
    if (traditional.getValue() && !simplified.getValue()) {
      simplified.setValue(LanguageApp.translate(traditional.getValue(), 'zh-TW', 'zh-CN'));
    }
    if (simplified.getValue() && !traditional.getValue()) {
      traditional.setValue(LanguageApp.translate(simplified.getValue(), 'zh-CN', 'zh-TW'));
    }
    if (english.getValue() && !traditional.getValue()) {
      traditional.setValue(LanguageApp.translate(english.getValue(), 'en-US', 'zh-TW'));
    }
    if (english.getValue() && !simplified.getValue()) {
      simplified.setValue(LanguageApp.translate(english.getValue(), 'en-US', 'zh-CN'));
    }
    if (traditional.getValue() && !english.getValue()) {
      let englishTranslation = LanguageApp.translate(traditional.getValue(), 'zh-TW', 'en-US');
      if (englishTranslation.split().length < 2) englishTranslation = englishTranslation.toLowerCase();
      english.setValue(englishTranslation);
    }

    // Fill in pinyin:
    if (traditional.getValue() && !pinyin.getValue()) {
      pinyin.setFormula('=PINYIN('+traditional.getA1Notation()+')');
      pinyin.setValue(pinyin.getValue());
    }
  }
  });
}

// This is a very long and unelegant function to generate the shell script.
function downloadAudio() {
  const sheet = SpreadsheetApp.getActiveSheet();

  let headerValues = sheet.getRange("A1:1").getValues()[0];
  const TRADITIONAL_COL = headerValues.indexOf("traditional")+1;
  const ENGLISH_COL = headerValues.indexOf("english")+1;

  let selection = sheet.getSelection().getActiveRange();

  // command to make a temporary folder
  let commands = "cd /tmp; mkdir mandarin-audio; cd ./mandarin-audio;" + "\n";
  for (let rangeRow = selection.getRow(); rangeRow <= selection.getLastRow(); rangeRow++)
  {
      let traditional = sheet.getRange(rangeRow, TRADITIONAL_COL).getValue();
      let english = sheet.getRange(rangeRow, ENGLISH_COL).getValue();
      // curl command to download the English audio
      commands += "curl -L 'https://translate.google.com/translate_tts?ie=UTF-8&tl=en-US&client=tw-ob&q="
      + encodeURIComponent(english).replace(/'/g, "%27") + "' > " + (rangeRow*3-2) + ".mp3" + "\n";
      // curl command to download the Chinese audio
      commands += "curl -L 'https://translate.google.com/translate_tts?ie=UTF-8&tl=zh-TW&client=tw-ob&q="
      + encodeURIComponent(traditional).replace(/'/g, "%27") + "' > " + (rangeRow*3-1) + ".mp3" + "\n";
      // sox command to generate silence
      commands += "sox -n -r 24000 -c 1 -b 16 " + (rangeRow*3) + ".wav trim 0.0 1" + "\n";
  }
  // sox command to convert mp3s to wav
  commands += 'for i in *.mp3; do sox "$i" -r 24000 -c 1 -b 16 "$(basename -s .mp3 "$i").wav"; done' + "\n"
  // sox command to merge the audio
  commands += 'sox $(ls *.wav | sort -n) ~/Downloads/Audio.wav' + "\n"
  // Move temporary files to trash
  commands += "cd ..; mv ./mandarin-audio ~/.Trash";

  let output = "<link rel=\"stylesheet\" href=\"https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css\" />"
  + "<p>These command are meant to be run in Terminal on a Mac! The selected rows will be exported to a single audio file.<br />Note: Requires <a href=\"https://formulae.brew.sh/formula/sox\">sox t installed</a>.<br /><strong>ALWAYS BE CAREFUL what you paste into Terminal.</strong></p>"
  + "<textarea readonly=\"true\" style=\"font-family:monospace; width:100%; height:150px; font-size: 9pt; overflow: auto; white-space: pre-wrap;\">" + commands + "</textarea>"
  + "<input type=\"button\" value=\"Copy to clipboard\" onClick=\"this.select(); document.execCommand('copy');\" />"
  let html = HtmlService.createHtmlOutput(output).setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, "Download selection as audio file");
}