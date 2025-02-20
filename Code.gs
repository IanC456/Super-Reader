function onOpen() {
  var ui = DocumentApp.getUi();
  ui.createMenu('Bionic Reading')
    .addItem('Apply Bionic Reading', 'bionicReadify')
    .addToUi();
}

function bionicReadify() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var paragraphs = body.getParagraphs();

  paragraphs.forEach(paragraph => {
    var textElement = paragraph.editAsText();
    var text = textElement.getText();
    var words = text.split(/(\s+|\W+)/);
    var index = 0;

    textElement.setBold(false); // Reset all formatting
    textElement.setForegroundColor("#000000"); // Default black text
    textElement.setFontFamily("Georgia"); // Apply Georgia font

    words.forEach(word => {
      if (word.match(/\w{2,}/)) { // Ensure word has at least 2 characters
        var boldLength = Math.ceil(word.length * 0.4);
        textElement.setBold(index, index + boldLength - 1, true); // Bold first 40%
        textElement.setForegroundColor(index + boldLength, index + word.length - 1, "#888888"); // Gray for rest
      }
      index += word.length;
    });

    // Increase line spacing to 1.5
    paragraph.setLineSpacing(1.5);
  });

  DocumentApp.getUi().alert("Bionic Reading applied!");
}
