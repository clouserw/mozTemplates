/*

To use this code:

1) Open your Google document

2) Click Tools -> Script editor

3) Replace the pre-filled code with:

  function onOpen(ev) {
      return MozTemplates.onOpen(ev);
  }

4) Click the '+' next to Libraries in the menu on the left and paste this
Script ID:

  1Sq4Q1P-thUtNisgq_HOlMdABnQpLTM309gjXVCE_-yiiMp5Rv1jHTcKQ

5) Choose the latest version (highest number).  I wouldn't suggest using HEAD.

6) Click Add.  Save the script and close the window.  Refresh your original
document.

7) The first time you choose a menu item you will probably need to grant
permissions to the script to modify the document.

*/

function onOpen(ev) {
    DocumentApp.getUi().createMenu('MozTemplates').addItem('Insert an agenda for today', 'MozTemplates.insertToday').addToUi();
}

function insertToday(ev) {

  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  var hrulerange = body.findElement(DocumentApp.ElementType.HORIZONTAL_RULE);

  if (hrulerange) {
    var actionItems = [
      ['Action Item(s)', 'Due Date', 'Owner', 'Status Notes'],
      ['','','',''],
      ['','','','']
    ];
    var previousActionItems = [
      ['Action Item(s) from last time', 'Due Date', 'Owner', 'Status Notes'],
      ['','','',''],
      ['','','','']
    ];
    var meetingNotes = [
      ['Topic', 'Notes'],
      ['',''],
      ['',''],
      ['',''],
      ['','']
    ];
    var  time = new Date();
    time = Utilities.formatDate(time, "PST", "MMMMM dd, YYYY");

    var headerStyle = {};
    headerStyle[DocumentApp.Attribute.BOLD] = true;
    headerStyle[DocumentApp.Attribute.FONT_SIZE] = '10';
    headerStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#fce5cd';

    var hruleparent = hrulerange.getElement().getParent();
    var hruleindex = body.getChildIndex(hruleparent);

    // We'll try to use this as an anchor to pull data from the previous meeting
    _attendees = body.findText('^Attendees:.*$');
    if (_attendees) {
      _attendees_text = _attendees.getElement().asText().getText();
    } else {
      _attendees_text = "Attendees: ";
    }

    t = body.insertTable(hruleindex + 1,meetingNotes);
    t.getRow(0).getCell(0).setAttributes(headerStyle);
    t.getRow(0).getCell(1).setAttributes(headerStyle);

    // The next sibling after the attendees list is the action items from the previous meeting!
    x = _attendees.getElement().getParent().getNextSibling();

    if (x) {
      t = body.insertTable(hruleindex + 1, x.copy().asTable());
      t.getRow(0).getCell(0).setText('Action Item(s) from last time');
      t.getRow(0).getCell(0).setAttributes(headerStyle);
    } else {
      t = body.insertTable(hruleindex + 1,previousActionItems);
      t.getRow(0).getCell(0).setAttributes(headerStyle);
      t.getRow(0).getCell(1).setAttributes(headerStyle);
      t.getRow(0).getCell(2).setAttributes(headerStyle);
      t.getRow(0).getCell(3).setAttributes(headerStyle);
      t.getRow(1).getCell(0).insertListItem(0,'').setGlyphType(DocumentApp.GlyphType.BULLET);
      t.getRow(2).getCell(0).insertListItem(0,'').setGlyphType(DocumentApp.GlyphType.BULLET);
      t.getRow(1).getCell(3).insertListItem(0,'').setGlyphType(DocumentApp.GlyphType.BULLET);
      t.getRow(2).getCell(3).insertListItem(0,'').setGlyphType(DocumentApp.GlyphType.BULLET);
    }

    t = body.insertTable(hruleindex + 1,actionItems);
    t.getRow(0).getCell(0).setAttributes(headerStyle);
    t.getRow(0).getCell(1).setAttributes(headerStyle);
    t.getRow(0).getCell(2).setAttributes(headerStyle);
    t.getRow(0).getCell(3).setAttributes(headerStyle);
    t.getRow(1).getCell(0).insertListItem(0,'').setGlyphType(DocumentApp.GlyphType.BULLET);
    t.getRow(2).getCell(0).insertListItem(0,'').setGlyphType(DocumentApp.GlyphType.BULLET);
    t.getRow(1).getCell(3).insertListItem(0,'').setGlyphType(DocumentApp.GlyphType.BULLET);
    t.getRow(2).getCell(3).insertListItem(0,'').setGlyphType(DocumentApp.GlyphType.BULLET);

    // Prefill from previous attendees
    body.insertParagraph(hruleindex + 1, _attendees_text);

    z = body.insertParagraph(hruleindex + 1,time);
    z.setHeading(DocumentApp.ParagraphHeading.HEADING2)
  }

}
