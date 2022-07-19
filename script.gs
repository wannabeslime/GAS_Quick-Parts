const ThisDoc = DocumentApp.getActiveDocument();
const TemplateDoc = DocumentApp.openById('') // template ID

function onOpen(e) {
  DocumentApp.getUi()
      .createMenu('Quick Parts')
      .addItem('CodeBlock', 'code_block')
      .addSeparator()
      .addSubMenu(DocumentApp.getUi().createMenu('Submenu')
          .addItem('A', 'sub_A')
          )
      .addToUi();
}

// Footholds at using arguments.
function code_block(){
  QuickParts('code_block')
}

// template
function template(){
  QuickParts('')
}


// functions
  // パーツのオフセットを取得
function GetTemplateOffsets(term) {
  var Templates = TemplateDoc.getBody();
  var paragraphs = Templates.getParagraphs();
  var term = "{" + term + "}"

  for (var i = 0; i < paragraphs.length; i++) {
    var paragraph = paragraphs[i];
    var text = paragraph.getText();
    let parent = paragraph.getParent();
    var start = parent.getChildIndex(paragraph);
    if(text == term){
      break;
    }
  }

  for (i; i < paragraphs.length; i++) {
    var paragraph = paragraphs[i];
    var text = paragraph.getText();
    let parent = paragraph.getParent();
    var end = parent.getChildIndex(paragraph);
    if(text == '/' + term){
      break;
    }
  }
  return {start, end}
}

  // パーツのインサート
function InsertPart(Type, Element, Cursor) {
  let body = ThisDoc.getBody();
  Logger.log(Type);
  switch (Type) {
    // Table
    case DocumentApp.ElementType.TABLE:
      body.insertTable(Cursor, Element)
      break;
    
    // Paragraph
    case DocumentApp.ElementType.PARAGRAPH:
      body.insertParagraph(Cursor, Element)
      break;

    // list
    case DocumentApp.ElementType.LIST_ITEM:
      body.insertListItem(Cursor, Element)
      break;

    // image
    case DocumentApp.ElementType.INLINE_IMAGE:
      body.insertImage(Cursor, Element)

    // default
    default:
      Logger.log('default')
  }
}

  // main
function QuickParts(part) {
  // init
  let body = TemplateDoc.getBody();
  let cursor = ThisDoc.getCursor();
  let element = cursor.getElement();
  let parent = element.getParent();
  let cursoroffset = parent.getChildIndex(element);
  let offsets = GetTemplateOffsets(part);
  Logger.log(offsets['start'], offsets['end']);

  // get_templates
  for (var i = offsets['start'] + 1; i < offsets['end']; i++) {
    let Part = body.getChild(i).copy();
    InsertPart(Part.getType(), Part, cursoroffset);
    cursoroffset += 1
  }
}
