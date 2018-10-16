/* What should the add-on do after it is installed */
function onInstall() {
  onOpen();
}

function onOpen(){
  DocumentApp.getUi() // return Ui class to be modified
  .createAddonMenu()
  .addItem('PW',"showSidebar") //addItem(name, function to invoke)
  .addToUi(); 
}

function showSidebar(){
  var html = HtmlService.createTemplateFromFile('PW') //return html template
  .evaluate() //evaluating the html template to return htmloutput
  .setTitle("PW JARVIS"); 
  DocumentApp.getUi().showSidebar(html); // using showSidbar function to modify the Ui
}

function main(){
  var removeHeading = ['abstract','acknowledgement','acknowledgements','appendix','bibliography','table','content','contents'];
  var para = DocumentApp.getActiveDocument().getBody().getParagraphs();
  //remove nullparagraphs and Figure descriptions
  var paragraphs = para.filter(function(paragraph){return (paragraph.getText().trim().length>0 && paragraph.getAttributes().ITALIC===null)})
  var obj = {};
  var paraList = [];
  //creating an array of object, each object contain pair (heading, array of paragraphs word)
  paragraphs.forEach(function(paragraph){
    var attribute = paragraph.getAttributes();
    if (String(attribute.HEADING)!=='Normal'){
      obj = {};
      obj[paragraph.getText().trim()] = [];
      paraList.push(obj)
    } else { 
      if (paraList.length>0){
        var latestObj = paraList[paraList.length-1];
        latestObj[Object.keys(latestObj)[0]].push(paragraph.getText().trim())
      }
    }
  })
  //remove abstract, acknowledgements, appendix, bibliography
  paraList = paraList.filter(function(paraObj){
    //getting the first word of the heading
    var heading = Object.keys(paraObj)[0].trim().toLowerCase().match(/\S+/g)[0]
    heading = heading.replace(/;/g, ""); //removing :,; after the heading, making it easier to check n remove
    return !(removeHeading.indexOf(heading) >= 0)
  })
  //Send the array of paragraphs word to wordCount() function
  paraList.forEach(function(paraObj){
    var heading = Object.keys(paraObj)[0]
    paraObj[heading] = wordCount(paraObj[heading])
  })
  return paraList
};

function wordCount(paragraphs){
  var removeFig = ['(fig','(figure','(fig.','(appendix'];
  var paraCount = 0;
  for (var i=0; i<paragraphs.length; i++){
    var paragraph = paragraphs[i];
    //getting list of words
    words = paragraph.match(/\S+/g);
    // Removing figure discription, usually starting with Figure 2
    if (words[0].toLowerCase()=='figure' && !isNaN(words[1][0])){
      break;
    } 
    words = words.filter(function(word){return !word.match(/\(\s*[A-Z]*\s*\)*/)})
    if (words !== null) {
    //removing (Fig 1)
      words.forEach(function(word){
        //can be simplified with regex
        if (removeFig.indexOf(word.toLowerCase())>=0 && !isNaN(words[words.indexOf(word)+1][0])){
          words.splice(words.indexOf(word),2);
        };
      });
      //removing the proper nouns (consecutive uppercase words counted as 1 word)
      words = words.filter(function(word){
        var i = words.indexOf(word);
        before = word;
        after = words[i+1];
        if (after === undefined) return true
        return !(/[A-Z]/.test(before[0]) && /[A-Z]/.test(after[0]))
      });
      paraCount += words.length;
    }
  }
  return paraCount
};

function create_bibli(){
  var body = DocumentApp.getActiveDocument().getBody();
  body.appendParagraph("Bibliography").setHeading(DocumentApp.ParagraphHeading.HEADING1); //set headings for Bibliography so that it is excluded
  var footnoteList = DocumentApp.getActiveDocument().getFootnotes(); // return a list of footnotes containing class paragraph
  //trim white space at the beginning of the footnote
  footnoteList = footnoteList.map(function(footnoteItem){
    var footnoteItem = footnoteItem.getFootnoteContents().getParagraphs()[0].copy()
    var text = footnoteItem.getText().trim();
    footnoteItem.setText(text)
    return footnoteItem
  })
  //remove same footnotes
  footnoteList.forEach(function(footnoteItem){
    var i = footnoteList.indexOf(footnoteItem);
    if (i < footnoteList.length-1){
      if (footnoteItem.getText() === footnoteList[i+1].getText()) {footnoteList.splice(i,1)}
    }
  })
  //sorting by Text
  footnoteList.sort(function(a,b){
    return a.getText().localeCompare(b.getText())
  })
  var style = {}
  style[DocumentApp.Attribute.FONT_SIZE] = 11;
  style[DocumentApp.Attribute.LINE_SPACING] = 1;
  style[DocumentApp.Attribute.INDENT_FIRST_LINE] = 0;
  style[DocumentApp.Attribute.INDENT_START] = 36;
  footnoteList.forEach(function(footnotePara){
    body.appendParagraph(footnotePara.setAttributes(style));
  });
}
