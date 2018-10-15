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
  .setTitle("PW word counting"); 
  DocumentApp.getUi().showSidebar(html); // using showSidbar function to modify the Ui
}

function main(){
  var para = DocumentApp.getActiveDocument().getBody().getParagraphs();
  //remove null paragraphs
  var paragraphs = para.filter(function(paragraph){return(paragraph.getText().length > 0)})
  //remove Figure descriptions
  paragraphs = paragraphs.filter(function(paragraph){return(paragraph.getAttributes().ITALIC===null)})
  //remove Bibliography
  var obj = {};
  var paraList = []; 
  paragraphs.forEach(function(paragraph){
    var attribute = paragraph.getAttributes();
    if (String(attribute.HEADING)!=='Normal'){
      obj = {};
      obj[paragraph.getText()] = [];
      paraList.push(obj)
    } else { 
      if (paraList.length>0){
        var latestObj = paraList[paraList.length-1];
        latestObj[Object.keys(latestObj)[0]].push(paragraph.getText())
      }
    }
  })
  //remove abstract, acknowledgements, appendix, bibliography
  paraList = paraList.filter(function(paraObj){
    var heading = Object.keys(paraObj)[0].trim().toLowerCase()
    heading = heading.replace(/;/g, "");
    return !(['abstract','acknowledgement','acknowledgements','appendix','bibliography'].indexOf(heading) >= 0)
  })
  paraList.forEach(function(paraObj){
    var heading = Object.keys(paraObj)[0]
    paraObj[heading] = wordCount(paraObj[heading])
  })
  return paraList
};

function wordCount(paragraphs){
  var paraCount = 0;
  for (var i=0; i<paragraphs.length; i++){
    var paragraph = paragraphs[i];
    //getting list of words
    words = paragraph.match(/\S+/g);
    if (words !== null) {
    //removing (Fig 1)
      words.forEach(function(word){
        if ((word.toLowerCase()==="(fig"||word.toLowerCase() === "(figure")&&!isNaN(words[words.indexOf(word)+1][0])){
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
  body.appendParagraph("Bibliography").setHeading(DocumentApp.ParagraphHeading.HEADING1);
  var footnotes = DocumentApp.getActiveDocument().getFootnotes(); // return a list of footnotes containing class paragraph
  var footnoteList = footnotes.map(function(footnote){return footnote.getFootnoteContents().getParagraphs()[0].getText().trim()})
  footnoteList.sort();
  footnoteList.forEach(function(footnoteText){
    body.appendParagraph(String(footnoteList.indexOf(footnoteText)+1)+") "+footnoteText)
  });
}
