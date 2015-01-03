/**
 * The onOpen function runs automatically when the Google Docs document is
 * opened. Use it to add custom menus to Google Docs that allow the user to run
 * custom scripts. For more information, please consult the following two
 * resources.
 *
 * Extending Google Docs developer guide:
 *     https://developers.google.com/apps-script/guides/docs
 *
 * Document service reference documentation:
 *     https://developers.google.com/apps-script/reference/document/
 */
function onOpen() {
  DocumentApp.getUi().createMenu("WriteWell").addItem("Open Sidebar", "writeWell2")
  .addItem("Open a Document with WriteWell", "documentSelect")
  .addItem("Generate multi-doc comments", "multiDocumentSelect")
  .addToUi();
  
  //UiApp.createApplication();
}

/**
* Unknown, possibly defunct:
*/
function writeWell(){
  var uiInstance = UiApp.createApplication()
  .setTitle('WriteWell')
  .setWidth(325);
  var panel = uiInstance.createVerticalPanel();
  uiInstance.add(panel);
  var handler = uiInstance.createServerHandler('click').addCallbackElement(panel);
  var box = uiInstance.createListBox().setId("select").setName("mySelector");//.setStyleAttributes({'padding':'25px','background':'#ffffee'}).setHeight('100%').setWidth('100%');;
  var images = getImages();
  images = cleanArray(images);
  for(var i = 0; i<images.length; i++){
    box.addItem("image"+i,i); 
    uiInstance.add(uiInstance.createLabel(images[i].getLinkUrl()));
  }
  panel.add(box);
  panel.add(uiInstance.createButton("Add Icon").addClickHandler(handler).setId("AddItemButton"));
  DocumentApp.getUi().showSidebar(uiInstance);  
  //UserProperties.deleteAllProperties();
}

/**
* Initializes the WriteWell Toolbar on the right side of the screen
* (the menu on the right that says WriteWell in cursive, the sidebar)
* Called in onOpen() above.
*/
function writeWell2(){
  var app= UiApp.createApplication()
  .setTitle('WriteWell')
  //.setWidth(250);
  .setWidth(290);
  var panel = app.createVerticalPanel().setHorizontalAlignment(UiApp.HorizontalAlignment.CENTER).setId("VPanel");//.setStyleAttribute('border', '0066ff');
  app.add(app.createScrollPanel().add(panel).setWidth("100%").setHeight("100%"));
  panel.add(app.createImage("http://www.writewellgroup.com/writewell.bmp").setWidth(290));//Sets banner
  var handler = app.createServerHandler('click').addCallbackElement(panel);
  var changeHandler = app.createServerHandler('changeCat').addCallbackElement(panel);
  var category = app.createListBox().setId("select").setName("mySelector").addChangeHandler(changeHandler);
  //.setStyleAttributes({'padding':'25px','background':'#ffffee'}).setHeight('100%').setWidth('100%');;
  category.addItem("Punctuation", 0);
  category.addItem("Usage", 1);
  category.addItem("Content", 2);
  category.addItem("Form", 3);
  
  var itemList = app.createListBox().setId("select2").setName("itemList").addChangeHandler(changeHandler).setTag("0");
  var initImages = getImageList(0);
  for(var i = 2; i<initImages.length; i+=3){
    itemList.addItem(initImages[i],i/3);
  }
  panel.add(category);
  panel.add(itemList);
  var iconAdder = app.createButton("Add Icon").addClickHandler(handler).setId("AddItemButton");
  //var itemSelectButtons = app.createHorizontalPanel();
  panel.add(iconAdder);
  var previewImage = app.createImage().setId("PreviewImage").setUrl(initImages[0]).addClickHandler(handler).setHeight(50).setWidth(50);
  panel.add(app.createLabel("Image Preview (Do Not Drag)"));
  panel.add(previewImage)
  var favorites = app.createListBox().setId("Favorites").setName("favorites").addChangeHandler(changeHandler);
  for(var i = 0; i<25; i++){
    if(favorites[i]!=null){
      favorites.addItem("Favorite #"+i,i);
    }
    else{
      favorites.addItem("Favorite #"+i, i);
    }
  }  
  //var favoritesBar = app.createHorizontalPanel();
  panel.add(favorites);
  //if(DocumentApp.getActiveDocument().getId()=="1KCpjGJUnVfhi9SkpJPfDyURfW3Xds0CidYyoat3oT6U"){
    panel.add(app.createButton("Set as Favorite").addClickHandler(handler).setId("favoriteSetter"));
  //}
  //favoritesBar.add(favorites);
  //panel.add(favoritesBar);
  //panel.add(app.createButton("Add Favorite").addClickHandler(handler).setId("favoriteGetter"));//Deprecated
  var favoriter = app.createLabel().setId("Favoriter").setText("Favorites");
  panel.add(favoriter);
  panel.add(populateGrid().setId("myGrid"));
  var miscButtons = app.createHorizontalPanel();
  miscButtons.add(app.createButton("Set Selection to Red").addClickHandler(handler).setId("redTexter"));
  miscButtons.add(app.createButton("Generate Comments").addClickHandler(handler).setId("commenter"));
  panel.add(miscButtons);
  DocumentApp.getUi().showSidebar(app);  
}

function click(eventInfo){
  var app = UiApp.getActiveApplication();
  var orig = eventInfo.parameter.source;
  if(orig == "AddItemButton" || orig == "PreviewImage"){
    var index = parseInt(eventInfo.parameter.mySelector);
    var item = parseInt(eventInfo.parameter.itemList);
    var images = getImageList(index);//imageResource.getBody().getImages();
    images = cleanArray(images);
    var cursor =  DocumentApp.getActiveDocument().getCursor();
    if(cursor!=null){
      cursor.insertInlineImage(UrlFetchApp.fetch(images[item*3]).getBlob()).setLinkUrl(images[item*3+1]).setHeight(24).setWidth(24);
    }
    else{
      DocumentApp.getUi().alert("Cursor not detected");
    }
  }
  else if(orig == "favoriteSetter"){
    var index = eventInfo.parameter.favorites;
    var cat = eventInfo.parameter.mySelector;
    var item = eventInfo.parameter.itemList;
    UserProperties.setProperty("userFavorites"+index,""+cat+item);
    //app.close();
    //DocumentApp.getUi().showSidebar(app);
    app.close();
    UiApp.createApplication();
    writeWell2();
    //var url = getImageList(parseInt(cat))[parseInt(item)*3];
    //var image = app.getElementById("gridFavoriteImage"+index).setUrl(url).setWidth(250);
    //getFavoritesBoxList(app.getElementById("Favorites")); 
    
  }
  else if(orig == "favoriteGetter"){
    var index = eventInfo.parameter.favorites;
    var userProperty = getFavorites()[index];
    var cat = parseInt(userProperty.substring(0,1));
    var item = parseInt(userProperty.substring(1));
    var images = getImageList(cat);
    DocumentApp.getActiveDocument().getCursor().insertInlineImage(UrlFetchApp.fetch(images[item*3]).getBlob()).setLinkUrl(images[item*3+1]).setHeight(24).setWidth(24);
  } else if(orig == "redTexter"){
    if(DocumentApp.getActiveDocument().getCursor()!=null){
      DocumentApp.getUi().alert("Cursor not detected");
      return;
    }
    var selection = DocumentApp.getActiveDocument().getSelection();
    if(selection == null){
      var offset = DocumentApp.getActiveDocument().getCursor().getSurroundingTextOffset();//.getSurroundingText().setForegroundColor("#ff0000");
      if(offset<1){
        offset = 1;
      }
      DocumentApp.getActiveDocument().getCursor().insertText(" ");
      DocumentApp.getActiveDocument().getCursor().getSurroundingText().setForegroundColor(offset, offset, "#ff0000").setBold(true).setFontSize(10);
    }
    else{
    selection = selection.getSelectedElements();
    for(var i = 0; i<selection.length; i++){
      var select = selection[i];
      var element = select.getElement();
      if(element.getText()== null){
        break;
      }
      var start = select.getStartOffset();
      var end = select.getEndOffsetInclusive();
      var text = element.setForegroundColor(start, end, "#ff0000").setBold(true).setFontSize(10);
      }
    }
  }
  else if(orig == "commenter"){
    var body = DocumentApp.getActiveDocument().getBody();
    var images = body.getImages();
    for(var i = 0; i<images.length; i++){
      images[i] = images[i].getLinkUrl();
    }
    var out = "";
    if(images.length == 0){
      DocumentApp.getUi().alert("Error: No comments detected");
    }
    body.appendParagraph("Automatically Generated Comments").setAlignment(DocumentApp.HorizontalAlignment);
    var absoluteTotal = 0;
    for(var i = 0; i<4; i++){//Category loop
      out+= catDictionary(i) + ":\n";
      var catImages = getImageList(i);
      var values = getImgValues(i);
      var ptTotal = 0;
      for(var x = 1; x<catImages.length; x+=3){//Image check loop
        var count = getCount(images,catImages[x]);
        if(count!=0){//Only prints if #of errors is not 0
          var pts = count*values[parseInt((x-1)/3)];
          out+= "\t" + catImages[x+1] + ": " + count + "*("+ values[(x-1)/3] + "pts each) = " + (pts) + "pts\n"; //Generates individual comment
          ptTotal+= pts;
        }
      }
      out+= "Subtotal: " + ptTotal +"pts\n\n"; 
      absoluteTotal += ptTotal;
    }
    out+= "\nTotal: " + absoluteTotal +"pts";
    body.appendParagraph(out).editAsText().setForegroundColor("#ff0000").setBold(true);//Adds text and sets the color to red.
  }
}

function changeCat(eventInfo){
  var app = UiApp.getActiveApplication();
  var index = parseInt(eventInfo.parameter.mySelector);
  var orig = eventInfo.parameter.source;
  if(orig == "select"){
    var items = app.getElementById('select2').clear();
    var images = getImageList(index);
    for(var i = 2; i<images.length; i+=3){
      items.addItem(images[i],i/3);
    }
    
    //app.getElementById("VPANEL").setHorizontalAlignment(UiApp.HorizontalAlignment.RIGHT).setId("VPanel");
   // DocumentApp.getUi().showSidebar(app);
    
   
  }
  if(orig == "select2" || orig == "select"){
    var image = app.getElementById("PreviewImage");
    var images = getImageList(index);
    var index2 = parseInt(eventInfo.parameter.itemList);
    image.setUrl(images[index2*3]).setWidth(50).setHeight(50);
  }
  if(orig == "Favorites"){
    /*var index = parseInt(eventInfo.parameter.favorites);
    var userProperty = getFavorites()[index];
    var cat = parseInt(userProperty.substring(0,1));
    var item = parseInt(userProperty.substring(1));
    var image = getImageList(cat);
    app.getElementById("FavoritePreview").setUrl(image[item]*3);*/
    //Broken will fix!
  }
  return app;
}

function cleanArray(array){
  var out = [];
  for(var i = 0, x = 0; i<array.length; i++){
    if(array[i]!=null){
      out[x] = array[i];
      x++;
    }
  }
  return out;
}

function getImages(){
  var id = "1ou_WBlXyV-yIEkZ595OTnlnz9la_km0Rfp3KKcZpyg0";
  var imageResource = DocumentApp.openById(id);
  var images = imageResource.getBody().getImages();
  return images;
}

/** Generates a list of images
  * index defines the category
  */
function getImageList(index){
  var app = UiApp.getActiveApplication();
  index = parseInt(index);
  var images = [];//Begin initialization of images
  if(index == 0){//These are the images from the punctuation category
    images[0] = "http://www.writewellgroup.com/bigicons/P_P1COMMA.GIF";//The location of the image file
    images[1] = "http://www.writewellgroup.com/p1.html#commap1";//The image file's link
    images[2] = "Comma needed: P.1: compound sentence";//The Name of the file for the drop down.
    
    images[3] = "http://www.writewellgroup.com/bigicons/P_P1COMMA_NO.GIF";  
    images[4] = "http://www.writewellgroup.com/p1.html#nocommap1";
    images[5] = "No comma needed: P.1";
    
    images[6] = "http://www.writewellgroup.com/bigicons/P_P1SEMI.GIF";
    images[7] = "http://www.writewellgroup.com/p1.html#semip1";
    images[8] = "Semicolon needed: P.1: compound sentence";
    
    images[9] = "http://www.writewellgroup.com/bigicons/P_P2A.GIF";
    images[10] = "http://www.writewellgroup.com/p1.html#semip1";
    images[11] = "Comma needed: Introductory Clause";
    
    images[12] = "http://www.writewellgroup.com/bigicons/P_P2B.GIF";
    images[13] = "http://www.writewellgroup.com/p2.html#commap2b";
    images[14] = "Comma needed: Introductory Phrase";
    
    images[15] = "http://www.writewellgroup.com/bigicons/P_P2C.GIF";
    images[16] = "http://www.writewellgroup.com/p2.html#commap2c";
    images[17] = "Comma needed: Introductory Word";
  
    images[18] = "http://www.writewellgroup.com/bigicons/P_P2_NO.GIF";
    images[19] = "http://www.writewellgroup.com/p2.html#nocommap2";
    images[20] = "No Comma Needed: P.2";
    
    images[21] = "http://www.writewellgroup.com/bigicons/P_P3.GIF";
    images[22] = "http://www.writewellgroup.com/p3.html";
    images[23] = "Comma needed: Non-Essential";
    
    images[24] = "http://www.writewellgroup.com/bigicons/P_P3_NO.GIF";
    images[25] = "http://www.writewellgroup.com/p3.html#nocommap3";
    images[26] = "No Comma Needed: P.3";
    
    images[27] = "http://www.writewellgroup.com/bigicons/P_P4.GIF";
    images[28] = "http://www.writewellgroup.com/p456.html#commap4";
    images[29] = "Comma needed: Sentence Interrupter";
    
    images[30] = "http://www.writewellgroup.com/bigicons/p_p4_no.gif";
    images[31] = "http://www.writewellgroup.com/p456.html#nocommap4";
    images[32] = "No Comma Needed: P.4";
    
    images[33] = "http://www.writewellgroup.com/bigicons/P_P5COMMA.GIF";
    images[34] = "http://www.writewellgroup.com/p456.html#commap5";
    images[35] = "Comma: Series of Expressions";
    
    images[36] = "http://www.writewellgroup.com/bigicons/P_P5COLON.GIF";
    images[37] = "http://www.writewellgroup.com/p456.html#colonp5";
    images[38] = "Colon: Introduce List or Appositive";
    
    images[39] = "http://www.writewellgroup.com/bigicons/P_P6.GIF";
    images[40] = "http://www.writewellgroup.com/p456.html#commap6";
    images[41] = "Comma needed: Separate Coordinate Adjs.";
    
    images[42] = "http://www.writewellgroup.com/bigicons/p_quotation.gif";
    images[43] = "http://www.writewellgroup.com/quotation.html";
    images[44] = "Quotation Marks Needed";
    
    images[45] = "http://www.writewellgroup.com/bigicons/P_P7.GIF";
    images[46] = "http://www.writewellgroup.com/quoting.html#commap7";
    images[47] = "Comma needed: Direct Discourse";
    
    images[48] = "http://www.writewellgroup.com/bigicons/P_P8.GIF";
    images[49] = "http://www.writewellgroup.com/p8.html";
    images[50] = "Comma needed: Protect Sentence Meaning";
    
    images[51] = "http://www.writewellgroup.com/bigicons/P_P9.GIF";
    images[52] = "http://www.writewellgroup.com/p9.html"
    images[53] = "Comma: Places, Dates, Titles, etc";
    
    images[54] = "http://www.writewellgroup.com/bigicons/P_P10A.GIF";
    images[55] = "http://www.writewellgroup.com/capitalization.html#cap10a";
    images[56] = "Capitalize First Letter of a Sentence";
    
    images[57] = "http://www.writewellgroup.com/bigicons/P_P10B.GIF";
    images[58] = "http://www.writewellgroup.com/capitalization.html#cap10b";
    images[59] = "Capitalize the Title of a Book";

    images[60] = "http://www.writewellgroup.com/bigicons/P_P11.GIF";
    images[61] = "http://www.writewellgroup.com/capitalization.html#cap11";
    images[62] = "Capitalize Proper Nouns & Adjs";
    
    images[63] = "http://www.writewellgroup.com/bigicons/P_P12.GIF";
    images[64] = "http://www.writewellgroup.com/capitalization.html#cap12";
    images[65] = "Capitalize Parts of Proper Nouns";
    
    images[66] = "http://www.writewellgroup.com/bigicons/P_P13A.GIF";
    images[67] = "http://www.writewellgroup.com/italicizing.htm#italicize13a";
    images[68] = "Italicize Words Under Discussion";
    
    images[69] = "http://www.writewellgroup.com/bigicons/P_P13B.GIF";
    images[70] = "http://www.writewellgroup.com/italicizing.htm#italicize13b";
    images[71] = "Italicize Letters Spoken as Letters";
    
    images[72] = "http://www.writewellgroup.com/bigicons/P_P13C.GIF";
    images[73] = "http://www.writewellgroup.com/italicizing.htm#italicize13c";
    images[74] = "Italicize Foreign Words or Phrases";
    
    images[75] = "http://www.writewellgroup.com/bigicons/P_P13D.GIF";
    images[76] = "http://www.writewellgroup.com/italicizing.htm#italicize13d";
    images[77] = "Italicize Names of Works & Ships";
    
    images[78] = "http://www.writewellgroup.com/bigicons/U_SPELLING.GIF";
    images[79] = "http://www.writewellgroup.com/spelling.html#spelling";
    images[80] = "Spelling";
    
    images[81] = "http://www.writewellgroup.com/bigicons/F_APOSTROPHE.GIF";
    images[82] = "http://www.writewellgroup.com/spelling.html#apostrophe";
    images[83] = "Apostrophe Errors";
  }
  else if(index == 1){//Category for Usage
    images[0] = "http://www.writewellgroup.com/bigicons/U_GOOD.GIF";
    images[1] = "http://www.writewellgroup.com/good.html#greatsentence";
    images[2] = "Great Sentence!";
                
    images[3] = "http://www.writewellgroup.com/icons/U_VERBTENSE.GIF";
    images[4] = "http://www.writewellgroup.com/tense.html";
    images[5] = "Verb Tense";
    
    images[6] = "http://www.writewellgroup.com/bigicons/u_agreement.gif";
    images[7] = "http://www.writewellgroup.com/verbagreement.html";
    images[8] = "Agreement of Verb with Noun";
    
    images[9] = "http://www.writewellgroup.com/bigicons/U_FRAGMENT.GIF";
    images[10] = "http://www.writewellgroup.com/modparfrag.html#fragment";
    images[11] = "Sentence Fragment";
  
    images[12] = "http://www.writewellgroup.com/icons/U_PASSIVE.GIF";
    images[13] = "http://www.writewellgroup.com/tense.html#passive";
    images[14] = "Passive Voice";
    
    images[15] = "http://www.writewellgroup.com/bigicons/U_AWKWARD.GIF";
    images[16] = "http://www.writewellgroup.com/awkward.html#awkward";
    images[17] = "Awkward Wording";
    
    images[18] = "http://www.writewellgroup.com/bigicons/U_CLARIFY.GIF";
    images[19] = "http://www.writewellgroup.com/awkward.html#unclear";
    images[20] = "Unclear Wording";
    
    images[21] = "http://www.writewellgroup.com/bigicons/U_WORDY.GIF";
    images[22] = "http://www.writewellgroup.com/awkward.html#wordy";
    images[23] = "Wordy";
    
    images[24] = "http://www.writewellgroup.com/bigicons/U_STRINGY.GIF";
    images[25] = "http://www.writewellgroup.com/awkward.html#stringy";
    images[26] = "Stringy";
    
    images[27] = "http://www.writewellgroup.com/bigicons/U_VARYSTRUCTURE.GIF";
    images[28] = "http://www.writewellgroup.com/awkward.html#vary";
    images[29] = "Vary Sentence Structure";
    
    images[30] = "http://www.writewellgroup.com/bigicons/U_MODIFIER.GIF";
    images[31] = "http://www.writewellgroup.com/modparfrag.html#modifier";
    images[32] = "Misplaced Modifier";
    
    images[33] = "http://www.writewellgroup.com/bigicons/U_PARALLEL.GIF";
    images[34] = "http://www.writewellgroup.com/modparfrag.html#parallelideas";
    images[35] = "Parallel Ideas";
    
    images[36] = "http://www.writewellgroup.com/bigicons/U_WORDCHOICE.GIF";
    images[37] = "http://www.writewellgroup.com/wordchoice.html#diction";
    images[38] = "Diction (Word Choice)";
    
    images[39] = "http://www.writewellgroup.com/bigicons/U_SLANG.GIF";
    images[40] = "http://www.writewellgroup.com/wordchoice.html#jargon";
    images[41] = "Jargon";
    
    images[42] = "http://www.writewellgroup.com/bigicons/u_proofread.gif";
    images[43] = "http://www.writewellgroup.com/proofread.html#proofread";
    images[44] = "Proofread";
    
    images[45] = "http://www.writewellgroup.com/bigicons/U_REDUNDANT.GIF";
    images[46] = "http://www.writewellgroup.com/wordchoice.html#redundant";
    images[47] = "Redundant Word Usage";
    
    images[48] = "http://www.writewellgroup.com/bigicons/U_PRONOUNREF.GIF";
    images[49] = "http://www.writewellgroup.com/pronoun.html#reference";
    images[50] = "Pronoun Reference";
    
    images[51] = "http://www.writewellgroup.com/bigicons/U_PRONOUNNUM.GIF";
    images[52] = "http://www.writewellgroup.com/pronoun.html#number";
    images[53] = "Pronoun Agreement";
    
    images[54] = "http://www.writewellgroup.com/bigicons/u_pronouncase.gif";
    images[55] = "http://www.writewellgroup.com/pronoun.html#case";
    images[56] = "Pronoun Case";
  }
  else if(index == 2){//Category for Content
    images[0] = "http://www.writewellgroup.com/bigicons/C_GOOD.GIF"
    images[1] = "http://www.writewellgroup.com/good.html#greatthinking";
    images[2] = "Great Thinking!";    
    
    images[3] = "http://www.writewellgroup.com/bigicons/C_DETAIL_NO.GIF";
    images[4] = "http://www.writewellgroup.com/contentissues.html#detail";
    images[5] = "Insufficient Detail";
    
    images[6] = "http://www.writewellgroup.com/bigicons/C_DISPROPORTIONATE_NO.GIF";
    images[7] = "http://www.writewellgroup.com/contentissues.html#disproportionate";
    images[8] = "Disappropiate Detail";
    
    images[9] = "http://www.writewellgroup.com/bigicons/C_SIGNIFIGANCE.gif";
    images[10] = "http://www.writewellgroup.com/significance.htm";
    images[11] = "Significance?";
    
    images[12] = "http://www.writewellgroup.com/bigicons/C_EXAGGERATED_NO.GIF";
    images[13] = "http://www.writewellgroup.com/contentissues.html#exaggeration";
    images[14] = "Point is Exaggerated";
    
    images[15] = "http://www.writewellgroup.com/bigicons/C_INTERPRETATION_NO.GIF";
    images[16] = "http://www.writewellgroup.com/facts.html#interpretation";
    images[17] = "Suspect Interpretation";
    
    images[18] = "http://www.writewellgroup.com/bigicons/C_OVERSIMPLIFICATION_NO.GIF";
    images[19] = "http://www.writewellgroup.com/contentissues.html#oversimplification";
    images[20] = "Oversimplification";
    
    images[21] = "http://www.writewellgroup.com/bigicons/C_NONTOPICAL_NO.GIF";
    images[22] = "http://www.writewellgroup.com/contentissues.html#nontopical";
    images[23] = "Non-topical Issue";
    
    images[24] = "http://www.writewellgroup.com/bigicons/C_FACTS_NO.GIF";
    images[25] = "http://www.writewellgroup.com/facts.html#facts";
    images[26] = "Facts are Garbled";
    
    images[27] = "http://www.writewellgroup.com/bigicons/C_FLAWEDSENTENCE_NO.GIF";
    images[28] = "http://www.writewellgroup.com/sequence.html";
    images[29] = "Flawed Idea Sequence";
    
    images[30] = "http://www.writewellgroup.com/bigicons/C_REDUNDANT.GIF";
    images[31] = "http://www.writewellgroup.com/contentissues.html#redundancy";
    images[32] = "Redundancy";
    
    images[33] = "http://www.writewellgroup.com/bigicons/C_QUOTE_NO.GIF";
    images[34] = "http://www.writewellgroup.com/quote.html#quoteneeded";
    images[35] = "Quote Needed";
    
    images[36] = "http://www.writewellgroup.com/bigicons/C_QUOTE.GIF";
    images[37] = "http://www.writewellgroup.com/quote.html#quote";
    images[38] = "Good Quotation";
    
    images[39] = "http://www.writewellgroup.com/bigicons/C_MLA.gif";
    images[40] = "http://www.writewellgroup.com/quoting.html#mla";
    images[41] = "MLA Citation Needed";
  }
  else{//Category for Form
    images[0] = "http://www.writewellgroup.com/bigicons/F_NEWPARA.GIF";
    images[1] = "http://www.writewellgroup.com/newparagraph.html";
    images[2] = "New Paragraph";                  
    
    images[3] = "http://www.writewellgroup.com/bigicons/F_TOPIC.GIF";
    images[4] = "http://www.writewellgroup.com/thesistopic.html#goodtopic";
    images[5] = "Good Topic Sentence";
    
    images[6] = "http://www.writewellgroup.com/bigicons/F_TOPIC_NO.GIF";
    images[7] = "http://www.writewellgroup.com/thesistopic.html#weaktopic";
    images[8] = "Weak Topic Sentence";
    
    images[9] = "http://www.writewellgroup.com/bigicons/F_TRANSITION.GIF";
    images[10] = "http://www.writewellgroup.com/transition.html";
    images[11] = "Good Transition";
    
    images[12] = "http://www.writewellgroup.com/bigicons/F_TRANSITION_NO.GIF";
    images[13] = "http://www.writewellgroup.com/transition.html#weaktransition";
    images[14] = "Poor Coherence";
    
    images[15] = "http://www.writewellgroup.com/bigicons/F_CONC_SENTENCE.GIF";
    images[16] = "http://www.writewellgroup.com/transition.html#goodconclusion";
    images[17] = "Good Concluding Sentence";
    
    images[18] = "http://www.writewellgroup.com/bigicons/F_CONC_SENTENCE_NO.GIF";
    images[19] = "http://www.writewellgroup.com/transition.html#weakconclusion";
    images[20] = "Weak Concluding Sentence";
    
    images[21] = "http://www.writewellgroup.com/bigicons/F_INTRO.GIF";
    images[22] = "http://www.writewellgroup.com/essay.html#goodintro";
    images[23] = "Good Introductory Paragraph";
    
    images[24] = "http://www.writewellgroup.com/bigicons/F_INTRO_NO.GIF";
    images[25] = "http://www.writewellgroup.com/essay.html#weakintro";
    images[26] = "Weak Introduction";
    
    images[27] = "http://www.writewellgroup.com/bigicons/F_BODY.GIF";
    images[28] = "http://www.writewellgroup.com/essay.html#goodbody";
    images[29] = "Good Body Paragraph";
    
    images[30] = "http://www.writewellgroup.com/bigicons/F_BODY_NO.GIF";
    images[31] = "http://www.writewellgroup.com/essay.html#weakbody";
    images[32] = "Weak Body Paragraph";
    
    images[33] = "http://www.writewellgroup.com/bigicons/F_CONCLUSION.GIF";
    images[34] = "http://www.writewellgroup.com/essay.html#goodconcl";
    images[35] = "Good Concluding Paragraph";
    
    images[36] = "http://www.writewellgroup.com/bigicons/F_CONCLUSION_NO.GIF";
    images[37] = "http://www.writewellgroup.com/essay.html#weakconcl";
    images[38] = "Weak Concluding Paragraph";
  }//End Initialization of Images
  
  return images;
}

function getImgValues(cat){
  var out = [];
  var length = getImageList(cat).length/3;
  if (cat == 0){
    for(var i = 0; i<length; i++){
      out[i] = -1;
    }
    //No exceptions (with images)
  }
  else if(cat == 1){//usage
    for(var i = 0; i<length; i++){//Generic Usage Generation
      out[i] = -1;//Generic
    }
    out[0] = 3; //Great Sentence Exception
    out[3] = -3; //Fragment Exception
  }
  else if(cat == 2){//content
    for(var i = 0; i<length; i++){
      out[i] = -2;//Generic
    }
    out[12] = 3;//Good Quote Choice
    out[0] = 3;//Good Thinking
  }
  else if(cat == 3){//form
    out[0] = -2;//Paragrah needed
    out[1] = 3;//Strong topic sentence
    out[2] = -3;//Weak topic sentence
    out[3] = 2; //Good transition
    out[5] = 2;//Good concluding sentence
    out[7] = 5; //Superb Intro
    out[8] = -3; //Not superb intro
    out[9] = 5; //Superb body
    out[10] = -3;//Not superb body
    out[11] = 5;//Superb body
    out[12] = -3; //not superb body;
  }
  return out;
}

function getFavorites(){
  var favorites = [];
  for(var i = 0; i<25; i++){
    var userProperty = UserProperties.getProperty("userFavorites"+i);
    if(userProperty!=null){
      favorites[i] = userProperty;
    }
    else{
      favorites[i] = "00";
    }
  }
  return favorites;
}

function populateGrid(){
  var app = UiApp.getActiveApplication();
  var numRows = 5;
  var numColumns = 5;
  var grid = app.createGrid(numRows, numColumns);
  var index = 0;
  var imageClickHandler = app.createServerHandler("imageClicker").addCallbackElement(app.getElementById("VPanel"));
  var favorites = getFavorites();
  for(var row = 0; row<numRows; row++){
    for(var column = 0; column<numColumns; column++){
      var userProperty = favorites[index];
      var cat = parseInt(userProperty.substring(0,1));
      var item = parseInt(userProperty.substring(1));
      var images = getImageList(cat);
      grid.setWidget(row, column, app.createImage().setId("gridFavoriteImage"+index).setUrl(images[item*3]).addClickHandler(imageClickHandler)
                   .setHeight(50).setWidth(50));
      index++;
    }
  }
  return grid;
}

function imageClicker(eventInfo){
  var app = UiApp.getActiveApplication();
  var orig = eventInfo.parameter.source;
  var num = +orig.replace(/\D/g, "");
  var index = parseInt(num);//orig.substring(orig.length-1));
  var userProperty = getFavorites()[index];
  var cat = parseInt(userProperty.substring(0,1));
  var item = parseInt(userProperty.substring(1));
  var images = getImageList(cat);
  var type = eventInfo.parameter.eventType;
  if(type == "click"){
    DocumentApp.getActiveDocument().getCursor().insertInlineImage(UrlFetchApp.fetch(images[item*3]).getBlob()).setLinkUrl(images[item*3+1]).setHeight(24).setWidth(24);
  }
  //app.add(app.createLabel(Logger.getLog()));
}


function multiDocumentSelect(){
  var app = UiApp.createApplication().setTitle("Generate Comments").setWidth(1000).setHeight(700);//UiApp.getActiveApplication();
  var selectionHandler = app.createServerHandler("multiDocGenerator");
   var box = app.createDocsListDialog().setDialogTitle("Select Files to Open").setMultiSelectEnabled(true).addSelectionHandler(selectionHandler).showDocsPicker();
   app.add(app.createVerticalPanel().setId("Panel"));
  //app.createDialogBox().add(app).show();
  DocumentApp.getUi().showDialog(app);
}

function multiDocGenerator(eventInfo){
  var items = eventInfo.parameter.items;
  var count = 0;
  var images = [];
  for(var i = 0; i<items.length; i++){
    var array = DocumentApp.openById(items[i].id).getBody().getImages();
    for(var x = 0; x<array.length; x++){
      images[count] = array[x];
      count++;
    }
  }
  for(var i = 0; i<images.length; i++){
    images[i] = images[i].getLinkUrl();
  }
  var out = "";
  if(images.length == 0){
    DocumentApp.getUi().alert("Error: No comments detected");
  }
  var body = DocumentApp.getActiveDocument().getBody();
  body.appendParagraph("Automatically Generated Comments").setAlignment(DocumentApp.HorizontalAlignment);
  var absoluteTotal = 0;
  for(var i = 0; i<4; i++){//Category loop
    out+= catDictionary(i) + ":\n";
    var catImages = getImageList(i);
    var values = getImgValues(i);
    var ptTotal = 0;
    for(var x = 1; x<catImages.length; x+=3){//Image check loop
      var count = getCount(images,catImages[x]);
      if(count!=0){//Only prints if #of errors is not 0
        var pts = count*values[parseInt((x-1)/3)];
        out+= "\t" + catImages[x+1] + ": " + count + "*("+ values[(x-1)/3] + "pts each) = " + (pts) + "pts\n"; //Generates individual comment
        ptTotal+= pts;
      }
    }
    out+= "Subtotal: " + ptTotal +"pts\n\n"; 
    absoluteTotal += ptTotal;
  }
  out+= "\nTotal: " + absoluteTotal +"pts";
  body.appendParagraph(out).editAsText().setForegroundColor("#ff0000").setBold(true);//Adds text and sets the color to red
  DocumentApp.getUi().alert("Done");
}

function documentSelect(){
  var app = UiApp.createApplication().setTitle("Open a Document With WriteWell").setWidth(1000).setHeight(700);//UiApp.getActiveApplication();
  var selectionHandler = app.createServerHandler("selectHandler");
   var box = app.createDocsListDialog().setDialogTitle("Select File to Open").addSelectionHandler(selectionHandler).showDocsPicker();
   app.add(app.createVerticalPanel().setId("Panel"));
  //app.createDialogBox().add(app).show();
  DocumentApp.getUi().showDialog(app);
  //DocumentApp.getUi().showDialog(app); 
  //return app;
}
  
function selectHandler(eventInfo){
  var document = eventInfo.parameter.items[0].id;//Selection???//The document iI want to copyy
  //var document = "1fMphfPcf3DQKoG-6UDMbScPLQFQuqrTYXiI-Jt1Gp7k";
  var app = UiApp.getActiveApplication();
  //var app = UiApp.createApplication();
  //var key = "1KCpjGJUnVfhi9SkpJPfDyURfW3Xds0CidYyoat3oT6U"; //Copys the WriteWell Template.
  var key = DocumentApp.getActiveDocument().getId(); //Copys from the current template
  var template = DocsList.getFileById(key);// get the template model
  var documentReal = DocsList.getFileById(document);
  var destination = documentReal.getName()+"-Edited";
  var baseDocId = DocsList.copy(template,destination).getId();// make a copy of firstelement and give it new basedocname build from the serie(to keep margins etc...)
  var newDoc = DocumentApp.openById(baseDocId);//The copy of the document!
  var body = newDoc.getBody().setText("");
  //var docToCopy = DocumentApp.openById(document);
  var trueDoc = DocumentApp.openById(document);
  var otherBody = trueDoc.getBody();//DocumentApp.openById(document).getBody();
  var totalElements = otherBody.getNumChildren();
  for( var j = 0; j < totalElements; ++j ) {
    var element = otherBody.getChild(j).copy();
    var type = element.getType();
    if( type == DocumentApp.ElementType.PARAGRAPH )
      body.appendParagraph(element);
    else if( type == DocumentApp.ElementType.TABLE )
      body.appendTable(element);
    else if( type == DocumentApp.ElementType.LIST_ITEM )
      body.appendListItem(element);
    else if( type == DocumentApp.ElementType.INLINE_IMAGE )
      body.appendImage(element);
    else if( type == DocumentApp.ElementType.PAGE_BREAK)
      body.appendPageBreak(element);
    else if( type == DocumentApp.ElementType.HORIZONTAL_RULE)
      body.appendHorizontalRule(element);

    // add other element types as you want

    //else
      //throw new Error("According to the doc this type couldn't appear in the body: "+type);
  }
  var editors = trueDoc.getEditors();
  for( var x = 0; x<editors.length; x++){
    newDoc.addEditor(editors[x]);
  }
  //app.close();
  //var app = UiApp.createApplication();
  var newLink = app.createAnchor("Open new Document", newDoc.getUrl());
  app.add(newLink);
  DocumentApp.getUi().showDialog(app);
  //app.add(app.createAnchor("Open the New Document", doc.getUrl()));
}

function getCount(array, searchTerm){
  var count = 0;
  for(var i = 0; i<array.length; i++){
    if(array[i] == searchTerm){
      count++;
    }
  }
  return count;
}

function catDictionary(index){
  if(index == 0){
    return "Punctuation";
  }
  else if(index == 1){
    return "Usage";
  }
  else if(index == 2){
    return "Content";
  }
  else{
    return "Form";
  }
}
