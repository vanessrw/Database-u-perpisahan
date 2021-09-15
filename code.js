// DO GET FUNCTION
// function untuk mengirim file html ke orang yang membuka link
 
function doGet() {
 var html = HtmlService.createTemplateFromFile('IndexHtml').evaluate()
 .addMetaTag("viewport", "width=device-width, initial-scale=1")
 .setTitle("Palancas")
 .setFaviconUrl("https://sanurbsd-tng.sch.id/assets/img/logo_2.png"); // icon webnya serviam, nanti bisa diganti
 
 return html;
};
 
// SET CSS
// function utk masukin css ke index.html
 
function includeStyleFile() {
 return HtmlService.createHtmlOutputFromFile("IndexCss").getContent();
}
 
function includeScriptFile() {
 return HtmlService.createHtmlOutputFromFile("IndexJs").getContent();
}
 
// CREATE GOOGLE SLIDE LINK 
// function untuk membuat google slide tersendiri masing-masing untuk setiap siswa
// link spreadsheet prototype: https://docs.google.com/spreadsheets/d/1awlbVXy1XY2qyQGAYO3Qqew9vYhPyOEcqxoEwuAm5SM/edit#gid=0
 
// SEND GOOGLE SLIDE VIA EMAIL 
// function untuk mengirim link google slide yang sudah dibuat diatas kepada setiap siswa. alamat email diambil dari google spreadsheet
// link spreadsheet prototype: https://docs.google.com/spreadsheets/d/1awlbVXy1XY2qyQGAYO3Qqew9vYhPyOEcqxoEwuAm5SM/edit#gid=0
 
// RENDER STUDENTS
// ambil data dari google spreadsheet, trud jadiin html
var studArr = [];
function renderStudents(studentClass) {
 studArr.length = 0;
 let url = "https://docs.google.com/spreadsheets/d/1awlbVXy1XY2qyQGAYO3Qqew9vYhPyOEcqxoEwuAm5SM/edit#gid=0"
 let Sheet = SpreadsheetApp.openByUrl(url);
 let ClassSheet = Sheet.getSheetByName(studentClass)
 
 let RowCount = ClassSheet.getLastRow();
 for (i=2; i<=RowCount; i++) {
   const x = i.toString();
     var itemObj = {
       id: 0,
       name: "",
       url: ""
     }
   itemObj.id = ClassSheet.getRange('A'+x).getValues().toString();
   itemObj.name = ClassSheet.getRange('B'+x).getValues().toString();
   itemObj.url = ClassSheet.getRange('E'+x).getValues().toString();
   studArr.push(itemObj);
 
   };
 
   return studArr;
};
 
// PREVIEW SLIDE
function showPreview() {
 var e = SlidesApp.openByUrl("https://docs.google.com/presentation/d/1pL51UmZYlNuQd8wZlMWC2GuOES-6n2EbHgAUrOJ3sNI/edit#slide=id.ge057e29d2d_14_14")
 var blob = Utilities.newBlob(e);
 var image = blob.getAs("image/jpeg");
 var file = DriveApp.createFile(image);
 
 let Destination = DriveApp.getFolderById("1tPWJGSWX_TE9r4j00fsXD4tRXTqzaQy8");
 Destination.createFile(file)
}
 
 
// CREATE NEW SLIDE/PAGE 
// function untuk membuat slide baru dalam google slide setiap kali terdapat pengirim yang menuliskan palancas ke orang tersebut
// link google slide prototype: https://docs.google.com/presentation/d/1pL51UmZYlNuQd8wZlMWC2GuOES-6n2EbHgAUrOJ3sNI/edit#slide=id.p
 

 
function createSlides(Array, Url){
  /*
  Array = [
    {
      writting: "hoiii",
      sender: "intiw",
      theme: 2,
      imgUrl: "https://drive.google.com/uc?id=1Zmb4Rp8V4tVklYbQIpW0XDfq7vKY9dvF&export=download"
    }
  ]*/
  for(i=0; i<Array.length; i++){
    appendSlide(Array, i, Url);
  }
}
 
function appendSlide(Array, i, Url) {

  let Presentation = SlidesApp.openByUrl(Url);
  let Template = SlidesApp.openByUrl('https://docs.google.com/presentation/d/1MOlAHiZ8tamlflpQV29UvAP6wOOeUQ9QoJcbgc98fHU/edit?usp=sharing');
  
 
  if(Array[i].imgUrl != ''){
    Presentation.insertSlide(Presentation.getSlides().length - 1, Template.getSlides()[Array[i].theme], SlidesApp.SlideLinkingMode.LINKED); 
    var Slide = Presentation.getSlides()[Presentation.getSlides().length - 2];
    for (x=0; x<Slide.getPageElements().length ; x++){
      if (Slide.getPageElements()[x].getTitle() == 'Img'){
        /*  // ini lain v
        var decoded = Utilities.base64Decode(Array[i].imgUrl);
        var blob = Utilities.newBlob(decoded, MimeType.JPEG, "nameOfImage");
        // ini lain ^*/
        /*
        let base64 = Array[i].imgUrl;
        let base64Edit = base64.replace("data:image/jpeg;base64,", "");
        var decoded = Utilities.base64Decode(base64Edit);
        var blob = Utilities.newBlob(decoded, MimeType.JPEG, "nameOfImage");
        let file = currentFolder.createFile(blob)
        let sharedFile = file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.OWNER);
        let fileUrl = sharedFile.getUrl();
*/
        var currentFolder = DriveApp.getFolderById("12D7Ylcd425CkGNowBKjnvuCXmGSjanLX");

        var base64 = Array[i].imgUrl
        let base64Edit
        if (base64.search("data:image/jpeg;base64,") == 0 ) {
          base64Edit = base64.replace("data:image/jpeg;base64,", "");
        } else if (base64.search("data:image/png;base64,") == 0 ) {
          base64Edit = base64.replace("data:image/png;base64,", "");
        } else if (base64.search("data:image/gif;base64,") == 0 ) {
          base64Edit = base64.replace("data:image/gif;base64,", "");
        }
       
        var decoded = Utilities.base64Decode(base64Edit);
        var blob = Utilities.newBlob(decoded, MimeType.JPEG, "nameOfImage");
        let file = currentFolder.createFile(blob)
        let sharedFile = file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT);
        let fileUrl = sharedFile.getDownloadUrl();
        Slide.getPageElements()[x].asImage().replace(fileUrl);

        let Width = Slide.getPageElements()[x].asImage().getWidth();
        let Height = Slide.getPageElements()[x].asImage().getHeight();
        let Top = Slide.getPageElements()[x].asImage().getTop();
        let Left = Slide.getPageElements()[x].asImage().getLeft();
        if(Width < Template.getSlides()[0].getImages()[0].getWidth()){
          for (y=0; y<Slide.getPageElements().length ; y++){
            if (Slide.getPageElements()[y].getTitle() == 'Border'){
              if(Slide.getPageElements()[y].asShape().getFill().getSolidFill().getColor().getColorType() == 'RGB'){
                z = 8;
              }else{
                z = 0;
              }
              Slide.getPageElements()[y].setWidth(Width+28);
              Slide.getPageElements()[y].setLeft(Left-14+z);
            }
          }
        }if(Height < Template.getSlides()[0].getImages()[0].getHeight()){
          for (y=0; y<Slide.getPageElements().length ; y++){
            if (Slide.getPageElements()[y].getTitle() == 'Border'){
              if(Slide.getPageElements()[y].asShape().getFill().getSolidFill().getColor().getColorType() == 'RGB'){
                z = 8;
              }else{
                z = 0;
              }
              Slide.getPageElements()[y].setHeight(Height+65);
              Slide.getPageElements()[y].setTop(Top-15+z);
            }
          }
        }
      }
    }
    insertText(Slide, Array, i);
  }else{
    Presentation.insertSlide(Presentation.getSlides().length - 1, Template.getSlides()[Array[i].theme+1], SlidesApp.SlideLinkingMode.LINKED); 
    var Slide = Presentation.getSlides()[Presentation.getSlides().length - 2];
    insertText(Slide, Array, i);
  }
}
 
function insertText(Slide, Array, i){
  for (x=0; x<Slide.getPageElements().length ; x++){
    if (Slide.getPageElements()[x].getTitle() == 'Msg'){
      Slide.getPageElements()[x].asShape().getText().appendText(Array[i].writting);
      Slide.getPageElements()[x].asShape().getText().getTextStyle().setFontSize(30);
      if(Array[i].sender != ''){
        Slide.getPageElements()[x].asShape().getText().appendParagraph(' ');
        Slide.getPageElements()[x].asShape().getText().appendParagraph(Array[i].sender).getRange().getTextStyle().setFontSize(30);
      };
      if(Array[i].theme == 4){
        Slide.getPageElements()[x].asShape().getText().getTextStyle().setFontFamily('Gaegu').setForegroundColor('#5e7db0');
      }else{
        Slide.getPageElements()[x].asShape().getText().getTextStyle().setFontFamily('Gaegu').setForegroundColor('#f9fbe6');
      };
      Slide.getPageElements()[x].asShape().getText().getTextStyle().setBold(true);
      Slide.getPageElements()[x].asShape().getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      Slide.unlink();
    }
  }
}


 


function saveFile(e) {
  var blob = Utilities.newBlob(e.bytes, e.mimeType, e.filename);
  
  let Destination = DriveApp.getFolderById("12D7Ylcd425CkGNowBKjnvuCXmGSjanLX");
  var url = Destination.createFile(blob).getUrl();
  Logger.log(url)

  let arr = [{
    sender: "ada deh",
    writting: "tulisan",
    imgUrl: url,
    theme: 2
  }]
  createSlides(arr, 'https://docs.google.com/presentation/d/1pL51UmZYlNuQd8wZlMWC2GuOES-6n2EbHgAUrOJ3sNI/edit#slide=id.SLIDES_API918760830_0')

}

function base64toBlob(url) {
  var currentFolder = DriveApp.getFolderById("12D7Ylcd425CkGNowBKjnvuCXmGSjanLX");

  var base64 = url
  var decoded = Utilities.base64Decode(base64);
  var blob = Utilities.newBlob(decoded, MimeType.JPEG, "nameOfImage");
  let file = currentFolder.createFile(blob)
  let sharedFile = file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT);
  let fileUrl = sharedFile.getDownloadUrl();
  Logger.log(fileUrl)
} 

 
/**
 * @OnlyCurrentDoc
*/
 
/**
 * Change these to match the column names you are using for email 
 * recipient addresses and email sent column.
*/
const RECIPIENT_COL  = "Recipient";
const EMAIL_SENT_COL = "Email Sent";
 
/** 
 * Creates the menu item "Mail Merge" for user to run scripts on drop-down.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Merge')
      .addItem('Send Emails', 'sendEmails')
      .addToUi();
}
 
/**
 * Send emails from sheet data.
 * @param {string} subjectLine (optional) for the email draft message
 * @param {Sheet} sheet to read data from
*/
function sendEmails(subjectLine, sheet=SpreadsheetApp.getActiveSheet()) {
  // option to skip browser prompt if you want to use this code in other projects
  if (!subjectLine){
    subjectLine = Browser.inputBox("Mail Merge", 
                                      "Type or copy/paste the subject line of the Gmail " +
                                      "draft message you would like to mail merge with:",
                                      Browser.Buttons.OK_CANCEL);
                                      
    if (subjectLine === "cancel" || subjectLine == ""){ 
    // if no subject line finish up
    return;
    }
  }
  
  // get the draft Gmail message to use as a template
  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);
  
  // get the data from the passed sheet
  const dataRange = sheet.getDataRange();
  // Fetch displayed values for each row in the Range HT Andrew Roberts 
  // https://mashe.hawksey.info/2020/04/a-bulk-email-mail-merge-with-gmail-and-google-sheets-solution-evolution-using-v8/#comment-187490
  // @see https://developers.google.com/apps-script/reference/spreadsheet/range#getdisplayvalues
  const data = dataRange.getDisplayValues();

  // assuming row 1 contains our column headings
  const heads = data.shift(); 
  
  // get the index of column named 'Email Status' (Assume header names are unique)
  // @see http://ramblings.mcpher.com/Home/excelquirks/gooscript/arrayfunctions
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);
  
  // convert 2d array into object array
  // @see https://stackoverflow.com/a/22917499/1027723
  // for pretty version see https://mashe.hawksey.info/?p=17869/#comment-184945
  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

  // used to record sent emails
  const out = [];

  // loop through all the rows of data
  obj.forEach(function(row, rowIdx){
    // only send emails is email_sent cell is blank and not hidden by filter
    if (row[EMAIL_SENT_COL] == ''){
      try {
        const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);

        // @see https://developers.google.com/apps-script/reference/gmail/gmail-app#sendEmail(String,String,String,Object)
        // if you need to send emails with unicode/emoji characters change GmailApp for MailApp
        // Uncomment advanced parameters as needed (see docs for limitations)
        GmailApp.sendEmail(row[RECIPIENT_COL], msgObj.subject, msgObj.text, {
          htmlBody: msgObj.html,
          // bcc: 'a.bbc@email.com',
          // cc: 'a.cc@email.com',
          // from: 'an.alias@email.com',
          // name: 'name of the sender',
          // replyTo: 'a.reply@email.com',
          // noReply: true, // if the email should be sent from a generic no-reply email address (not available to gmail.com users)
          attachments: emailTemplate.attachments,
          inlineImages: emailTemplate.inlineImages
        });
        // modify cell to record email sent date
        out.push([new Date()]);
      } catch(e) {
        // modify cell to record error
        out.push([e.message]);
      }
    } else {
      out.push([row[EMAIL_SENT_COL]]);
    }
  });
  
  // updating the sheet with new data
  sheet.getRange(2, emailSentColIdx+1, out.length).setValues(out);
  
  /**
   * Get a Gmail draft message by matching the subject line.
   * @param {string} subject_line to search for draft message
   * @return {object} containing the subject, plain and html message body and attachments
  */
  function getGmailTemplateFromDrafts_(subject_line){
    try {
      // get drafts
      const drafts = GmailApp.getDrafts();
      // filter the drafts that match subject line
      const draft = drafts.filter(subjectFilter_(subject_line))[0];
      // get the message object
      const msg = draft.getMessage();

      // Handling inline images and attachments so they can be included in the merge
      // Based on https://stackoverflow.com/a/65813881/1027723
      // Get all attachments and inline image attachments
      const allInlineImages = draft.getMessage().getAttachments({includeInlineImages: true,includeAttachments:false});
      const attachments = draft.getMessage().getAttachments({includeInlineImages: false});
      const htmlBody = msg.getBody(); 

      // Create an inline image object with the image name as key 
      // (can't rely on image index as array based on insert order)
      const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj) ,{});

      //Regexp to search for all img string positions with cid
      const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
      const matches = [...htmlBody.matchAll(imgexp)];

      //Initiate the allInlineImages object
      const inlineImagesObj = {};
      // built an inlineImagesObj from inline image matches
      matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);

      return {message: {subject: subject_line, text: msg.getPlainBody(), html:htmlBody}, 
              attachments: attachments, inlineImages: inlineImagesObj };
    } catch(e) {
      throw new Error("Oops - can't find Gmail draft");
    }

    /**
     * Filter draft objects with the matching subject linemessage by matching the subject line.
     * @param {string} subject_line to search for draft message
     * @return {object} GmailDraft object
    */
    function subjectFilter_(subject_line){
      return function(element) {
        if (element.getMessage().getSubject() === subject_line) {
          return element;
        }
      }
    }
  }
  
  /**
   * Fill template string with data object
   * @see https://stackoverflow.com/a/378000/1027723
   * @param {string} template string containing {{}} markers which are replaced with data
   * @param {object} data object used to replace {{}} markers
   * @return {object} message replaced with data
  */
  function fillInTemplateFromObject_(template, data) {
    // we have two templates one for plain text and the html body
    // stringifing the object means we can do a global replace
    let template_string = JSON.stringify(template);

    // token replacement
    template_string = template_string.replace(/{{[^{}]+}}/g, key => {
      return escapeData_(data[key.replace(/[{}]+/g, "")] || "");
    });
    return  JSON.parse(template_string);
  }

  /**
   * Escape cell data to make JSON safe
   * @see https://stackoverflow.com/a/9204218/1027723
   * @param {string} str to escape JSON special characters from
   * @return {string} escaped string
  */
  function escapeData_(str) {
    return str
      .replace(/[\\]/g, '\\\\')
      .replace(/[\"]/g, '\\\"')
      .replace(/[\/]/g, '\\/')
      .replace(/[\b]/g, '\\b')
      .replace(/[\f]/g, '\\f')
      .replace(/[\n]/g, '\\n')
      .replace(/[\r]/g, '\\r')
      .replace(/[\t]/g, '\\t');
  };
}