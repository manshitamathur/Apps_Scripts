//create and send pdfs
function onOpen(e){
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu("Additional Options").addItem("Send Mails", "createBulkPDFs").addToUi();
  
}
function createBulkPDFs(){
   const docFile = DriveApp.getFileById("1D3W4f1mnbdh0VstVwUH6SXpbrwUgiiEpM1yOpz3Em4w");
   const tempFolder = DriveApp.getFolderById("1c4DPJO3gRu3WGkgyoG-Dsz2eUQ8_ElRE");
   const pdfFolder = DriveApp.getFolderById("1p5ftXR5ul01nyeiDs75ASvk-FHNzUegT");
   const currentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("onboarding-Interns");
    
   const data = currentSheet.getRange(2,1,currentSheet.getLastRow() - 1,15).getValues();
   
   let errors=[];
   var row_no=-1;
   data.forEach(row => {
        var res = currentSheet.getRange(findRow(row[2]),12,currentSheet.getLastRow() - 1,1).getValue();
        if(res == "") 
        {
         try{
          createPDF(row[1],row[5],row[0],row[2],row[1]+" "+row[0],docFile,tempFolder,pdfFolder,currentSheet);
          //errors.push(["Success"]);
         }
         catch(err){
         //errors.push([err]);
         }
       
       //currentSheet.getRange(2,6,currentSheet.getLastRow() - 1,1).setValue(errors);
        }  
    });  
   
}

function sendEmail(email,pdfFile,designation)
{
  GmailApp.sendEmail("manshitamathur@gmail.com", "Offer Letter - THE MENTOR", "we are delighted to send you an offer as a "+designation, {
  attachments  : [pdfFile],
  name : "The Mentor"
  });
}

function createPDF(name,designation,date,email,pdfName,docFile,tempFolder,pdfFolder,currentSheet) {

//   let firstName = "Manshita Sanjay";
//   let lastName = "Mathur";
//   let designation = "Django Intern";
//   let date = "June 01 ,2021";
 
 //doc id (offerletter Template)              1D3W4f1mnbdh0VstVwUH6SXpbrwUgiiEpM1yOpz3Em4w 
   //temp folder                                1c4DPJO3gRu3WGkgyoG-Dsz2eUQ8_ElRE
   //pdf folder(offer letters folder in drive)  1p5ftXR5ul01nyeiDs75ASvk-FHNzUegT
   
   const tempFile = docFile.makeCopy(tempFolder);
   const tempDocFile = DocumentApp.openById(tempFile.getId());
   const body = tempDocFile.getBody();
   body.replaceText("{name}",name);
  var dict = {"hr":"Human Resource",
   "gd":"Marketing","cc":"Marketing",
   "tech":"Technical","smm":"Marketing",
   "content creation":"Marketing",
   "video creation":"Marketing",
   "sales":"Sales",
   "mi":"Management",
   "bda":"Business Development Associate"};
   body.replaceText("{designation}",dict[designation.toLowerCase().trim()]+" Intern");
   const temp1 = date.toString();
   const day=temp1.split(" ");
   var onboarding_date = day[1]+" "+day[2]+","+day[3];
   body.replaceText("{doj}",onboarding_date);
   tempDocFile.saveAndClose();
   
   const pdfContentBlob = tempFile.getAs(MimeType.PDF);
   const pdfFile = pdfFolder.createFile(pdfContentBlob).setName(pdfName+"-Offer-Letter");
   row_no =findRow(email); 
   currentSheet.getRange(row_no,12).setValue(pdfFile.getName());
   currentSheet.getRange(row_no,13).setValue(pdfFile.getUrl());
  // currentSheet.getRange(row_no,6).setValue(errors);
  let e_=[];
  var email_status = currentSheet.getRange(row_no,14).getValue();
  if (email_status=="")  
 {
  try{
   sendEmail(email,pdfFile,designation);
   e_.push("sent");
   
  } catch(e)
  {
    e_.push("pending");
  }
  finally{
  tempFolder.removeFile(tempFile);
  currentSheet.getRange(row_no,14).setValue(e_);
  }
  
  } 
   
   
   
  
}


function findRow(searchVal) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var columnCount = sheet.getDataRange().getLastColumn();

  var i = data.flat().indexOf(searchVal); 
  var columnIndex = i % columnCount
  var rowIndex = ((i - columnIndex) / columnCount);

  Logger.log({columnIndex, rowIndex }); // zero based row and column indexes of searchVal

  return i >= 0 ? rowIndex + 1 : -1;
}



