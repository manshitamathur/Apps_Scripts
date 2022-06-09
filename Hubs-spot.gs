var ui = SpreadsheetApp.getUi();
function onOpen(e){
  
  ui.createMenu("View Details").addItem("Get Emails by Label", "GetHubspotForm").addToUi();
  
}

function GetHubspotForm(){
//  var input = ui.prompt('Label Name', 'Enter the label name that is assigned to your emails:', Browser.Buttons.OK_CANCEL);
//  
//  if (input.getSelectedButton() == ui.Button.CANCEL){
//    return;
//  }
//  
  var label = GmailApp.getUserLabelByName("HubSpot_Form");
  var threads = label.getThreads();
  
  for(var i = threads.length - 1; i >=0; i--){
    var messages = threads[i].getMessages();
    
    for (var j = 0; j <messages.length; j++){
      var message = messages[j];
      if (message.isUnread()){
        extractDetails(message);
        GmailApp.markMessageRead(message);
      }
    }
    //threads[i].removeLabel(label);
    
  }
  
}

function extractDetails(message){
var emailData = {

//body: "Null",
Payment_Id: "Null",
Amount :"Null",
email:"Null",
contact : "Null"
}

var emailKeywords = {
Payment_Id: "pay",
Amount :"₹",
email:"Customer Details",
contact :"+91",
}

var bodycontents = message.getPlainBody();

var regExp;


emailData.contact = bodycontents.match(/[\+]?\d{10}/).toString().trim();
emailData.email = bodycontents.match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/gi).toString().trim();

var START = bodycontents.lastIndexOf("CONTACT"); 

var my_string = bodycontents.substring(START).toString();
var temp = my_string.replace("CONTACT","");

var Product_Details = bodycontents.search("Product:")
var START1 = bodycontents.lastIndexOf("Product:");
var END1 = bodycontents.indexOf("CONTACT"); 

var my_string1 = bodycontents.substring(START1,END1).toString();
var temp2 = my_string1.replace("Product:","");
var END3= getPosition(temp2,"\n", 2)
var temp3 = temp2.substring(0,END3);
var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

//var emailDataArr = [];

//for(var propName in emailData){
//emailDataArr.push(emailData[propName]);

//activeSheet.appendRow([emailData.contact,email,emailData.Amount,namess]);
activeSheet.appendRow([emailData.contact,emailData.email,temp,temp3]);
//,[emailData.contact,emailData.email,emailData.Amount]
}

/////////////emailData.contact,emailData.email
function getPosition(string, subString, index) {
  return string.split(subString, index).join(subString).length;
}



