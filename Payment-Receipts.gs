//var ui = SpreadsheetApp.getUi();
//function onOpen(e){
//  
//  ui.createMenu("View Details").addItem("Get Emails by Label", "PaymentsuccessTheMentor").addToUi();
//  
//}

function PaymentsuccessTheMentor(){
//  var input = ui.prompt('Label Name', 'Enter the label name that is assigned to your emails:', Browser.Buttons.OK_CANCEL);
//  
//  if (input.getSelectedButton() == ui.Button.CANCEL){
//    return;
//  }
//  
  var label = GmailApp.getUserLabelByName("Razorpay-Mini-Receipts");
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
    
    
  }
  
  ////////////////////////////////////////
  
  
  var label1 = GmailApp.getUserLabelByName("Payment_MTP");
  var threads1 = label1.getThreads();
  
  for(var i = threads1.length - 1; i >=0; i--){
    var messages1 = threads1[i].getMessages();
    
    for (var j = 0; j <messages1.length; j++){
      var message1 = messages1[j];
      if (message1.isUnread()){
        extractDetails1(message1);
        GmailApp.markMessageRead(message1);
      }
    }
    
    
  }
  
  /////////////////////////////////////////
  
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

//var regExp;

//regExp = new RegExp("(?<=" + emailKeywords.Payment_Id + ").*");
//emailData.Payment_Id = bodycontents.match(regExp).toString().trim();

//regExp = new RegExp("(?<="+"₹"+").[0-9].[0-9]+[0-9]+");
//emailData.Amount = bodycontents.match(regExp).toString().trim();

//regExp = new RegExp("(?<=" + emailKeywords.email +").*");
//emailData.email = bodycontents.match(regExp).toString().trim();

//

//regExp= new RegExp("/^\+91\d{10}$/");
//regExp = new RegExp("(?<=" + emailKeywords.contact + ").[0-9].[0-9]+[0-9]+");
var check_pt1 =bodycontents.search("This means that money was deducted from the customer's account.");
var check_pt2 =bodycontents.search("Refund has been initiated");
if(check_pt1== -1 &&  check_pt2==-1){
emailData.contact = bodycontents.match(/[\+]?\d{12}|\(\d{3}\)\s?-\d{6}/).toString().trim();
emailData.email = bodycontents.match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/gi).toString().trim();
const emailarr = emailData.email.split(",");
var email = emailarr[emailarr.length - 1];

emailData.Amount = bodycontents.match(/₹\S+/g).toString().trim();
var START = (emailData.Amount.length)/2;
let amt = emailData.Amount.slice(START+1, -1);
//fetching payment id
var start_ = bodycontents.lastIndexOf("Payment Id");
var END = bodycontents.indexOf("Amount");
//var temp_len = "Payment Id".length; //10
var pay_id = bodycontents.substring(start_+10,END);

var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mini-Receipts");


activeSheet.appendRow([pay_id,emailData.contact,email,amt]);

}
}




/////////////////////////
function extractDetails1(message){
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

//regExp = new RegExp("(?<=" + emailKeywords.Payment_Id + ").*");
//emailData.Payment_Id = bodycontents.match(regExp).toString().trim();

//regExp = new RegExp("(?<=" + emailKeywords.Amount +").[0-9].[0-9]+[0-9]+");
//emailData.Amount = bodycontents.match(regExp).toString().trim();

//regExp = new RegExp("(?<=" + emailKeywords.email +").*");
//emailData.email = bodycontents.match(regExp).toString().trim();

//

//regExp= new RegExp("/^\+91\d{10}$/");
//regExp = new RegExp("(?<=" + emailKeywords.contact + ").[0-9].[0-9]+[0-9]+");
 emailData.contact = bodycontents.match(/[\+]?\d{12}|\(\d{3}\)\s?-\d{6}/).toString().trim();
  emailData.email = bodycontents.match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/gi).toString().trim();
const emailarr = emailData.email.split(",");
var email = emailarr[emailarr.length - 1];

emailData.Amount = bodycontents.match(/₹\S+/g).toString().trim();
//fetching payment id
var start_ = bodycontents.lastIndexOf("Payment Id");
var END_ = bodycontents.indexOf("Paid On");
//var temp_len = "Payment Id".length; //10
var pay_id = bodycontents.substring(start_+10,END_);
var namess ;
if (bodycontents.search("First Name")!= -1)
{
var START = bodycontents.lastIndexOf("First Name");
var END = bodycontents.indexOf("Email"); 

var my_string = bodycontents.substring(START,END);
var temp = my_string.replace("First Name","");
var names = temp.replace("Last Name","");
var new_ = names.replace("\n"," ");
const name_arr= names.split("\n").toString();
const name_arr_new= name_arr.split(",").toString();

namess = names;
}
else{
var START = bodycontents.lastIndexOf("Name");
var END = bodycontents.indexOf("Email"); 

var my_string = bodycontents.substring(START,END);
var temp = my_string.replace("Name","");
namess=temp.toString();


}
var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Detailed-Receipts");

//var emailDataArr = [];

//for(var propName in emailData){
//emailDataArr.push(emailData[propName]);

activeSheet.appendRow([pay_id,emailData.contact,email,emailData.Amount,namess]);
//,[emailData.contact,emailData.email,emailData.Amount]
}

