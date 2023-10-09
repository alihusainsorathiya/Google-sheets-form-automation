var START_ROW = 2;
var START_COL = 1;
var subject = "Your Email Subject Here";

// Your Slack Email
var slackEmail = 'channelemail@yourdomain.slack.com';
function sendEmailToUser(){
 SpreadsheetApp.flush();
//  Utilities.sleep(1000);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // var sheet = ss.getSheetByName("Form Responses 1");
 
  var sheet = ss.getSheetByName("report");
  // var data = sheet.getRange(START_ROW,START_COL,sheet.getLastRow()-1,23).getValues();

 Utilities.sleep(2000);
//  SpreadsheetApp.flush();
var lastRow = sheet.getLastRow();
var lastColumn = sheet.getLastColumn();
Logger.log(lastRow);
Logger.log(lastColumn);
// var lastColumn = sheet.getLastColumn().g;
  var lastRowValue = sheet.getRange(lastRow,1,1,60).getValues();
  Logger.log("The value of the last row in column A is: " + lastRowValue);


Logger.log(lastRowValue);
  // data.forEach(function(row,i){

// Logger.log("6"+lastRowValue[0][6]);

lastRowValue.forEach(function(row,i){
Logger.log("last row: " +row[1]);


    // var name = row[1];
    // var number = row[2];
    // var email = row[41];
    // var category = row[4];
    // var details = row[5];

    // var subject = "CAFU - Free Vehicle Health Checkup Report";
    // var EmailBody = "Dear " + name + "," + "<br><br>"+
    //                 "We recieved your problem with the following details" + "<br><br>"+
    //                 "Data" +data[i]+
    //                 // "Description : " + details + "<br><br>"+
    //                 // "Category : " + category + "<br><br>"+
    //                 // "A Customer Service Executive will contact you shortly "+ "<br><br>"
    //                 // "Regards,"+ "<br><br>"+
                    
    //                 "ABC Solutions."+ "<br><br>";
// console.log("Row :" +row[17]);



// var EmailBody = "<b>Timestamp</b>:"+ row[0]+"<br><br>"+
// "<b>Order ID</b>:"+ row[1]+"<br><br>"+
// "<b>VIN Number</b>:"+ row[2]+"<br><br>"+
// "<b>Pilot Name</b>:"+ row[3]+"<br><br>"+
// "<b>Mileage (Kms)</b>:"+ row[4]+"<br><br>"+
// "<b>Next service Mileage (Kms)</b>:"+ row[5]+"<br><br>"+
// "<b>MINOR SERVICE [Oil Filter check-up]</b>:"+ row[6]+"<br><br>"+
// "<b>MINOR SERVICE [Oil grade check-up]</b>:"+ row[7]+"<br><br>"+
// "<b>MINOR SERVICE [Engine coolant levels]</b>:"+ row[8]+"<br><br>"+
// "<b>MINOR SERVICE [Steering Fluid levels]</b>:"+ row[9]+"<br><br>"+
// "<b>MINOR SERVICE [Washer fluid levels]</b>:"+ row[10]+"<br><br>"+
// "<b>MINOR SERVICE [Wiper blade check-up]</b>:"+ row[11]+"<br><br>"+
// "<b>Fire Extinguisher availability</b>:"+ row[12]+"<br><br>"+
// "<b>Fire Extinguisher Expiry Date</b>:"+ row[13]+"<br><br>"+
// "<b>Battery Health</b>:"+ row[14]+"<br><br>"+
// "<b>Battery Make</b>:"+ row[15]+"<br><br>"+
// "<b>Front Right Tyre PSI</b>:"+ row[16]+"<br><br>"+
// "<b>Front Left Tyre PSI</b>:"+ row[17]+"<br><br>"+
// "<b>Back Right Tyre PSI</b>:"+ row[18]+"<br><br>"+
// "<b>Back Left Tyre PSI</b>:"+ row[19]+"<br><br>"+
// "<b>Tyre Health Inspection [Front - Right]</b>:"+ row[20]+"<br><br>"+
// "<b>Tyre Health Inspection [Front - Left]</b>:"+ row[21]+"<br><br>"+
// "<b>Tyre Health Inspection [Rear - Right]</b>:"+ row[22]+"<br><br>"+
// "<b>Tyre Health Inspection [Rear - Left]</b>:"+ row[23]+"<br><br>"+
// "<b>Tyre Health comments</b>:"+ row[24]+"<br><br>"+
// "<b>Brake Pad Inspection [Front - Right]</b>:"+ row[25]+"<br><br>"+
// "<b>Brake Pad Inspection [Front - Left]</b>:"+ row[26]+"<br><br>"+
// "<b>Brake Pad Inspection [Rear - Right]</b>:"+ row[27]+"<br><br>"+
// "<b>Brake Pad Inspection [Rear - Left]</b>:"+ row[28]+"<br><br>"+
// "<b>Brake Fluid Check-up</b>:"+ row[29]+"<br><br>"+
// "<b>Air Filter Check</b>:"+ row[30]+"<br><br>"+
// "<b>Lights Inspection [Front - Right]</b>:"+ row[31]+"<br><br>"+
// "<b>Lights Inspection [Front - Left]</b>:"+ row[32]+"<br><br>"+
// "<b>Lights Inspection [Rear - Right]</b>:"+ row[33]+"<br><br>"+
// "<b>Lights Inspection [Rear - Left]</b>:"+ row[34]+"<br><br>"+
// "<b>Lights Inspection [Right Brake Light]</b>:"+ row[35]+"<br><br>"+
// "<b>Lights Inspection [Left Brake Light]</b>:"+ row[36]+"<br><br>"+
// "<b>General Feedback</b>:"+ row[37]+"<br><br>"+
// "<b>VAS OrderID</b>:"+ row[38]+"<br><br>"+
// "<b>Customer ID</b>:"+ row[39]+"<br><br>"+
// "<b>Customer Name</b>:"+ row[40]+"<br><br>"+
// "<b>Customer Email</b>:"+ row[41]+"<br><br>"+
// "<b>Customer Phone Number (without dialcode should start without 0)</b>:"+ row[42]+"<br><br>"+
// "<b>Vehicle Make</b>:"+ row[43]+"<br><br>"+
// "<b>Vehicle Model</b>:"+ row[44]+"<br><br>"+
// "<b>Vehicle Year</b>:"+ row[45]+"<br><br>"+
// "<b>Vehicle Number Plate</b>:"+ row[46]+"<br><br>"+
// "<b>Vehicle Number Plate Emirati</b>:"+ row[47]+"<br><br>"+
// "<b>Timeslot</b>:"+ row[48]+"<br><br>";




// // Logger.log(EmailBody);

// var header = "";
// var footer = "";


// For Multiple recipients
// var bcc1and2 = "xyz@example.com,abc@example.com";

    // GmailApp.sendEmail(email,subject,"",{'bcc':slackEmail, htmlBody:EmailBody});


  });



// var EmailBody = "<b>Timestamp</b>:"+ lastRowValue[0][0]+"<br><br>"+
// "<b>Order ID</b>:"+ lastRowValue[0][1]+"<br><br>"+
// "<b>VIN Number</b>:"+ lastRowValue[0][2]+"<br><br>"+
// "<b>Pilot Name</b>:"+ lastRowValue[0][3]+"<br><br>"+
// "<b>Mileage (Kms)</b>:"+ lastRowValue[0][4]+"<br><br>"+
// "<b>Next service Mileage (Kms)</b>:"+ lastRowValue[0][5]+"<br><br>"+
// "<b>MINOR SERVICE [Oil Filter check-up]</b>:"+ lastRowValue[0][6]+"<br><br>"+
// "<b>MINOR SERVICE [Oil grade check-up]</b>:"+ lastRowValue[0][7]+"<br><br>"+
// "<b>MINOR SERVICE [Engine coolant levels]</b>:"+ lastRowValue[0][8]+"<br><br>"+
// "<b>MINOR SERVICE [Steering Fluid levels]</b>:"+ lastRowValue[0][9]+"<br><br>"+
// "<b>MINOR SERVICE [Washer fluid levels]</b>:"+ lastRowValue[0][10]+"<br><br>"+
// "<b>MINOR SERVICE [Wiper blade check-up]</b>:"+ lastRowValue[0][11]+"<br><br>"+
// "<b>Fire Extinguisher availability</b>:"+ lastRowValue[0][12]+"<br><br>"+
// "<b>Fire Extinguisher Expiry Date</b>:"+ lastRowValue[0][13]+"<br><br>"+
// "<b>Battery Health</b>:"+ lastRowValue[0][14]+"<br><br>"+
// "<b>Battery Make</b>:"+ lastRowValue[0][15]+"<br><br>"+
// "<b>Front Right Tyre PSI</b>:"+ lastRowValue[0][16]+"<br><br>"+
// "<b>Front Left Tyre PSI</b>:"+ lastRowValue[0][17]+"<br><br>"+
// "<b>Back Right Tyre PSI</b>:"+ lastRowValue[0][18]+"<br><br>"+
// "<b>Back Left Tyre PSI</b>:"+ lastRowValue[0][19]+"<br><br>"+
// "<b>Tyre Health Inspection [Front - Right]</b>:"+ lastRowValue[0][20]+"<br><br>"+
// "<b>Tyre Health Inspection [Front - Left]</b>:"+ lastRowValue[0][21]+"<br><br>"+
// "<b>Tyre Health Inspection [Rear - Right]</b>:"+ lastRowValue[0][22]+"<br><br>"+
// "<b>Tyre Health Inspection [Rear - Left]</b>:"+ lastRowValue[0][23]+"<br><br>"+
// "<b>Tyre Health comments</b>:"+ lastRowValue[0][24]+"<br><br>"+
// "<b>Brake Pad Inspection [Front - Right]</b>:"+ lastRowValue[0][25]+"<br><br>"+
// "<b>Brake Pad Inspection [Front - Left]</b>:"+ lastRowValue[0][26]+"<br><br>"+
// "<b>Brake Pad Inspection [Rear - Right]</b>:"+ lastRowValue[0][27]+"<br><br>"+
// "<b>Brake Pad Inspection [Rear - Left]</b>:"+ lastRowValue[0][28]+"<br><br>"+
// "<b>Brake Fluid Check-up</b>:"+ lastRowValue[0][29]+"<br><br>"+
// "<b>Air Filter Check</b>:"+ lastRowValue[0][30]+"<br><br>"+
// "<b>Lights Inspection [Front - Right]</b>:"+ lastRowValue[0][31]+"<br><br>"+
// "<b>Lights Inspection [Front - Left]</b>:"+ lastRowValue[0][32]+"<br><br>"+
// "<b>Lights Inspection [Rear - Right]</b>:"+ lastRowValue[0][33]+"<br><br>"+
// "<b>Lights Inspection [Rear - Left]</b>:"+ lastRowValue[0][34]+"<br><br>"+
// "<b>Lights Inspection [Right Brake Light]</b>:"+ lastRowValue[0][35]+"<br><br>"+
// "<b>Lights Inspection [Left Brake Light]</b>:"+ lastRowValue[0][36]+"<br><br>"+
// "<b>General Feedback</b>:"+ lastRowValue[0][37]+"<br><br>"+
// "<b>VAS OrderID</b>:"+ lastRowValue[0][38]+"<br><br>"+
// "<b>Customer ID</b>:"+ lastRowValue[0][39]+"<br><br>"+
// "<b>Customer Name</b>:"+ lastRowValue[0][40]+"<br><br>"+
// "<b>Customer Email</b>:"+ lastRowValue[0][41]+"<br><br>"+
// "<b>Customer Phone Number (without dialcode should start without 0)</b>:"+ lastRowValue[0][42]+"<br><br>"+
// "<b>Vehicle Make</b>:"+ lastRowValue[0][43]+"<br><br>"+
// "<b>Vehicle Model</b>:"+ lastRowValue[0][44]+"<br><br>"+
// "<b>Vehicle Year</b>:"+ lastRowValue[0][45]+"<br><br>"+
// "<b>Vehicle Number Plate</b>:"+ lastRowValue[0][46]+"<br><br>"+
// "<b>Vehicle Number Plate Emirati</b>:"+ lastRowValue[0][47]+"<br><br>"+
// "<b>Timeslot</b>:"+ lastRowValue[0][48]+"<br><br>";


var EmailBody = "<b>Timestamp</b>:"+ lastRowValue[0][0]+"<br><br>"+
"<b>Order ID</b>:"+ lastRowValue[0][1]+"<br><br>"+
"<b>VIN Number</b>:"+ lastRowValue[0][2]+"<br><br>"+
"<b>Pilot Name</b>:"+ lastRowValue[0][3]+"<br><br>"+
"<b>Mileage (Kms)</b>:"+ lastRowValue[0][4]+"<br><br>"+
"<b>Next service Mileage (Kms)</b>:"+ lastRowValue[0][5]+"<br><br>"+
"<b>MINOR SERVICE [Oil Filter check-up]</b>:"+ lastRowValue[0][6]+"<br><br>"+
"<b>MINOR SERVICE [Oil grade check-up]</b>:"+ lastRowValue[0][7]+"<br><br>"+
"<b>MINOR SERVICE [Engine coolant levels]</b>:"+ lastRowValue[0][8]+"<br><br>"+
"<b>MINOR SERVICE [Steering Fluid levels]</b>:"+ lastRowValue[0][9]+"<br><br>"+
"<b>MINOR SERVICE [Washer fluid levels]</b>:"+ lastRowValue[0][10]+"<br><br>"+
"<b>MINOR SERVICE [Wiper blade check-up]</b>:"+ lastRowValue[0][11]+"<br><br>"+
"<b>Fire Extinguisher availability</b>:"+ lastRowValue[0][12]+"<br><br>"+
"<b>Fire Extinguisher Expiry Date</b>:"+ lastRowValue[0][13]+"<br><br>"+
"<b>Battery Health Percentage</b>:"+ lastRowValue[0][14]+"<br><br>"+
"<b>Battery Make</b>:"+ lastRowValue[0][15]+"<br><br>"+
"<b>Front Right Tyre PSI</b>:"+ lastRowValue[0][16]+"<br><br>"+
"<b>Front Left Tyre PSI</b>:"+ lastRowValue[0][17]+"<br><br>"+
"<b>Back Right Tyre PSI</b>:"+ lastRowValue[0][18]+"<br><br>"+
"<b>Back Left Tyre PSI</b>:"+ lastRowValue[0][19]+"<br><br>"+
"<b>Tyre Health Inspection [Front - Right]</b>:"+ lastRowValue[0][20]+"<br><br>"+
"<b>Tyre Health Inspection [Front - Left]</b>:"+ lastRowValue[0][21]+"<br><br>"+
"<b>Tyre Health Inspection [Rear - Right]</b>:"+ lastRowValue[0][22]+"<br><br>"+
"<b>Tyre Health Inspection [Rear - Left]</b>:"+ lastRowValue[0][23]+"<br><br>"+
"<b>Tyre Health comments</b>:"+ lastRowValue[0][24]+"<br><br>"+
"<b>Brake Pad Inspection [Front - Right]</b>:"+ lastRowValue[0][25]+"<br><br>"+
"<b>Brake Pad Inspection [Front - Left]</b>:"+ lastRowValue[0][26]+"<br><br>"+
"<b>Brake Pad Inspection [Rear - Right]</b>:"+ lastRowValue[0][27]+"<br><br>"+
"<b>Brake Pad Inspection [Rear - Left]</b>:"+ lastRowValue[0][28]+"<br><br>"+
"<b>Brake Fluid Check-up</b>:"+ lastRowValue[0][29]+"<br><br>"+
"<b>Air Filter Check</b>:"+ lastRowValue[0][30]+"<br><br>"+
"<b>Lights Inspection [Front - Right]</b>:"+ lastRowValue[0][31]+"<br><br>"+
"<b>Lights Inspection [Front - Left]</b>:"+ lastRowValue[0][32]+"<br><br>"+
"<b>Lights Inspection [Rear - Right]</b>:"+ lastRowValue[0][33]+"<br><br>"+
"<b>Lights Inspection [Rear - Left]</b>:"+ lastRowValue[0][34]+"<br><br>"+
"<b>Lights Inspection [Right Brake Light]</b>:"+ lastRowValue[0][35]+"<br><br>"+
"<b>Lights Inspection [Left Brake Light]</b>:"+ lastRowValue[0][36]+"<br><br>"+
"<b>Engine Size</b>:"+ lastRowValue[0][37]+"<br><br>"+
"<b>Tyre Front Right (Width/Profile/Rim)</b>:"+ lastRowValue[0][38]+"<br><br>"+
"<b>Tyre Front Left (Width/Profile/Rim)</b>:"+ lastRowValue[0][39]+"<br><br>"+
"<b>Tyre Back Right (Width/Profile/Rim)</b>:"+ lastRowValue[0][40]+"<br><br>"+
"<b>Tyre Back Left (Width/Profile/Rim)</b>:"+ lastRowValue[0][41]+"<br><br>"+
"<b>General Feedback</b>:"+ lastRowValue[0][42]+"<br><br>"+
"<b>VAS OrderID</b>:"+ lastRowValue[0][43]+"<br><br>"+
"<b>Customer ID</b>:"+ lastRowValue[0][44]+"<br><br>"+
"<b>Customer Name</b>:"+ lastRowValue[0][45]+"<br><br>"+
"<b>Customer Email</b>:"+ lastRowValue[0][46]+"<br><br>"+
"<b>Customer Phone Number (without dialcode should start without 0)</b>:"+ lastRowValue[0][47]+"<br><br>"+
"<b>Vehicle Make</b>:"+ lastRowValue[0][48]+"<br><br>"+
"<b>Vehicle Model</b>:"+ lastRowValue[0][49]+"<br><br>"+
"<b>Vehicle Year</b>:"+ lastRowValue[0][50]+"<br><br>"+
"<b>Vehicle Number Plate</b>:"+ lastRowValue[0][51]+"<br><br>"+
"<b>Vehicle Number Plate Emirati</b>:"+ lastRowValue[0][52]+"<br><br>"+
"<b>Timeslot</b>:"+ lastRowValue[0][53]+"<br><br>";
// "<h1><b>CONCLUSION : </h1></b><br><br>"+
// "<b>Mileage Report</b>:"+ lastRowValue[0][54]+"<br><br>"+
// "<b>Tyre Report</b>:"+ lastRowValue[0][55]+"<br><br>"+
// "<b>Tyre PSI report</b>:"+ lastRowValue[0][56]+"<br><br>"+
// "<b>Fire Extinguisher</b>:"+ lastRowValue[0][57]+"<br><br>"+
// "<b>Battery Status</b>:"+ lastRowValue[0][58]+"<br><br>"+
// "<b>Brake Pad Status</b>:"+ lastRowValue[0][59]+"<br><br>";


var mileageReport = lastRowValue[0][54];

var tyreReport=lastRowValue[0][55]
var tyrePSIReport = lastRowValue[0][56];
var fireExtinguisherStatus= lastRowValue[0][57];
var batteryHealthStatus= lastRowValue[0][58];
var brakePadStatus=lastRowValue[0][59];

var testarray = [];


// testarray.push(EmailBody);
var Conclusion = "<h1><b>CONCLUSION : </h1></b><br><br>";
testarray.push(Conclusion);
if(mileageReport!=""&& mileageReport!=null)
{Conclusion= Conclusion+ mileageReport+"<br><br>";

testarray.push(mileageReport+"<br><br>");
}
if(tyreReport!=""&& tyreReport!=null)
{Conclusion= Conclusion+ tyreReport+"<br><br>";

testarray.push(tyreReport+"<br><br>");
}

if(tyrePSIReport!=""&& tyrePSIReport!=null)
{Conclusion= Conclusion+ tyrePSIReport+"<br><br>";

testarray.push(tyrePSIReport+"<br><br>");
}

if(fireExtinguisherStatus!=""&& fireExtinguisherStatus!=null)
{Conclusion= Conclusion+ fireExtinguisherStatus+"<br><br>";
testarray.push(fireExtinguisherStatus+"<br><br>");
}

if(batteryHealthStatus!=""&& batteryHealthStatus!=null)
{Conclusion= Conclusion+ batteryHealthStatus+"<br><br>";

testarray.push(batteryHealthStatus+"<br><br>");
}

if(brakePadStatus!=""&& brakePadStatus!=null)
{Conclusion= Conclusion+ brakePadStatus+"<br><br>";

testarray.push(brakePadStatus)+"<br><br>";
}

// var finalEmailBody = EmailBody+Conclusion;

// Logger.log(finalEmailBody);

var abcde = testarray.toString().replace(","," ");
// Logger.log(testarray);
var abc = EmailBody+Conclusion;

Logger.log(abc);
// Logger.log("last :" +lastRowValue[0][59]);

 var email = lastRowValue[0][46];
// Logger.log(EmailBody);
// For Multiple recipients
// var bcc1and2 = "xyz@example.com,abc@example.com";
 Utilities.sleep(2000);
    GmailApp.sendEmail(email,subject,"",{'bcc':slackEmail, htmlBody:abc});

SpreadsheetApp.flush();
}
