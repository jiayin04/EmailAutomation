## Project-Based Learning
Contains all the test and trial project I carried out for comprehensive learning and strengthening. 


# Email Automation
In administrative task, we often meet the situation of sending multiple emails with same subject and contents which are burdening the workloads of the manpower in focusing on more high-level execution. Especially when the email sending involves more than 100 people in groups, thus, I have tried and tested out the execution of email sending with the aid of App Script. 

**a) Step 1: Create a Google Sheet/ Submit Form(Link to Google Sheet)**

![image](https://github.com/jiayin04/DummyProject/assets/154343987/2e98b9d5-783e-404e-9b20-dcba6d280d20 | width=250, height=250)

> Place all the necessary attributes in the excel sheet if you are creating google sheet
> Sample of Google Sheet Template:

![image](https://github.com/jiayin04/DummyProject/assets/154343987/c72263d3-fd5c-4ba5-96ff-b45b49d67ff0)


**b) Step 2: Go to Extension and Open App Script**

![image](https://github.com/jiayin04/DummyProject/assets/154343987/1e95bdc5-9a33-47f1-90bf-8ceceb915165)

> In app script, we will start to code out our desired function as we can see over the menu
> Customized menu:

```function onOpen() {


  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Send Membership Email")
    .addItem("Send Email Now", "main") //main is to link back to the main function
    .addToUi();


};
```

**c) Step 3: Code the Main Function**
- In email automation, we have few important functions, like email sending function, email validation function, menu customization function and main function which execute the entire script.

> Variable Declaration:

```var ss_id = "13GiYvizQhhLUvpXtFdJlql6BFUz7HdyZhf1FJFDXD2A";//Sheet ID
var sheetName = "Automate Email"; // sheet name
var startRow = 2; //First row after your header.


//Email data
var subject = "Welcome to Our Club + Join Our Discord Server"; //The email subject
```

> Send Email Function:

```function sendEmail(membership, sheetURL) {

  //Members Details
  var members = {
    "name": membership[1],
    "email": membership[6],
  };

  if (!isValidEmail(members.email)) { //Call Email Validation function
    return members.email;
  }

  var body = HtmlService.createTemplateFromFile("email"); //Include Email HTML file
  body.name = members.name;

  try {
    MailApp.sendEmail({
      to: members.email,
      subject: subject,
      htmlBody: body.evaluate().getContent(),
    });
  }
  catch (error) {
    console.log(error);
  }
};
```
* For avoiding email sending error, we may add function to check the validity of the email in case the email is not accurate or invalid.
  
  > Email Validation:
  
  ```function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
  }; 
  ```

> Main function (Calling sendEmail function):

```function main() {
  var ss = SpreadsheetApp.openById(ss_id);
  var sheet = ss.getSheetByName(sheetName);

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  var range = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol);
  var rangeVals = range.getValues();

  var badEmailList = [];

  //Loop through range values.
  for (var row = 0; row < rangeVals.length; row++) {

    if (!rangeVals[row][14]) {
      var badEmail = sendEmail(rangeVals[row], ss.getUrl);

      var date = new Date();
      var updatedDate = Utilities.formatDate(date, "GMT+08", "yyyy-MM-dd HH:mm:ss");
      rangeVals[row][15] = updatedDate; //UPDATE THE SEND TIME(O)
      sheet.getRange(row + startRow, 15).setValue(updatedDate);

      if (badEmail) {
        badEmailList.push(badEmail);
        sheet.getRange(row + startRow, 16).setValue("Email Not Send!");

      }
      else {

        // Add checkboxes for the newly submitted rows
        sheet.getRange(row + startRow, 14).insertCheckboxes().setValue(true);
        sheet.getRange(row + startRow, 16).setValue("-");
      };
    }
    else {
      continue;
    };


  };
};
```

**d) Step 4: Customize Own Email Format**
> Open another file named it as Email, in the file, we will use HTML and CSS for creating the basic framework and style of the email.

```<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <base target="_top">

<style>
    .header{
        background-color:rgb(246, 210, 215); 
        color:black; 
        height: 50px; 
        width: auto;
        text-align:center;
        padding-top: 20px;
    }

</style>
</head>

<body>
  <div style="margin:3%; border:2px solid rgb(162, 159, 159);">
    <div style="margin:3%;">
      <h2 class="header">
        OSC Membership Confirmation
      </h2>
    </div>

    <div style="margin:3%; text-align: justify;">
      <p>Hello
        <?= name ?>,
      </p>

      <p>
        Warm greetings from XXX Club! We are glad to receive your registration of our club's membership.
        With that being said, welcome aboard to XXX! <br> <br>

        The link below is the link to join our Discord server to join our weekly sessions and interactive activities: <br><br>
        <a href="https://discord.com/"> Discord Server Link </a> <br> <br>

        For any further enquiries or more information, feel free to contact us by replying to this email or through our
        Instagram account @XXX!
        Thank you and have a nice day! <br> <br>
      </p>

      <p>
        Best regards, <br>
        XXX Club 
      </p>
    </div>
  </div>
</body>

</html>
```

**Step 5: Automate the Email with App Script**
- Go to 'Trigger'

![image](https://github.com/jiayin04/DummyProject/assets/154343987/41f71929-1ba2-4dd4-a190-03558a6e4a41) 

- Make changes based on your need (Time Driven, Spreadsheet, or Calendar)

![image](https://github.com/jiayin04/DummyProject/assets/154343987/58121897-8ede-48b7-9705-55e521f6fd78)

~ DONE!!!!! We made it ^_^

