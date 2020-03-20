ss = SpreadsheetApp.getActiveSpreadsheet();
introcallNotesSheet = ss.getSheetByName("Introcall Notes");
settingsSheet = ss.getSheetByName("Settings");
historyLogSheet = ss.getSheetByName("History log");

callOwner = introcallNotesSheet.getRange("B1").getValue();

companyName = introcallNotesSheet.getRange("B4").getValue();
country = introcallNotesSheet.getRange("B5").getValue();
companyType = introcallNotesSheet.getRange("B6").getValue();

callType = introcallNotesSheet.getRange("B8").getValue();

attendeeName1 = introcallNotesSheet.getRange("B9").getValue();
attendeeRole1 = introcallNotesSheet.getRange("B10").getValue();
attendeeEmail1 = introcallNotesSheet.getRange("B11").getValue();
attendeeName2 = introcallNotesSheet.getRange("B12").getValue();
attendeeRole2 = introcallNotesSheet.getRange("B13").getValue();
attendeeEmail2 = introcallNotesSheet.getRange("B14").getValue();
attendeeName3 = introcallNotesSheet.getRange("B15").getValue();
attendeeRole3 = introcallNotesSheet.getRange("B16").getValue();
attendeeEmail3 = introcallNotesSheet.getRange("B17").getValue();
attendeeName4 = introcallNotesSheet.getRange("B18").getValue();
attendeeRole4 = introcallNotesSheet.getRange("B19").getValue();
attendeeEmail4 = introcallNotesSheet.getRange("B20").getValue();

projectType = introcallNotesSheet.getRange("E2").getValue();
introcallNotesSheet.getRange("E3").setNumberFormat('@STRING@');
introcallNotesSheet.getRange("E5").setNumberFormat('@STRING@');
numberUsers = introcallNotesSheet.getRange("E3").getValue();
framework = introcallNotesSheet.getRange("E4").getValue();
numberLanguages = introcallNotesSheet.getRange("E5").getValue();

timeline = new Date(introcallNotesSheet.getRange("E6").getValue());

if (timeline != "") {

    Utilities.formatDate(timeline, "Europe/Paris", "dd/MM/yyyy");
};

liveWithPrismic = introcallNotesSheet.getRange("E7").getValue();
workWithAgency = introcallNotesSheet.getRange("E8").getValue();
heardAboutPrismic = introcallNotesSheet.getRange("E9").getValue();

createdRepo = introcallNotesSheet.getRange("E11").getValue();
currentCMS = introcallNotesSheet.getRange("E12").getValue();

currentWorkflow = introcallNotesSheet.getRange("E13:E16").getValues().reduce(function(ar, e) {
  if (e[0]) ar.push(e[0])
  return ar;
},[]);

whyPrismic = introcallNotesSheet.getRange("E17:E20").getValues().reduce(function (ar, e) {
    if (e[0]) ar.push(e[0])
    return ar;
}, []);

valuableInfos = [
    introcallNotesSheet.getRange("G2").getValue(),
    introcallNotesSheet.getRange("G3").getValue(),
    introcallNotesSheet.getRange("G4").getValue(),
    introcallNotesSheet.getRange("G5").getValue(),
    introcallNotesSheet.getRange("G6").getValue(),
    introcallNotesSheet.getRange("G7").getValue(),
    introcallNotesSheet.getRange("G8").getValue(),
    introcallNotesSheet.getRange("G9").getValue(),
    introcallNotesSheet.getRange("G10").getValue(),
    introcallNotesSheet.getRange("H2").getValue(),
    introcallNotesSheet.getRange("H3").getValue(),
    introcallNotesSheet.getRange("H4").getValue(),
    introcallNotesSheet.getRange("H5").getValue(),
    introcallNotesSheet.getRange("H6").getValue(),
    introcallNotesSheet.getRange("H7").getValue(),
    introcallNotesSheet.getRange("H8").getValue(),
    introcallNotesSheet.getRange("H9").getValue(),
    introcallNotesSheet.getRange("H10").getValue()
];
keyFeatures = [
    introcallNotesSheet.getRange("G12").getValue(),
    introcallNotesSheet.getRange("G13").getValue(),
    introcallNotesSheet.getRange("G14").getValue(),
    introcallNotesSheet.getRange("H12").getValue(),
    introcallNotesSheet.getRange("H13").getValue(),
    introcallNotesSheet.getRange("H14").getValue()
];
dealAlerts = [
    introcallNotesSheet.getRange("G16").getValue(),
    introcallNotesSheet.getRange("G17").getValue(),
    introcallNotesSheet.getRange("G18").getValue(),
    introcallNotesSheet.getRange("G19").getValue(),
    introcallNotesSheet.getRange("G20").getValue()
];
featureRequests = [
    introcallNotesSheet.getRange("H16").getValue(),
    introcallNotesSheet.getRange("H17").getValue(),
    introcallNotesSheet.getRange("H18").getValue(),
    introcallNotesSheet.getRange("H19").getValue(),
    introcallNotesSheet.getRange("H20").getValue()
];
nextSteps = [
    introcallNotesSheet.getRange("J2").getValue(),
    introcallNotesSheet.getRange("J3").getValue(),
    introcallNotesSheet.getRange("J4").getValue(),
    introcallNotesSheet.getRange("J5").getValue(),
    introcallNotesSheet.getRange("J6").getValue(),
    introcallNotesSheet.getRange("J7").getValue()
];

sentToSales = introcallNotesSheet.getRange("J8").getValue();

docsToShare = [
    introcallNotesSheet.getRange("J10").getValue(),
    introcallNotesSheet.getRange("J11").getValue(),
    introcallNotesSheet.getRange("J12").getValue(),
    introcallNotesSheet.getRange("J13").getValue(),
    introcallNotesSheet.getRange("J14").getValue(),
    introcallNotesSheet.getRange("J15").getValue(),
    introcallNotesSheet.getRange("J16").getValue(),
    introcallNotesSheet.getRange("J17").getValue(),
    introcallNotesSheet.getRange("J18").getValue(),
    introcallNotesSheet.getRange("J19").getValue(),
    introcallNotesSheet.getRange("J20").getValue()
];
allNotes =
    companyName + " (" + country + " - " + companyType + ") Introcall Notes:\n" +
    /**/
    introcallNotesSheet.getRange("A1").getValue() + ": " + callOwner + "\n" +
    introcallNotesSheet.getRange("A8").getValue() + ": " + callType + "\n\n" +
    /**/
    "Attendees:\n" +
    attendeeName1 + " (" + attendeeRole1 + " - " + attendeeEmail1 + ")\n" +
    attendeeName2 + " (" + attendeeRole2 + " - " + attendeeEmail2 + ")\n" +
    attendeeName3 + " (" + attendeeRole3 + " - " + attendeeEmail3 + ")\n" +
    attendeeName4 + " (" + attendeeRole4 + " - " + attendeeEmail4 + ")\n\n" +
    /**/
    introcallNotesSheet.getRange("D2").getValue() + ": " + projectType + "\n" +
    introcallNotesSheet.getRange("D3").getValue() + ": " + numberUsers + "\n" +
    introcallNotesSheet.getRange("D4").getValue() + ": " + framework + "\n" +
    introcallNotesSheet.getRange("D5").getValue() + ": " + numberLanguages + "\n" +
    introcallNotesSheet.getRange("D6").getValue() + ": " + timeline + "\n" +
    introcallNotesSheet.getRange("D7").getValue() + ": " + liveWithPrismic + "\n" +
    introcallNotesSheet.getRange("D8").getValue() + ": " + workWithAgency + "\n" +
    introcallNotesSheet.getRange("D9").getValue() + ": " + heardAboutPrismic + "\n\n" +
    /**/
    "PIPE ASSIGNMENT: " + sentToSales + "\n\n" +
    /**/
    introcallNotesSheet.getRange("D11").getValue() + ": " + createdRepo + "\n" +
    introcallNotesSheet.getRange("D12").getValue() + ": " + currentCMS + "\n\n" +
    /**/
    introcallNotesSheet.getRange("D13").getValue() + ":\n" +
    currentWorkflow + "\n\n" +
    /**/
    introcallNotesSheet.getRange("D17").getValue() + ":\n" +
    whyPrismic + "\n\n" +
    /**/
    introcallNotesSheet.getRange("G1").getValue() + ":\n" +
    "- " + valuableInfos[0] + "\n- " + valuableInfos[1] + "\n- " + valuableInfos[2] + "\n- " + valuableInfos[3] + "\n- " + valuableInfos[4] + "\n- " + valuableInfos[5] + "\n- " + valuableInfos[6] + "\n- "
    + valuableInfos[7] + "\n- " + valuableInfos[8] + "\n- " + valuableInfos[9] + "\n- " + valuableInfos[10] + "\n- " + valuableInfos[11] + "\n- " + valuableInfos[12] + "\n- " + valuableInfos[13] + "\n- " + valuableInfos[14] + "\n- " + valuableInfos[15] + "\n- " + valuableInfos[16] + "\n- " + valuableInfos[17] + "\n\n" +
    /**/
    introcallNotesSheet.getRange("G11").getValue() + ":\n" +
    "- " + keyFeatures[0] + "\n- " + keyFeatures[1] + "\n- " + keyFeatures[2] + "\n- " + keyFeatures[3] + "\n- " + keyFeatures[4] + "\n- " + keyFeatures[5] + "\n\n" +
    /**/
    introcallNotesSheet.getRange("G15").getValue() + ":\n" +
    "- " + dealAlerts[0] + "\n" +
    "- " + dealAlerts[1] + "\n" +
    "- " + dealAlerts[2] + "\n" +
    "- " + dealAlerts[3] + "\n" +
    "- " + dealAlerts[4] + "\n\n" +
    /**/
    introcallNotesSheet.getRange("H15").getValue() + ":\n" +
    "- " + featureRequests[0] + "\n" +
    "- " + featureRequests[1] + "\n" +
    "- " + featureRequests[2] + "\n" +
    "- " + featureRequests[3] + "\n" +
    "- " + featureRequests[4] + "\n\n" +
    /**/
    introcallNotesSheet.getRange("J1").getValue() + ":\n" +
    "- " + nextSteps[0] + "\n" +
    "- " + nextSteps[1] + "\n" +
    "- " + nextSteps[2] + "\n" +
    "- " + nextSteps[3] + "\n" +
    "- " + nextSteps[4] + "\n" +
    "- " + nextSteps[5] + "\n\n" +
    /**/
    introcallNotesSheet.getRange("J9").getValue() + ":\n" +
    "- " + docsToShare[0] + "\n" +
    "- " + docsToShare[1] + "\n" +
    "- " + docsToShare[2] + "\n" +
    "- " + docsToShare[3] + "\n" +
    "- " + docsToShare[4] + "\n" +
    "- " + docsToShare[5] + "\n" +
    "- " + docsToShare[6] + "\n" +
    "- " + docsToShare[7] + "\n" +
    "- " + docsToShare[8] + "\n" +
    "- " + docsToShare[9]
    ;


function onOpen(e) {
    introcallNotesSheet.getRange("E3").setNumberFormat('@STRING@');
    introcallNotesSheet.getRange("E5").setNumberFormat('@STRING@');
    SpreadsheetApp.getUi()
        .createMenu("Controls")
        .addItem("Generate Notes", "generateNotes")
        .addItem("Save Infos", "saveInfos")
        .addItem("Reset Notes", "resetNotes")
        .addToUi();
}


function generateNotes() {
    introcallNotesSheet.getRange("E3").setNumberFormat('@STRING@');
    introcallNotesSheet.getRange("E5").setNumberFormat('@STRING@');
    if (sentToSales == "-= SELECT VALUE =-") {
        SpreadsheetApp.getUi().alert("Please select a pipe (cell J8)");
    }

    else {
        SpreadsheetApp.getUi().alert(allNotes);
        saveInfos();
    }

}

function saveInfos() {
    introcallNotesSheet.getRange("E3").setNumberFormat('@STRING@');
    introcallNotesSheet.getRange("E5").setNumberFormat('@STRING@');
    var ls = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1GF_-Z_qgdDahAy1htljOoL1ZRgANwvjq_vXoN2tR9rw/edit");
    var historyLogSheetExternal = ls.getSheetByName("History Log V2");

    if (sentToSales == "-= SELECT VALUE =-") {
        SpreadsheetApp.getUi().alert("Please select a pipe (cell J8)");
    }

    else {
        historyLogSheetExternal.appendRow([
            Utilities.formatDate(new Date(), "Europe/Paris", "dd/MM/yyyy - HH:mm"),
            callOwner,
            companyName,
            country,
            companyType,
            callType,
            attendeeName1,
            attendeeRole1,
            attendeeEmail1.trim(),
            attendeeName2,
            attendeeRole2,
            attendeeEmail2.trim(),
            attendeeName3,
            attendeeRole3,
            attendeeEmail3.trim(),
            attendeeName4,
            attendeeRole4,
            attendeeEmail4.trim(),
            projectType,
            numberUsers,
            framework,
            numberLanguages,
            timeline,
            liveWithPrismic,
            workWithAgency,
            heardAboutPrismic,
            createdRepo,
            currentCMS,
            currentWorkflow.toString(),
            whyPrismic.toString(),
            valuableInfos.toString(),
            keyFeatures.toString(),
            dealAlerts.toString(),
            featureRequests.toString(),
            nextSteps.toString(),
            docsToShare.toString(),
            sentToSales.toString()
        ]);
        historyLogSheet.appendRow([
            Utilities.formatDate(new Date(), "Europe/Paris", "dd/MM/yyyy - HH:mm"),
            callOwner,
            companyName,
            country,
            companyType,
            callType,
            attendeeName1,
            attendeeRole1,
            attendeeEmail1.trim(),
            attendeeName2,
            attendeeRole2,
            attendeeEmail2.trim(),
            attendeeName3,
            attendeeRole3,
            attendeeEmail3.trim(),
            attendeeName4,
            attendeeRole4,
            attendeeEmail4.trim(),
            projectType,
            numberUsers,
            framework,
            numberLanguages,
            timeline,
            liveWithPrismic,
            workWithAgency,
            heardAboutPrismic,
            createdRepo,
            currentCMS,
            currentWorkflow.toString(),
            whyPrismic.toString(),
            valuableInfos.toString(),
            keyFeatures.toString(),
            dealAlerts.toString(),
            featureRequests.toString(),
            nextSteps.toString(),
            docsToShare.toString(),
            sentToSales.toString()
        ]);
    }
}

function resetFields() {
    introcallNotesSheet.getRange("E3").setNumberFormat('@STRING@');
    introcallNotesSheet.getRange("E5").setNumberFormat('@STRING@');
    var assignToSalesOptions = settingsSheet.getRange("J2:J3");
    var assignToSales = SpreadsheetApp.newDataValidation().requireValueInRange(assignToSalesOptions).build();

    introcallNotesSheet.getRange("B4:B6").setValue(" ");
    introcallNotesSheet.getRange("B8:B20").setValue(" ");
    introcallNotesSheet.getRange("E2:E9").setValue(" ");
    introcallNotesSheet.getRange("E11:E20").setValue(" ");
    introcallNotesSheet.getRange("G2:H10").setValue(" ");
    introcallNotesSheet.getRange("G12:H14").setValue(" ");
    introcallNotesSheet.getRange("G16:H20").setValue(" ");
    introcallNotesSheet.getRange("J2:J7").setValue(" ");
    introcallNotesSheet.getRange("J8").setDataValidation(assignToSales);
    introcallNotesSheet.getRange("J8").setValue("-= SELECT VALUE =-");
    introcallNotesSheet.getRange("J10:J20").setValue(" ");
}

function resetNotes() {
    var response = SpreadsheetApp.getUi().alert('You are about to reset these values.', 'Do you want to save before erasing?', SpreadsheetApp.getUi().ButtonSet.YES_NO);
    if (response == SpreadsheetApp.getUi().Button.YES) {
        saveInfos();
        resetFields();
    }

    else {
        resetFields();
    }
}

