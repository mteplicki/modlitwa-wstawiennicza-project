import Utils from "./modules/utils";
import Variables from "./modules/variables";
import UIOperations from "./modules/ui_operations";
import EmailOperations from "./modules/email_operations";
import FirebaseInit from "./modules/firebase_init";

function showUsersSidebar(): void {
    try {
        Logger.log("showUsersSidebar");
        let ss = SpreadsheetApp.getActiveSpreadsheet();
        if (ss.getActiveSheet().getName() !== "Uczestnicy") {
            throw new Error("Przełącz się na arkusz 'Uczestnicy'");
        }
        let widget = HtmlService.createTemplateFromFile(
            "src/templates/UsersSidebar",
        ).evaluate().setTitle("Modlitwa wstawiennicza MOST");
        SpreadsheetApp.getUi().showSidebar(widget);
    } catch (e: any) {
        Utils.handleError(e);
    }
}

function addUser() {
    UIOperations.showLoading();
    try {
        Logger.log("addUser");
        let ss = SpreadsheetApp.getActiveSpreadsheet();
        if (ss.getActiveSheet().getName() !== "Uczestnicy") {
            throw new Error("Przełącz się na arkusz 'Uczestnicy'");
        }
        let widget = HtmlService.createTemplateFromFile(
            "src/templates/AddUser",
        ).evaluate().setTitle("Modlitwa wstawiennicza MOST");
        SpreadsheetApp.getUi().showModalDialog(widget, "Dodaj uczestnika");
    } catch (e: any) {
        Utils.handleError(e);
    }
}

function addUserCallback(name : string, email : string, checked : boolean){
    try {
        Logger.log("addUserCallback");
        let [, sheet] = Utils.getActiveSheetByName("Uczestnicy");
        sheet.insertRowBefore(3)
        let range = sheet.getRange("A3:B3");
        range.setValues([[name, email]]);
        if (checked) {
            let invitationText = Variables.getVariable("invitation_text");
            EmailOperations.sendEmail({to: email, subject: "Zaproszenie do modlitwy wstawienniczej MOST", text: invitationText, name: name});
        }
        FirebaseInit.firestore.createDocument(`intentions/${email}`, {name: name, email: email});
        UIOperations.showDialog("Sukces", null, null, "Użytkownik został dodany");
        
    } catch (e: any) {
        Utils.handleError(e);
    }
}

function removeUser() {
    UIOperations.showLoading();
    try {
        Logger.log("removeUser");
        let [, sheet] = Utils.getActiveSheetByName("Uczestnicy");
        let range = sheet.getActiveRange();
        if (!range) {
            throw new Error("Zaznacz wiersz do usunięcia");
        }
        let row = range.getRow();
        if (row < 3) {
            throw new Error("Nie możesz usunąć tego wiersza");
        }
        let name = sheet.getRange(row, 1).getValue();
        let email = sheet.getRange(row, 2).getValue();
        sheet.deleteRow(row);
        FirebaseInit.firestore.deleteDocument(`intentions/${email}`);
        UIOperations.showDialog("Sukces", null, null, "Użytkownik " + name + " został usunięty"); 
    } catch (e: any) {
        Utils.handleError(e);
    }
}

function sendMail() {
    UIOperations.showLoading();
    try{
        Logger.log("sendMail");
        let [, sheet] = Utils.getActiveSheetByName("Uczestnicy");
        let range = sheet.getActiveRange();
        if (!range) {
            throw new Error("Zaznacz wiersz do wysłania");
        }
        if (range.getRow() < 3) {
            throw new Error("Nie możesz zaznaczyć wiersza nagłówka");
        }
        let email = ""
        let name = ""
        for (let i = 0; i < range.getNumRows(); i++) {
            let row = range.getRow() + i;
            email += `${sheet.getRange(row, 2).getValue()},`;
        }
        // remove last comma
        email = email.substring(0, email.length - 1);
        let template = HtmlService.createTemplateFromFile("src/templates/EmailSender");
        template.email = email;
        let html = template.evaluate();
        let widget = HtmlService.createHtmlOutput(html).setWidth(450).setHeight(300);
        SpreadsheetApp.getUi().showModalDialog(widget, "Wyślij wiadomość");
    } catch (e: any) {
        Utils.handleError(e);
    }
}

function sendMailCallback(emails : string, subject : string, text : string){
    try {
        Logger.log("sendMailCallback");
        for (let email of emails.split(",").map((email : string) => email.trim()).filter((email : string) => email !== "")) {
            EmailOperations.sendEmail({to: email, subject: subject, text: text});
        }
        UIOperations.showDialog("Sukces", null, null, "Wiadomości zostały wysłane");
    } catch (e: any) {
        Utils.handleError(e);
    }
}

