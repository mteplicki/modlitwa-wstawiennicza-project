import Utils from "./modules/utils";
import Variables from "./modules/variables";
import UIOperations from "./modules/ui_operations";
import EmailOperations from "./modules/email_operations";

function showUsersSidebar(): void {
    try {
        Logger.log("showUsersSidebar");
        let ss = SpreadsheetApp.getActiveSpreadsheet();
        if (ss.getActiveSheet().getName() !== "Uczestnicy") {
            throw new Error("Przełącz się na arkusz 'Uczestnicy'");
        }
        let widget = HtmlService.createHtmlOutputFromFile(
            "src/templates/UsersSidebar",
        ).setTitle("Modlitwa wstawiennicza MOST");
        SpreadsheetApp.getUi().showSidebar(widget);
    } catch (e: any) {
        Utils.handleError(e);
    }
}

function addUser() {
    try {
        Logger.log("addUser");
        let ss = SpreadsheetApp.getActiveSpreadsheet();
        if (ss.getActiveSheet().getName() !== "Uczestnicy") {
            throw new Error("Przełącz się na arkusz 'Uczestnicy'");
        }
        let widget = HtmlService.createHtmlOutputFromFile(
            "src/templates/AddUser",
        ).setTitle("Modlitwa wstawiennicza MOST");
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
        UIOperations.showDialog("Sukces", null, null, "Użytkownik został dodany");
    } catch (e: any) {
        Utils.handleError(e);
    }
}

function removeUser() {
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
        UIOperations.showDialog("Sukces", null, null, "Użytkownik " + name + " został usunięty");
    } catch (e: any) {
        Utils.handleError(e);
    }
}

function sendMail() {
    try{
        Logger.log("sendMail");
        let [, sheet] = Utils.getActiveSheetByName("Uczestnicy");
        let range = sheet.getActiveRange();
        if (!range) {
            throw new Error("Zaznacz wiersz do usunięcia");
        }
        let row = range.getRow();
        let name = sheet.getRange(row, 1).getValue();
        let email = sheet.getRange(row, 2).getValue();

        let template = HtmlService.createTemplateFromFile("src/templates/EmailSender");
        template.email = email;
        let html = template.evaluate();
        let widget = HtmlService.createHtmlOutput(html).setWidth(450).setHeight(300);
        SpreadsheetApp.getUi().showModalDialog(widget, "Wyślij wiadomość");
    } catch (e: any) {
        Utils.handleError(e);
    }
}

function sendMailCallback(email : string, subject : string, text : string){
    try {
        Logger.log("sendMailCallback");
        EmailOperations.sendEmail({to: email, subject: subject, text: text});
        UIOperations.showDialog("Sukces", null, null, "Wiadomość została wysłana");
    } catch (e: any) {
        Utils.handleError(e);
    }
}

