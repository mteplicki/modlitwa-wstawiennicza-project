import Utils from "./modules/utils";
import Variables from "./modules/variables";
import UIOperations from "./modules/ui_operations";
import SheetOperations from "./modules/sheet_operations";

function showUsersSidebar(): void {
    try {
        let ss = SpreadsheetApp.getActiveSpreadsheet();
        if (ss.getActiveSheet().getName() !== "Uczestnicy") {
            throw new Error("Przełącz się na arkusz 'Uczestnicy'");
        }
        let widget = HtmlService.createHtmlOutputFromFile(
            "src/templates/UsersSidebar",
        ).setTitle("Modlitwa wstawiennicza MOST");
        SpreadsheetApp.getUi().showSidebar(widget);
    } catch (e: any) {
        if (e instanceof Error) {
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + e.message);
        } else {
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + String(e));
        }
    }
}

function addUser() {
    try {
        let ss = SpreadsheetApp.getActiveSpreadsheet();
        if (ss.getActiveSheet().getName() !== "Uczestnicy") {
            throw new Error("Przełącz się na arkusz 'Uczestnicy'");
        }
        let widget = HtmlService.createHtmlOutputFromFile(
            "src/templates/AddUserSidebar",
        ).setTitle("Modlitwa wstawiennicza MOST");
        SpreadsheetApp.getUi().showSidebar(widget);
    } catch (e: any) {
        if (e instanceof Error) {
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + e.message);
        } else {
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + String(e));
        }
    }
}

function addUserCallback(name : string, email : string, checked : boolean){
    try {
        let ss = SpreadsheetApp.getActiveSpreadsheet();
        if (ss.getActiveSheet().getName() !== "Uczestnicy") {
            throw new Error("Przełącz się na arkusz 'Uczestnicy'");
        }
        let sheet = ss.getSheetByName("Users");
        if (!sheet) {
            throw new Error("Nie znaleziono arkusza o nazwie Users");
        }
        sheet.insertRowBefore(3)
        let range = sheet.getRange("A3:B3");
        range.setValues([[name, email]]);
        if (checked) {
            let invitationText = Variables.getVariable("invitation_text");
            let template = HtmlService.createTemplateFromFile("src/templates/EmailTemplate");
            template.notTable = true
            template.text = invitationText;
            let html = template.evaluate().getContent();
            MailApp.sendEmail({
                to: email,
                subject: "[Modlitwa wstawiennicza MOST] Zaproszenie do modlitwy",
                htmlBody: html,
                name: "Modlitwa wstawiennicza MOST",
            });
        }
        UIOperations.showDialog("Sukces", null, "Użytkownik został dodany");
    } catch (e: any) {
        if (e instanceof Error) {
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + e.message);
        } else {
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + String(e));
        }
    }
}

function removeUser() {
    try {
        let ss = SpreadsheetApp.getActiveSpreadsheet();
        if (ss.getActiveSheet().getName() !== "Uczestnicy") {
            throw new Error("Przełącz się na arkusz 'Uczestnicy'");
        }
        let sheet = ss.getSheetByName("Users");
        if (!sheet) {
            throw new Error("Nie znaleziono arkusza o nazwie Users");
        }
        let range = sheet.getActiveRange();
        if (!range) {
            throw new Error("Zaznacz wiersz do usunięcia");
        }
        let row = range.getRow();
        let name = sheet.getRange(row, 1).getValue();
        let email = sheet.getRange(row, 2).getValue();
        sheet.deleteRow(row);
        UIOperations.showDialog("Sukces", null, "Użytkownik " + name + " został usunięty");
    } catch (e: any) {
        if (e instanceof Error) {
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + e.message);
        } else {
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + String(e));
        }
    }
}

function sendMail() {
    try{
        
    }
}

function sendMailCallback(){

}

