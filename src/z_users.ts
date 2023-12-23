import Utils from "./modules/utils";
import Variables from "./modules/variables";
import UIOperations from "./modules/ui_operations";
import SheetOperations from "./modules/sheet_operations";

function showUsersSidebar() : void {
    try {
        let widget = HtmlService.createHtmlOutputFromFile('src/templates/UsersSidebar').setTitle('Modlitwa wstawiennicza MOST')
        SpreadsheetApp.getUi().showSidebar(widget);
    } catch (e: any) {
        if (e instanceof Error) {
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + e.message)
        } else {
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + String(e))
        }
    }
}

function addUser() {

}

function removeUser() {

}

function sendMail() {

}
