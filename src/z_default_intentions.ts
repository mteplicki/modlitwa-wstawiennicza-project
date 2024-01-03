import Utils from "./modules/utils";
import Variables from "./modules/variables";
import UIOperations from "./modules/ui_operations";
import SheetOperations from "./modules/sheet_operations";
import EmailOperations from "./modules/email_operations";

function showDefaultIntentionsSidebar(): void {
    try {
        Logger.log("showUsersSidebar");
        let [ss, sheet] = Utils.getActiveIntencjeOgólneOrCykliczne();
        let widget = HtmlService.createHtmlOutputFromFile(
            "src/templates/IntentionsSidebar",
        ).setTitle("Modlitwa wstawiennicza MOST");
        SpreadsheetApp.getUi().showSidebar(widget);
    } catch (e: any) {
        Utils.handleError(e);
    }
}

function addDefaultIntention() {
    try {
        Logger.log("addDefaultIntention");
        let [ss, sheet] = Utils.getActiveIntencjeOgólneOrCykliczne();
        let template = HtmlService.createTemplateFromFile(
            "src/templates/AddIntention",
        )
        template.sheet = sheet.getName()
        let sheetName = sheet.getName() === "Intencje-ogólne" ? "intencję ogólną" : "intencję cykliczną"
        template.sheetName = sheetName
        let widget = template.evaluate().setWidth(400).setHeight(300);
        SpreadsheetApp.getUi().showModalDialog(widget, `Dodaj ${sheetName}`);
    } catch (e: any) {
        Utils.handleError(e);
    }
}

function addDefaultIntentionCallback(sheetName : string, name : string, intention : string){
    try {
        Logger.log("addDefaultIntentionCallback");
        let [ss, sheet] = Utils.getActiveSheetByName(sheetName);
        sheet.insertRowBefore(3);
        let range = sheet.getRange("A3:C3");
        range.setValues([[Utilities.getUuid(), name, intention]]);
        UIOperations.showDialog('Sukces', null, null, `Dodano intencję do arkusza ${sheetName}`)
    } catch (e: any) {
        Utils.handleError(e);
    }
}

function removeDefaultIntention() {
    try {
        Logger.log("removeDefaultIntention");
        let [ss, sheet] = Utils.getActiveIntencjeOgólneOrCykliczne();
        let currentRange = sheet.getActiveRange();
        if (currentRange === null) {
            throw new Error("Nie wybrano zakresu. Wybierz intencję do usunięcia.");
        }
        let currentRow = currentRange.getRow();
        sheet.deleteRow(currentRow);
        UIOperations.showDialog('Sukces', null, null, `Usunięto intencję z arkusza ${sheet.getName()}`)
    } catch (e: any) {
        Utils.handleError(e);
    }
}