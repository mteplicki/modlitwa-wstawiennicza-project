import Utils from "./modules/utils";
import Variables from "./modules/variables";
import UIOperations from "./modules/ui_operations";
import SheetOperations from "./modules/sheet_operations";
import EmailOperations from "./modules/email_operations";

function showDefaultIntentionsSidebar(): void {
    try {
        Logger.log("showUsersSidebar");
        let [ss, sheet] = Utils.getActiveIntencjeOgolneOrCykliczne();
        let widget = HtmlService.createTemplateFromFile(
            "src/templates/IntentionsSidebar",
        ).evaluate().setTitle("Modlitwa wstawiennicza MOST");
        SpreadsheetApp.getUi().showSidebar(widget);
    } catch (e: any) {
        Utils.handleError(e);
    }
}

function addDefaultIntention() {
    UIOperations.showLoading();
    try {
        Logger.log("addDefaultIntention");
        let [ss, sheet] = Utils.getActiveIntencjeOgolneOrCykliczne();
        let template = HtmlService.createTemplateFromFile(
            "src/templates/AddDefaultIntention",
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
        sheet.insertRowAfter(2)
        let range = sheet.getRange("A3:C3");
        range.clearFormat();
        range.setFontSize(11);
        range.setVerticalAlignment("middle");
        range.setValues([[Utilities.getUuid(), name, intention]]);
        UIOperations.showDialog('Sukces', null, null, `Dodano intencję do arkusza ${sheetName}`)
    } catch (e: any) {
        Utils.handleError(e);
    }
}

function removeDefaultIntention() {
    UIOperations.showLoading();
    try {
        Logger.log("removeDefaultIntention");
        let [ss, sheet] = Utils.getActiveIntencjeOgolneOrCykliczne();
        let currentRange = sheet.getActiveRange();
        if (currentRange === null) {
            throw new Error("Nie wybrano zakresu. Wybierz intencję do usunięcia.");
        }
        if (currentRange.getRow() < 3) {
            throw new Error("Nie można usunąć nagłówka.");
        }
        if (currentRange.getNumRows() !== 1) {
            throw new Error("Wybierz tylko jeden wiersz.");
        }
        Logger.log(currentRange.getRow())
        let row = currentRange.getRow()
        Logger.log(row)
        sheet.deleteRow(row);
        Logger.log("sukces!")
        UIOperations.showDialog('Sukces', null, null, `Usunięto intencję z arkusza ${sheet.getName()}`)
    } catch (e: any) {
        Utils.handleError(e);
    }
}