import SheetOperations from "./modules/sheet_operations";
import UIOperations from "./modules/ui_operations";
import Utils from "./modules/utils";
import Variables from "./modules/variables";
function onOpen(): void {
    SpreadsheetApp
        .getUi()
        .createMenu("Modlitwa wstawiennicza")
        .addItem("Otwórz panel intencji", "showIntentionSidebar")
        .addItem("Otwórz panel uczestników", "showUsersSidebar")
        .addItem("Otwórz panel domyślnych intencji", "showDefaultIntentionsSidebar")
        .addItem("Dokumentacja", "showDocumentation")
        .addToUi();
}

function doGet(e : any) {
    Logger.log(e);
    const url = "https://docs.google.com/spreadsheets/d/1LJuwTBUkpp_KVYZH89w0e-mU0lPJZ1Eau_z2oIa3oPA/edit"
    let result = UrlFetchApp.fetch(url)
    return HtmlService.createHtmlOutput(result.getContentText())
}

function doPost(e: any) {
    return true;
}

function showDocumentation(){
    UIOperations.openUrl("https://docs.google.com/document/d/1EGnnS1uNXitrftmpcXH-I1Pm-K5oPIm9kV8aHn2SLwM/edit?usp=sharing")
}

function onEditVariables(e : GoogleAppsScript.Events.SheetsOnEdit){
    Logger.log("onEditVariables")
    let [ss, sheet] = Utils.getSheetByName("Ustawienia");
    if (e.range.getColumn() === 3){
        let variable = sheet.getRange(e.range.getRow(), 1).getValue();
        let value = sheet.getRange(e.range.getRow(), 3).getValue();
        CacheService.getScriptCache().put(variable, value);
        Logger.log(`Saved ${variable} with value ${value}`)
    }
    
}

function onEditIntencje (e : GoogleAppsScript.Events.SheetsOnEdit){
    Logger.log("onEditIntencje")
    
}

function onEdit(e :GoogleAppsScript.Events.SheetsOnEdit){
    switch (e.range.getSheet().getName()) {
        case "Ustawienia":
            onEditVariables(e);
            break;
        case "Intencje":
            onEditIntencje(e);
            break;
        default:
            break;
    }
}

