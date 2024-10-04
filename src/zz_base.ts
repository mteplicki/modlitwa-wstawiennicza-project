import SheetOperations from "./modules/sheet_operations";
import UIOperations from "./modules/ui_operations";
import Utils from "./modules/utils";
import FirebaseInit from "./modules/firebase_init";

function testFirebase() {
    let result = FirebaseInit.firestore.createDocument("FirstCollection", {
        "name": "test!"
    });
    Logger.log(result);
}
function refreshVariables() {
    UIOperations.showLoading();
    try {
        Logger.log("refreshVariables");
        let [ss, sheet] = Utils.getSheetByName("Ustawienia");
        let variables = sheet.getRange("A2:C").getValues();
        variables.forEach((variable) => {
            PropertiesService.getDocumentProperties().setProperty(variable[0], variable[2]);
        });
        UIOperations.showDialog("Sukces", null, null, "Zaaktualizowane zmienne");
    } catch (e: any) {
        Utils.handleError(e);
    }
}

function showIntentionSidebar(): void {
    try {
        let widget = HtmlService.createTemplateFromFile(
            "src/templates/AdminSidebar",
        ).evaluate().setTitle("Modlitwa wstawiennicza MOST");
        SpreadsheetApp.getUi().showSidebar(widget);
    } catch (e: any) {
        Utils.handleError(e);
    }
}

function showDocumentation() {
    UIOperations.openUrl("https://docs.google.com/document/d/1EGnnS1uNXitrftmpcXH-I1Pm-K5oPIm9kV8aHn2SLwM/edit?usp=sharing")
}

function onEditVariables(e: GoogleAppsScript.Events.SheetsOnEdit) {
    Logger.log("onEditVariables")
    let [ss, sheet] = Utils.getSheetByName("Ustawienia");
    if (e.range.getColumn() === 3) {
        UIOperations.showLoading();
        let variable = sheet.getRange(e.range.getRow(), 1).getValue();
        let value = sheet.getRange(e.range.getRow(), 3).getValue();
        PropertiesService.getDocumentProperties().setProperty(variable, value);
        Logger.log(`Saved ${variable} with value ${value}`)
        UIOperations.showDialog("Sukces", null, null, "Zaaktualizowano zmienną");
    }
}

function onEditIntencje(e: GoogleAppsScript.Events.SheetsOnEdit) {
    Logger.log("onEditIntencje")
    const [ss, sheet] = Utils.getSheetByName("Intencje");
    if (e.range.getColumn() === 2) {
        let date = sheet.getRange(e.range.getRow(), 2).getValue() as string;
        Logger.log(date);
        let dateObj = new Date(date);
        let dateStr = Utilities.formatDate(dateObj, "Europe/Warsaw", "yyyy-MM-dd HH:mm:ss");
        Logger.log(dateStr);
        sheet.getRange(e.range.getRow(), 2).setValue(dateStr);
        sheet.getRange(e.range.getRow(), 3).setValue(dateStr)
        Logger.log(`Saved ${date} with value ${dateStr}`)
        let [start, end] = SheetOperations.getRangeArray(SheetOperations.getRange())
        let [startObj, endObj] = [new Date(start), new Date(end)]
        if (dateObj.getTime() < startObj.getTime() || dateObj.getTime() > endObj.getTime()) {
            UIOperations.showLoading();
            SheetOperations.refreshFilter();
            let exclude_uuids = SheetOperations.insertCykliczneIntecje(sheet);
            SheetOperations.refreshFilter(exclude_uuids);
            UIOperations.showDialog("Sukces", null, null, "Zaaktualizowane intencje", 200, 450);
        }
    }

}

function onEditInstallable(e: GoogleAppsScript.Events.SheetsOnEdit) {
    try {
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
    catch (error: any) {
        Logger.log(error);
        if (e.range) {
            e.range.getSheet().getRange(e.range.getRow(), e.range.getColumn()).setValue(e.oldValue);
        }
    }
}
function onOpenInstallable() {
    SpreadsheetApp
        .getUi()
        .createMenu("Modlitwa wstawiennicza")
        .addItem("Otwórz panel intencji", "showIntentionSidebar")
        .addItem("Otwórz panel uczestników", "showUsersSidebar")
        .addItem("Otwórz panel domyślnych intencji", "showDefaultIntentionsSidebar")
        .addItem("Odśwież ustwienia", "refreshVariables")
        .addItem("Dokumentacja", "showDocumentation")
        .addToUi();
    SheetOperations.refresh();
    showIntentionSidebar();
}

// function onOpen() {
//     SpreadsheetApp
//         .getUi()
//         .createMenu("Modlitwa wstawiennicza")
//         .addItem("Otwórz panel intencji", "showIntentionSidebar")
//         .addItem("Otwórz panel uczestników", "showUsersSidebar")
//         .addItem("Otwórz panel domyślnych intencji", "showDefaultIntentionsSidebar")
//         .addItem("Odśwież ustwienia", "refreshVariables")
//         .addItem("Dokumentacja", "showDocumentation")
//         .addToUi();
// }

function installTrigger() {
    ScriptApp.newTrigger("onEditInstallable")
        .forSpreadsheet(SpreadsheetApp.getActive())
        .onEdit()
        .create();

    ScriptApp.newTrigger("onOpenInstallable")
        .forSpreadsheet(SpreadsheetApp.getActive())
        .onOpen()
        .create();
}
