import SheetOperations from "./modules/sheet_operations";
import UIOperations from "./modules/ui_operations";
function onOpen(): void {
    SpreadsheetApp
        .getUi()
        .createMenu("Modlitwa wstawiennicza")
        .addItem("Otwórz panel intencji", "showIntentionSidebar")
        .addItem("Otwórz panel uczestników", "showUsersSidebar")
        .addItem("Otwórz panel domyślnych intencji", "showDefaultIntentionsSidebar")
        .addToUi();
}

function doGet(e : any) {
    Logger.log(e);
    const url = "https://docs.google.com/spreadsheets/d/1LJuwTBUkpp_KVYZH89w0e-mU0lPJZ1Eau_z2oIa3oPA/edit"
    // const url = "https://google.com"
    const html = `<html>
    <a id='url' href="${url}">${url}</a>
      <script>
         var winRef = window.open("${url}");
         winRef ? close() : window.alert('Skonfiguruj przeglądarkę, aby pozwalała na przekierowanie do ${url}') ;
         </script>
    </html>`; 
    return HtmlService.createHtmlOutput(html).setTitle("Modlitwa wstawiennicza");
}

function doPost(e: any) {
    return true;
}

