export namespace Utils {
  const ONE_SECOND = 1000;
  const ONE_MINUTE = ONE_SECOND * 60;
  const MAX_EXECUTION_TIME = ONE_MINUTE * (5);

  interface DayWeek {
    [index: number]: string;
  }

  export function openUrl( url : string ){
    Logger.log('openUrl. url: ' + url);
    const html = `<html>
  <a id='url' href="${url}">Click here</a>
    <script>
       var winRef = window.open("${url}");
       winRef ? google.script.host.close() : window.alert('Configure browser to allow popup to redirect you to ${url}') ;
       </script>
  </html>`; 
    Logger.log('openUrl. html: ' + html);
    var htmlOutput = HtmlService.createHtmlOutput(html).setWidth( 250 ).setHeight( 300 );
    Logger.log('openUrl. htmlOutput: ' + htmlOutput);
    SpreadsheetApp.getUi().showModalDialog( htmlOutput, `openUrl function in generic.gs is now opening a URL...` ); // https://developers.google.com/apps-script/reference/base/ui#showModalDialog(Object,String)  Requires authorization with this scope: https://www.googleapis.com/auth/script.container.ui  See https://developers.google.com/apps-script/concepts/scopes#setting_explicit_scopes
  }

  export const dayWeek: DayWeek = {
    0: "Niedziela",
    1: "Poniedziałek",
    2: "Wtorek",
    3: "Środa",
    4: "Czwartek",
    5: "Piątek",
    6: "Sobota",
  };

  export function getNieprzeczytanychWord(number: number) {
    switch (number) {
        case 0:
            return "nieprzeczytanych wiadomości"
        case 1:
            return "nieprzeczytana wiadomość"
        case 2:
            return "nieprzeczytane wiadomości"
        case 3:
            return "nieprzeczytane wiadomości"
        case 4:
            return "nieprzeczytane wiadomości"
        default:
            return "nieprzeczytanych wiadomości"
    }
}

  export const isTimeLeft = (START: number) => {
    return MAX_EXECUTION_TIME > Date.now() - START;
  };

  export function showDialog(title: string, warning: string | null = null, alert: string | null = null, info: string | null = null, height: number = 200, width: number = 450) {
    let ui = HtmlService.createTemplateFromFile('src/RefreshWindow')
    ui.show_warning = warning !== null;
    ui.show_alert = alert !== null;
    ui.show_info = info !== null;
    ui.warning = warning;
    ui.alert = alert;
    ui.info = info;
    ui.title = title;
    let evaluate_ui = ui
      .evaluate()
      .setWidth(width)
      .setHeight(height);
    SpreadsheetApp.getUi().showModalDialog(evaluate_ui, title);
  }

  export function showPickerDialog(range: string, height: number = 200, width: number = 450) {
    let ui = HtmlService.createTemplateFromFile('src/DateRangePicker')
    ui.range = range;
    let title = "Wybierz zakres dat"
    let evaluate_ui = ui
      .evaluate()
      .setWidth(width)
      .setHeight(height);
    SpreadsheetApp.getUi().showModalDialog(evaluate_ui, title);
  }

  export function getRange() : string {
    let sheet = SpreadsheetApp.getActive().getSheetByName('Intencje')
    if (sheet === null) {
      throw new Error("Sheet not found")
    }
    let value = sheet?.getRange("A1:I1").getValue() as string
    return value.split(": ")[1].split("[")[0].trim()
  }

  export function getRangeArray(): [string, string] {
    let range = getRange()
    let range_array = range.split(" - ")
    let start = range_array[0]
    let end = range_array[1]
    return [start, end]
  }

  export function refreshFilter() {
    let sheet = SpreadsheetApp.getActive().getSheetByName('Intencje')
    if (sheet === null) {
      throw new Error("Sheet not found")
    }
    if (sheet.getFilter() !== null) {
      sheet.getFilter()?.remove()
    }
    let range = sheet.getRange("A2:I")
    let filter = range.createFilter()
    let [start, end] = getRangeArray()
    start = start + " 00:00:00"
    end = end + " 23:59:59"
    let start_date = new Date(start)
    let end_date = new Date(end)
    start_date.setDate(start_date.getDate() - 1)
    end_date.setDate(end_date.getDate() + 1)
    let criteria1 = SpreadsheetApp.newFilterCriteria()
      .whenDateAfter(start_date)
      .build()
    filter.setColumnFilterCriteria(2, criteria1)
    let criteria2 = SpreadsheetApp.newFilterCriteria()
      .whenDateBefore(end_date)
      .build()
    filter.setColumnFilterCriteria(3, criteria2)
    let criteria3 = SpreadsheetApp.newFilterCriteria()
      .whenTextEqualTo("FALSE")
    filter.setColumnFilterCriteria(7, criteria3)
    return filter
  }

  export function parseIntentionGmail(intentions_all_sheet: GoogleAppsScript.Spreadsheet.Sheet | null, unstored_message: GoogleAppsScript.Gmail.GmailMessage): readonly [string, string, string] {
    
    let intentions_split = unstored_message.getPlainBody().split('\r\n\r\n--- Intencja: ---\r\n\r\n');

    let date = Utilities.formatDate(unstored_message.getDate(), "GMT+1", "yyyy-MM-dd HH:mm:ss");
    let name = intentions_split[0].split("--- Imię: ---\r\n\r\n")[1].trim();

    let intention = intentions_split[1]?.trim();
    return [date, name, intention];
  }

  export function insertIntention(date: string, name: string, intention: string, intentions_all_sheet: GoogleAppsScript.Spreadsheet.Sheet | null = null, za_parafian: boolean = false) {
    intentions_all_sheet = intentions_all_sheet === null ? SpreadsheetApp.getActive().getSheetByName('Intencje') : intentions_all_sheet;
    intentions_all_sheet?.insertRowBefore(3);
    let range = intentions_all_sheet?.getRange("A3:I3");
    let deleted = "FALSE";
    let intention_corrected = intention;
    let uuid = Utilities.getUuid();
    let za_parafian_string = za_parafian ? "TRUE" : "FALSE";

    range?.setValues([[uuid, date, date, name, intention, intention_corrected, deleted, "", za_parafian_string]]);
    range?.protect();  
  }

  function getRowVariable(sheet : GoogleAppsScript.Spreadsheet.Sheet,VARIABLE_NAME : string) : number {
    sheet.getFilter()?.remove()
    let row = sheet.getRange("A1:A").createTextFinder(VARIABLE_NAME).findNext()?.getRow()
    if (row === undefined) {
      throw new Error("Variable not found")
    }
    return row
  }

  export function saveVariable(VARIABLE_NAME : string, value : string) : void {
    let sheet = SpreadsheetApp.getActive().getSheetByName(MW_VARIABLES)
    if (sheet === null) {
      throw new Error("Sheet not found")
    }
    let last_text_range = getRowVariable(sheet, VARIABLE_NAME)
    try {
      sheet.getRange(`B${last_text_range}`).setValue(value)
    } catch (e) {
      throw new Error("Variable not found")
    }
  }

  const MW_VARIABLES = 'MW_VARIABLES';
  export function getVariable(VARIABLE_NAME : string) : any {
    let sheet = SpreadsheetApp.getActive().getSheetByName(MW_VARIABLES)
    if (sheet === null) {
      throw new Error("Sheet not found")
    }
    let last_text_range = getRowVariable(sheet, VARIABLE_NAME)

    try {
      let value = sheet.getRange(`B${last_text_range}`).getValue()
      return value
    } catch (e) {
      throw new Error("Variable not found")
    }
  }

  export function getFilteredValues(ss : GoogleAppsScript.Spreadsheet.Spreadsheet, sheet : GoogleAppsScript.Spreadsheet.Sheet) : string[][] {
    var url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/gviz/tq?tqx=out:csv&gid=" + sheet.getSheetId() + "&access_token=" + ScriptApp.getOAuthToken();
    var res = UrlFetchApp.fetch(url);
    var values = Utilities.parseCsv(res.getContentText());
    return values
  }

  export function range(size : number, startAt = 0) {
    return [...Array(size).keys()].map(i => i + startAt);
  }

  export function insertZaParafian(intentions_all_sheet: GoogleAppsScript.Spreadsheet.Sheet) {
    let filtered_values = getFilteredValues(SpreadsheetApp.getActive(), intentions_all_sheet).slice(1);
    if (filtered_values.map(value => value[8]).filter(value => value === "TRUE").length === 0) {
        let date = getRangeArray()[1];
        date = date + " 00:00:00";
        insertIntention(date, "-", "Za parafian", intentions_all_sheet, true);
        refreshFilter();
    }
}
}




