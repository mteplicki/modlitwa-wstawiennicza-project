namespace SheetOperations {
    export function getRange() : string {
        let sheet = SpreadsheetApp.getActive().getSheetByName('Intencje')
        if (sheet === null) {
          throw new Error("Sheet not found")
        }
        let value = sheet?.getRange("A1:I1").getValue() as string
        return value.split(": ")[1].split("[")[0].trim()
      }
    
      export function getRangeArray(): readonly [string, string] {
        let range = getRange()
        let range_array = range.split(" - ")
        let start = range_array[0]
        let end = range_array[1]
        return [start, end]
      }
    
      export function refreshFilter(exclude_uuids: string[] = []) : GoogleAppsScript.Spreadsheet.Filter {
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
        if (exclude_uuids.length > 0) {
          let criteria4 = SpreadsheetApp.newFilterCriteria().setHiddenValues(exclude_uuids)
          filter.setColumnFilterCriteria(1, criteria4.build())
        }
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
        // range?.protect();  
      }
    
    
    
      export function getFilteredValues(ss : GoogleAppsScript.Spreadsheet.Spreadsheet, sheet : GoogleAppsScript.Spreadsheet.Sheet) : string[][] {
        var url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/gviz/tq?tqx=out:csv&gid=" + sheet.getSheetId() + "&access_token=" + ScriptApp.getOAuthToken();
        var res = UrlFetchApp.fetch(url);
        var values = Utilities.parseCsv(res.getContentText());
        return values
      }
    
      export function range(size : number, startAt = 0) : number[] {
        return [...Array(size).keys()].map(i => i + startAt);
      }
    
      export function insertZaParafian(intentions_all_sheet: GoogleAppsScript.Spreadsheet.Sheet) : void {
        let filtered_values = getFilteredValues(SpreadsheetApp.getActive(), intentions_all_sheet).slice(1);
        let za_parafian = filtered_values.filter(value => value[8] === "TRUE");
        if (za_parafian.length === 0) {
            let date = getRangeArray()[1];
            date = date + " 00:00:00";
            insertIntention(date, "-", "Za parafian", intentions_all_sheet, true);
            refreshFilter();
        } else if (za_parafian.length > 1) {
            let exclude_uuids = za_parafian.map(value => value[0]).slice(1);
            refreshFilter(exclude_uuids);
        }
      }
}

export default SheetOperations;