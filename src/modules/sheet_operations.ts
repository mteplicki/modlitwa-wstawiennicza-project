import Utils from "./utils";
import UIOperations from "./ui_operations";

namespace SheetOperations {
  export function getRange(): string {
    let sheet = SpreadsheetApp.getActive().getSheetByName('Intencje')
    if (sheet === null) {
      throw new Error("Sheet not found")
    }
    let value = sheet?.getRange("A1:I1").getValue() as string
    return value.split(": ")[1].split("[")[0].trim()
  }

  export function getRangeArray(range: string): readonly [string, string] {
    let range_array = range.split(" - ")
    let start = range_array[0].trim()
    let end = range_array[1].trim()
    return [start, end]
  }

  export function refresh(showUI : boolean = true) {
    if (showUI) UIOperations.showLoading();
    let error: string | null = null;
    let unread: any[] = [];
    try {
      let [ss, intentions_all_sheet] = Utils.getSheetByName("Intencje");
      let START = Date.now();
      let unstored = GmailApp.search(
        'subject:"[Skrzynka intencji] Nowa intencja" -label:stored_sheet',
      );
      unstored.reverse();
      Logger.log(`Udało się pobrać ${unstored.length} wiadomości.`)

      for (let unstored_thread of unstored) {
        if (!Utils.isTimeLeft(START)) {
          throw new Error(
            "Przekroczono czas wykonania skryptu. Nie wszystkie intencje zostały zapisane. Spróbuj ponownie.",
          );
        }
        for (let unstored_message of unstored_thread.getMessages()) {
          let [date, name, intention] = SheetOperations.parseIntentionGmail(unstored_message);
          SheetOperations.insertIntention(
            date,
            name,
            intention,
            intentions_all_sheet,
          );
        }
        unstored_thread.addLabel(GmailApp.getUserLabelByName("stored_sheet"));
        unstored_thread.markRead();
      }
      Logger.log("Zapisano intencje z poczty.");
      unread = GmailApp.search(
        '-subject:"[Skrzynka intencji] Nowa intencja" is:unread',
      );
      Logger.log(`Znaleziono ${unread.length} nieprzeczytanych maili.`);
      SheetOperations.refreshFilter();
      Logger.log("Odświeżono filtr.");
      let exclude_uuids = SheetOperations.insertCykliczneIntecje(intentions_all_sheet);
      SheetOperations.refreshFilter(exclude_uuids);
      Logger.log("Dodano cykliczne intencje.");
    } catch (e: any) {
      if (e instanceof Error) {
        Logger.log(e.message);
        error = e.message;
      } else {
        Logger.log(String(e));
        error = String(e);
      }
    }
    let nieprzyczytane_word = UIOperations.getNieprzeczytanychWord(unread.length);
    let warning = unread.length > 0
      ? `W poczcie znajdują się ${unread.length} ${nieprzyczytane_word}. Otwórz Gmail i przeczytaj je.`
      : null;
    let alert = error !== null ? "Wystąpił błąd: " + error : null;
    let info = error === null ? "Odświeżono intencje." : null;
    let DIALOG_TITLE = "Odświeżanie intencji";
    if (showUI) UIOperations.showDialog(DIALOG_TITLE, warning, alert, info);
  }

  export function refreshFilter(exclude_uuids: string[] = []): GoogleAppsScript.Spreadsheet.Filter {
    let sheet = SpreadsheetApp.getActive().getSheetByName('Intencje')
    if (sheet === null) {
      throw new Error("Sheet not found")
    }
    if (sheet.getFilter() !== null) {
      sheet.getFilter()?.remove()
    }
    let range = sheet.getRange("A2:I")
    let filter = range.createFilter()
    let [start, end] = getRangeArray(SheetOperations.getRange())
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

  export function parseIntentionGmail(unstored_message: GoogleAppsScript.Gmail.GmailMessage): readonly [string, string, string] {

    let intentions_split = unstored_message.getPlainBody().split('\r\n\r\n--- Intencja: ---\r\n\r\n');

    let date = Utilities.formatDate(unstored_message.getDate(), "GMT+1", "yyyy-MM-dd HH:mm:ss");
    let name = intentions_split[0].split("--- Imię: ---\r\n\r\n")[1].trim();

    let intention = intentions_split[1]?.trim();
    return [date, name, intention];
  }

  export function insertIntention(date: string, name: string, intention: string, intentions_all_sheet: GoogleAppsScript.Spreadsheet.Sheet | null = null, cyclic_uuid: string = "") {
    intentions_all_sheet = intentions_all_sheet === null ? SpreadsheetApp.getActive().getSheetByName('Intencje') : intentions_all_sheet;
    intentions_all_sheet?.insertRowBefore(3);
    let range = intentions_all_sheet?.getRange("A3:I3");
    let deleted = "FALSE";
    let intention_corrected = intention;
    let uuid = Utilities.getUuid();

    range?.setValues([[uuid, date, date, name, intention, intention_corrected, deleted, "", cyclic_uuid]]);
  }



  export function getFilteredValues(ss: GoogleAppsScript.Spreadsheet.Spreadsheet, sheet: GoogleAppsScript.Spreadsheet.Sheet): string[][] {
    Logger.log("Pobieranie intencji")
    Logger.log(`ss.getId() : ${ss.getId()}`)
    Logger.log(`sheet.getSheetId() : ${sheet.getSheetId()}`)
    Logger.log(`sheet.getName() : ${sheet.getName()}`)
    Logger.log(`ScriptApp.getOAuthToken() : ${ScriptApp.getOAuthToken()}`)
    Logger.log(`url: https://docs.google.com/spreadsheets/d/${ss.getId()}/gviz/tq?tqx=out:csv&gid=${sheet.getSheetId()}`)
    var url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/gviz/tq?tqx=out:csv&gid=" + sheet.getSheetId();
    var res = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      method: "get",
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      }
    });
    var values = Utilities.parseCsv(res.getContentText());
    return values
  }

  export function range(size: number, startAt = 0): number[] {
    return [...Array(size).keys()].map(i => i + startAt);
  }

  export function insertCykliczneIntecje(intentions_all_sheet: GoogleAppsScript.Spreadsheet.Sheet): string[] {
    let filtered_values = getFilteredValues(SpreadsheetApp.getActive(), intentions_all_sheet).slice(1);
    Logger.log(`Pobrano ${filtered_values.length} intencji`)
    let [, cykliczne_arkusz] = Utils.getIntencjeCykliczne();
    let cykliczne_values = cykliczne_arkusz.getRange("A3:C").getValues() as string[][];
    let cykliczne_values_filtered = cykliczne_values.filter(value => value[0] !== "");
    let exclude_uuids: string[] = []
    for (let values of cykliczne_values_filtered) {
      let [uuid, name, intention] = values;
      let intentions_present = filtered_values.filter(value => value[8].includes(uuid));
      if (intentions_present.length === 0) {
        let date = getRangeArray(SheetOperations.getRange())[1];
        date = date + " 00:00:00";
        insertIntention(date, name, intention, intentions_all_sheet, uuid);
      } else if (intentions_present.length > 1) {
        intentions_present.sort((a, b) => new Date(b[1]).getTime() - new Date(a[1]).getTime());
        Logger.log(intentions_present)
        let intentions_excluded_uuids = intentions_present.map(value => value[0]).slice(1);
        exclude_uuids = exclude_uuids.concat(intentions_excluded_uuids)
      }
    }
    return exclude_uuids
  }
}

export default SheetOperations;