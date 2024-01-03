namespace Variables {
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
      export function getVariable(VARIABLE_NAME : string) : string {
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
}

export default Variables;
