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
        let sheet = SpreadsheetApp.getActive().getSheetByName(Ustawienia)
        if (sheet === null) {
          throw new Error("Sheet not found")
        }
        let last_text_range = getRowVariable(sheet, VARIABLE_NAME)
        try {
          sheet.getRange(`C${last_text_range}`).setValue(value)
          let cache = PropertiesService.getDocumentProperties()
          cache.setProperty(VARIABLE_NAME, value)
          Logger.log(`Saved ${VARIABLE_NAME} with value ${value}`)
        } catch (e) {
          throw new Error("Variable not found")
        }
      }
    
      const Ustawienia = 'Ustawienia';
      export function getVariable(VARIABLE_NAME : string) : string {    
        try {
          let value = PropertiesService.getDocumentProperties().getProperty(VARIABLE_NAME)
          if (value !== null) {
            return value
          }
          throw new Error("Variable not found")
        } catch (e) {
          throw new Error("Variable not found")
        }
      }

      export function synchronizeWithCache() : void {
        let sheet = SpreadsheetApp.getActive().getSheetByName(Ustawienia)
        if (sheet === null) {
          throw new Error("Sheet not found")
        }
        let range = sheet.getRange("A2:C")
        let range_values = range.getValues() as [string,string,string][]
        let cache = PropertiesService.getDocumentProperties()
        for (let row of range_values) {
          let [key, ,value] = row
          if (key !== "") {
            cache.setProperty(key, value)
            Logger.log(`Synchronized ${key} with value ${value}`)
          }
        }
      }
}

export default Variables;
