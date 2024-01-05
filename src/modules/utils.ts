import UIOperations from "./ui_operations";

namespace Utils {

    const ONE_SECOND = 1000
    const ONE_MINUTE = ONE_SECOND * 60
    const MAX_EXECUTION_TIME = ONE_MINUTE * (5)

    export function isTimeLeft(START: number) : boolean {
      return MAX_EXECUTION_TIME > Date.now() - START;
    };

    export const getRandomItem = <T>(set : Set<T>) : T => [...set][Math.floor(Math.random()*set.size)]

    export function handleError(e: any) {
        if (e instanceof Error) {
            Logger.log(e.message);
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + e.message);
        } else {
            Logger.log(String(e));
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + String(e));
        }
    }

    export function getSheetByName(name: string): readonly [GoogleAppsScript.Spreadsheet.Spreadsheet, GoogleAppsScript.Spreadsheet.Sheet] {
        let ss = SpreadsheetApp.getActive();
        let sheet = ss.getSheetByName(name);
        if (sheet === null) {
            throw new Error(`Arkusz ${name} nie istnieje`);
        }
        return [ss, sheet];
    }

    export function getActiveSheetByName(name: string): readonly [GoogleAppsScript.Spreadsheet.Spreadsheet, GoogleAppsScript.Spreadsheet.Sheet] {
        let ss = SpreadsheetApp.getActive();
        let sheet = ss.getActiveSheet();
        if (sheet.getName() !== name) {
            throw new Error(`Przełącz się na arkusz ${name}`);
        }
        return [ss, sheet];
    }

    interface SheetType {
      (): readonly [GoogleAppsScript.Spreadsheet.Spreadsheet, GoogleAppsScript.Spreadsheet.Sheet]
    }

    export const getIntencje : SheetType = () => getSheetByName("Intencje")
    export const getUczestnicy : SheetType = () => getSheetByName("Uczestnicy")
    export const getIntencjeOgolne : SheetType = () => getSheetByName("Intencje-ogólne")
    export const getIntencjeCykliczne : SheetType = () => getSheetByName("Intencje-cykliczne")

    export const getActiveIntencje : SheetType = () => getActiveSheetByName("Intencje")
    export const getActiveUczestnicy : SheetType = () => getActiveSheetByName("Uczestnicy")
    export const getActiveIntencjeOgolne : SheetType = () => getActiveSheetByName("Intencje-ogólne")
    export const getActiveIntencjeCykliczne : SheetType = () => getActiveSheetByName("Intencje-cykliczne")

    export const getActiveIntencjeOgolneOrCykliczne : SheetType = () => {
        let name = SpreadsheetApp.getActive().getActiveSheet().getName()
        switch (name) {
            case "Intencje-ogólne":
                return Utils.getActiveIntencjeOgolne()
            case "Intencje-cykliczne":
                return Utils.getActiveIntencjeCykliczne()
            default:
                throw new Error("Nieprawidłowy arkusz - musi być Intencje-ogólne lub Intencje-cykliczne")
        }
    }

}

export default Utils;