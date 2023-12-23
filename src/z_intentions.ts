import Utils from "./modules/utils";
import Variables from "./modules/variables";
import UIOperations from "./modules/ui_operations";
import SheetOperations from "./modules/sheet_operations";

function updateDateRange(range: string): void {
    try {
        let ss = SpreadsheetApp.getActive();
        if (ss.getActiveSheet().getName() !== "Intencje") {
            throw new Error("Przełącz się na arkusz 'Intencje'");
        }
        let sheet = ss.getSheetByName("Intencje");
        if (sheet === null) {
            throw new Error("Sheet 'Intencje' not found");
        }
        const [start, end] = SheetOperations.getRangeArray();
        const start_day = UIOperations.dayWeek[new Date(start).getDay()];
        const end_day = UIOperations.dayWeek[new Date(end).getDay()];
        sheet?.getRange("A1:H1").setValue(
            `Intencje z zakresu: ${range} [${start_day} - ${end_day}]`,
        );
        SheetOperations.refreshFilter();
        SheetOperations.insertZaParafian(sheet);
    } catch (e: any) {
        if (e instanceof Error) {
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + e.message);
        } else {
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + String(e));
        }
    }
}

function showDateRangePicker(): void {
    UIOperations.showPickerDialog(SheetOperations.getRange(), 400, 650);
}

function refresh(): void {
    let error: string | null = null;
    let unread: any[] = [];
    try {
        let ss = SpreadsheetApp.getActive();
        if (ss.getActiveSheet().getName() !== "Intencje") {
            throw new Error("Przełącz się na arkusz 'Intencje'");
        }
        let START = Date.now();
        let unstored = GmailApp.search(
            'subject:"[Skrzynka intencji] Nowa intencja" -label:stored_sheet',
        );
        let intentions_all_sheet = ss.getSheetByName("Intencje");
        if (intentions_all_sheet === null) {
            throw new Error("Sheet not found");
        }
        unstored.reverse();

        for (let unstored_thread of unstored) {
            if (!Utils.isTimeLeft(START)) {
                throw new Error(
                    "Przekroczono czas wykonania skryptu. Nie wszystkie intencje zostały zapisane. Spróbuj ponownie.",
                );
            }
            for (let unstored_message of unstored_thread.getMessages()) {
                let [date, name, intention] = SheetOperations.parseIntentionGmail(
                    intentions_all_sheet,
                    unstored_message,
                );
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
        unread = GmailApp.search(
            '-subject:"[Skrzynka intencji] Nowa intencja" is:unread',
        );
        SheetOperations.refreshFilter();
        SheetOperations.insertZaParafian(intentions_all_sheet);
    } catch (e: any) {
        if (e instanceof Error) {
            error = e.message;
        } else {
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
    UIOperations.showDialog(DIALOG_TITLE, warning, alert, info);
}

function insertFromDialog(date: string, name: string, intention: string): void {
    try {
        let ss = SpreadsheetApp.getActive();
        if (ss.getActiveSheet().getName() !== "Intencje") {
            throw new Error("Przełącz się na arkusz 'Intencje'");
        }
        SheetOperations.insertIntention(date, name, intention);
        SheetOperations.refreshFilter();
    } catch (e: any) {
        if (e instanceof Error) {
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + e.message);
        } else {
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + String(e));
        }
    }
}

function insertDialog(): void {
    try {
        let ss = SpreadsheetApp.getActive();
        if (ss.getActiveSheet().getName() !== "Intencje") {
            throw new Error("Przełącz się na arkusz 'Intencje'");
        }
        let html = HtmlService.createTemplateFromFile("src/templates/AddIntention")
            .evaluate().setWidth(400).setHeight(650);
        SpreadsheetApp.getUi().showModalDialog(html, "Dodaj intencję");
    } catch (e: any) {
        if (e instanceof Error) {
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + e.message);
        } else {
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + String(e));
        }
    }
}

function onOpen(): void {
    SpreadsheetApp
        .getUi()
        .createMenu("Modlitwa wstawiennicza")
        .addItem("Otwórz panel intencji", "showIntentionSidebar")
        .addItem("Otwórz panel uczestników", "showUsersSidebar")
        .addToUi();
}

function showIntentionSidebar(): void {
    try {
        let widget = HtmlService.createHtmlOutputFromFile(
            "src/templates/AdminSidebar",
        ).setTitle("Modlitwa wstawiennicza MOST");
        SpreadsheetApp.getUi().showSidebar(widget);
    } catch (e: any) {
        if (e instanceof Error) {
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + e.message);
        } else {
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + String(e));
        }
    }
}

function deleteIntention(): void {
    try {
        //not deleting intention, just hiding it by setting "deleted" column to true
        let ss = SpreadsheetApp.getActive();
        if (ss.getActiveSheet().getName() !== "Intencje") {
            throw new Error("Przełącz się na arkusz 'Intencje'");
        }
        let sheet = ss.getSheetByName("Intencje");
        if (sheet === null) {
            throw new Error("Sheet not found");
        }
        let range = sheet.getActiveRange();
        if (range === null) {
            throw new Error(
                "Nie wybrano zakresu. Kliknij na interesujący Cię wiersz.",
            );
        }
        let UUID = sheet.getRange(`A${range.getRow()}`).getValue();
        let row = sheet.getRange(`A3:A`).createTextFinder(UUID).findNext()
            ?.getRow();
        if (row === undefined) {
            throw new Error("UUID not found");
        }
        SpreadsheetApp.getActive().getSheetByName("Intencje")?.getRange(`G${row}`)
            .setValue("TRUE");
        SheetOperations.refreshFilter();
        UIOperations.showDialog(
            "Usunięto intencję",
            null,
            null,
            "Usunięto intencję.",
        );
    } catch (e: any) {
        if (e instanceof Error) {
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + e.message);
        } else {
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + String(e));
        }
    }
}

function assignIntentions(): void {
    try {
        let ss = SpreadsheetApp.getActive();
        if (ss.getActiveSheet().getName() !== "Intencje") {
            throw new Error("Przełącz się na arkusz 'Intencje'");
        }
        let usersSheet = ss.getSheetByName("Uczestnicy");
        if (usersSheet === null) {
            throw new Error("Sheet not found");
        }
        let intentionsSheet = ss.getSheetByName("Intencje");
        if (intentionsSheet === null) {
            throw new Error("Sheet not found");
        }
        let usersRange = usersSheet.getRange("B3:B");
        let usersValues = usersRange.getValues();
        let usersLength = usersValues.length;
        let intentionsRange = SheetOperations.getFilteredValues(ss, intentionsSheet)
            .slice(1);
        let intentionsLength = intentionsRange.length;
        let intentionsUUID = intentionsRange.map((value) => value[0]);
        let usersShuffled = usersValues
            .map((value) => ({ value, sort: Math.random() }))
            .sort((a, b) => a.sort - b.sort)
            .map(({ value }) => value)
            .map((value) => value[0]);
        let intentionsUUIDShuffled = intentionsUUID
            .map((value) => ({ value, sort: Math.random() }))
            .sort((a, b) => a.sort - b.sort)
            .map(({ value }) => value);
        let assigned = Math.floor(intentionsLength / usersLength);
        let unassigned = intentionsLength % usersLength;
        let last = 0;
        for (let user of usersShuffled) {
            let intentions: any[];
            if (unassigned > 0) {
                intentions = intentionsUUIDShuffled.slice(last, last + assigned + 1);
                last += assigned + 1;
                unassigned--;
            } else {
                intentions = intentionsUUIDShuffled.slice(last, last + assigned);
                last += assigned;
            }
            for (let UUID of intentions) {
                let row = intentionsSheet.getRange(`A3:A`).createTextFinder(UUID)
                    .findNext()?.getRow();
                if (row !== undefined) {
                    intentionsSheet.getRange(`H${row}`).setValue(user);
                } else {
                    throw new Error("UUID not found");
                }
            }
        }
        UIOperations.showDialog(
            "Przypisano intencje",
            null,
            null,
            "Przypisano intencje.",
        );
    } catch (e: any) {
        if (e instanceof Error) {
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + e.message);
        } else {
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + String(e));
        }
    }
}

function sendEmails(): void {
    try {
        let ss = SpreadsheetApp.getActive();
        if (ss.getActiveSheet().getName() !== "Intencje") {
            throw new Error("Przełącz się na arkusz 'Intencje'");
        }
        let template = HtmlService.createTemplateFromFile(
            "src/templates/SendEmails",
        );
        template.last_text = Variables.getVariable("last_text");
        let html = template.evaluate().setWidth(400).setHeight(650);
        SpreadsheetApp.getUi().showModalDialog(html, "Wyślij maile");
    } catch (e: any) {
        if (e instanceof Error) {
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + e.message);
        } else {
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + String(e));
        }
    }
}

function sendEmailsCallback(text: string): void {
    try {
        let ss = SpreadsheetApp.getActive();
        if (ss.getActiveSheet().getName() !== "Intencje") {
            throw new Error("Przełącz się na arkusz 'Intencje'");
        }
        Variables.saveVariable("last_text", text);
        let sheet = ss.getSheetByName("Intencje");
        if (sheet === null) {
            throw new Error("Sheet not found");
        }
        let range = sheet.getRange("A2:I");
        let values = SheetOperations.getFilteredValues(ss, sheet).slice(1);

        let names = values.map((value) => value[3]);
        let intentions = values.map((value) => value[5]);
        let emails = values.map((value) => value[7]);

        let groupedData: {
            [email: string]: { names: string[]; intentions: string[] };
        } = {};

        for (let i = 0; i < emails.length; i++) {
            let email = emails[i] as string;
            if (
                (typeof email == "string" && email.trim() === "") ||
                typeof email != "string"
            ) {
                throw new Error(
                    `Przydziel wszystkie intencje do uczestników. Wiersz ${i + 2
                    } nie ma przypisanego uczestnika.`,
                );
            }
            if (!groupedData[email]) {
                groupedData[email] = { names: [], intentions: [] };
            }
            groupedData[email].names.push(names[i]);
            groupedData[email].intentions.push(intentions[i]);
        }

        // Use groupedData for further processing
        for (let email in groupedData) {
            let dateRange = SheetOperations.getRangeArray() as readonly string[];
            dateRange = dateRange.map((date) => date.replace(/-/g, "."));
            let names = groupedData[email].names;
            let intentions = groupedData[email].intentions;
            let mailTemplate = HtmlService.createTemplateFromFile(
                "src/templates/EmailTemplate",
            );
            mailTemplate.names = names;
            mailTemplate.intentions = intentions;
            mailTemplate.text = text;
            let html = mailTemplate.evaluate().getContent();

            MailApp.sendEmail({
                to: email,
                subject: "[Modlitwa wstawiennicza MOST] Intencje " + dateRange[0] +
                    " - " + dateRange[1],
                htmlBody: html,
                name: "Modlitwa wstawiennicza MOST",
            });
        }
        UIOperations.showDialog("Wysłano maile", null, null, "Wysłano maile.");
    } catch (e: any) {
        if (e instanceof Error) {
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + e.message);
        } else {
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + String(e));
        }
    }
}

function generateIntentionsDoc(): void {
    try {
        let ss = SpreadsheetApp.getActive();
        if (ss.getActiveSheet().getName() !== "Intencje") {
            throw new Error("Przełącz się na arkusz 'Intencje'");
        }
        let rangeDate = SheetOperations.getRange();
        let doc = DocumentApp.create(`Intencje_${rangeDate}`);

        let sheet = ss.getSheetByName("Intencje");
        if (sheet === null) {
            throw new Error("Sheet not found");
        }
        ss.getEditors().forEach((editor) => doc.addEditor(editor));
        let filteredBody = SheetOperations.getFilteredValues(
            SpreadsheetApp.getActive(),
            sheet,
        ).slice(1);
        let mappedBody = filteredBody.map((value) => [value[3], value[5]]);

        let header = doc.addHeader().setText(`Intencje z zakresu: ${rangeDate}`);
        interface DocumentStyle {
            [key: string]: string | number | boolean;
        }
        let style: DocumentStyle = {};
        style[DocumentApp.Attribute.FONT_FAMILY] = "Calibri";
        style[DocumentApp.Attribute.FONT_SIZE] = 14;
        style[DocumentApp.Attribute.BOLD] = true;
        style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
            DocumentApp.HorizontalAlignment.CENTER;
        header.setAttributes(style);
        doc.getBody().setText(
            "Módlmy się w intencjach przesłanych grupie modlitwy wstawienniczej:",
        );
        let nextParagraph = doc.getBody().appendParagraph("");
        for (let value of mappedBody) {
            let style: DocumentStyle = {};
            style[DocumentApp.Attribute.BOLD] = true;
            let bullet = doc.getBody().appendListItem(`${value[0]}: `).setAttributes(
                style,
            );
            style[DocumentApp.Attribute.BOLD] = false;
            bullet.appendText(`${value[1]}`).setAttributes(style);
        }
        // let listItems = mappedBody.map(value => doc.getBody().appendListItem())
        UIOperations.openUrl(doc.getUrl());
    } catch (e: any) {
        if (e instanceof Error) {
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + e.message);
        } else {
            UIOperations.showDialog("Błąd", null, "Wystąpił błąd: " + String(e));
        }
    }
}
