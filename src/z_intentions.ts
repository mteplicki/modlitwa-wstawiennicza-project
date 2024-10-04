import Utils from "./modules/utils";
import Variables from "./modules/variables";
import UIOperations from "./modules/ui_operations";
import SheetOperations from "./modules/sheet_operations";
import EmailOperations from "./modules/email_operations";
import FirebaseInit from "./modules/firebase_init";

function updateDateRange(range: string): void {
    try {
        Logger.log(range);
        const [,sheet] = Utils.getActiveSheetByName("Intencje");
        const [start, end] = SheetOperations.getRangeArray(range);
        const start_day = UIOperations.dayWeek[new Date(start).getDay()];
        const end_day = UIOperations.dayWeek[new Date(end).getDay()];
        Logger.log(`start: ${start}; end: ${end}`);
        Logger.log(`start_day_number: ${new Date(start).getDay()}, end_day_number: ${new Date(end).getDay()}`);
        Logger.log(`start_day: ${start_day}, end_day: ${end_day}`)
        sheet.getRange("A1:H1").setValue(
            `Intencje z zakresu: ${range} [${start_day} - ${end_day}]`,
        );
        SheetOperations.refreshFilter();
        let exclude_uuids = SheetOperations.insertCykliczneIntecje(sheet);
        SheetOperations.refreshFilter(exclude_uuids);
    } catch (e: any) {
        Utils.handleError(e);
    }
}

function showDateRangePicker(): void {
    UIOperations.showPickerDialog(SheetOperations.getRange(), 400, 650);
}

function refresh(): void {
    SheetOperations.refresh();
}

function insertFromDialog(date: string, name: string, intention: string): void {
    try {
        SheetOperations.insertIntention(date, name, intention);
        SheetOperations.refreshFilter();
    } catch (e: any) {
        Utils.handleError(e);
    }
}

function insertDialog(): void {
    try {
        if (SpreadsheetApp.getActive().getActiveSheet().getName() !== "Intencje") {
            throw new Error("Przełącz się na arkusz 'Intencje'");
        }
        let html = HtmlService.createTemplateFromFile("src/templates/AddIntention")
            .evaluate().setWidth(400).setHeight(650);
        SpreadsheetApp.getUi().showModalDialog(html, "Dodaj intencję");
    } catch (e: any) {
        Utils.handleError(e);
    }
}

function deleteIntention(): void {
    UIOperations.showLoading();
    try {
        //not deleting intention, just hiding it by setting "deleted" column to true
        let [_, sheet] = Utils.getActiveSheetByName("Intencje");
        if (!sheet) {
            throw new Error("Sheet not found");
        }
        let range = sheet.getActiveRange();
        if (!range) {
            throw new Error("Nie wybrano zakresu. Kliknij na interesujący Cię wiersz.")
        }
        if (range.getRow() < 3) {
            throw new Error("Nie można usunąć wiersza nagłówka.")
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
        Utils.handleError(e);
    }
}

function assignIntentions(): void {
    UIOperations.showLoading();
    try {
        let [ss, usersSheet] = Utils.getSheetByName("Uczestnicy");
        let [, intentionsSheet] = Utils.getActiveSheetByName("Intencje");
        let usersRange = usersSheet.getRange("B3:B");
        let usersValues = usersRange.getValues() as string[][];
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
        let set_intentionsUUIDShuffled = new Set([...intentionsUUIDShuffled]);
        for (let user of usersShuffled) {
            let intentions: string[];
            if (unassigned > 0) {
                intentions = intentionsUUIDShuffled.slice(last, last + assigned + 1);
                last += assigned + 1;
                unassigned--;
            } else {
                intentions = intentionsUUIDShuffled.slice(last, last + assigned);
                last += assigned;
            }
            if (intentions.length === 0) {
                let UUID = Utils.getRandomItem(set_intentionsUUIDShuffled);
                let row = intentionsSheet.getRange(`A3:A`).createTextFinder(UUID)
                    .findNext()?.getRow();
                if (row !== undefined) {
                    let range = intentionsSheet.getRange(`H${row}`)
                    let oldValue = range.getValue() as string;
                    range.setValue(oldValue.trim() + " " + user.trim());
                } else {
                    throw new Error("UUID not found");
                }
            } else {
                for (let UUID of intentions) {
                    let row = intentionsSheet.getRange(`A3:A`).createTextFinder(UUID)
                        .findNext()?.getRow();
                    if (row !== undefined) {
                        intentionsSheet.getRange(`H${row}`).setValue(user.trim());
                    } else {
                        throw new Error("UUID not found");
                    }
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
        Utils.handleError(e);
    }
}

function sendEmails(): void {
    UIOperations.showLoading();
    try {
        let [ss, sheet] = Utils.getActiveSheetByName("Intencje");
        let template = HtmlService.createTemplateFromFile(
            "src/templates/SendEmails",
        );
        Logger.log(PropertiesService.getDocumentProperties().getProperty("last_text"));
        template.last_text = PropertiesService.getDocumentProperties().getProperty("last_text") || "";
        let html = template.evaluate().setWidth(400).setHeight(650);
        SpreadsheetApp.getUi().showModalDialog(html, "Wyślij maile");
    } catch (e: any) {
        Utils.handleError(e);
    }
}

function sendEmailsCallback(text: string): void {
    type GroupedData = {
        [email: string]: { names: string[]; intentions: string[] };
    };

    function groupedDataInitialize(emails: string[][]) : GroupedData {
        let groupedData: GroupedData = {};
        //flatten emails to one array
        let emails2 : string[]  = [];
        for (let emailList of emails) {
            emails2.push(...emailList);
        }
        //remove duplicates
        emails2 = [...new Set(emails2)];

        for (let email of emails2) {
            groupedData[email] = { names: [], intentions: [] };
        }
        return groupedData;
    }

    function getDefaultIntention() : string[][] {
        let [ss, sheet] = Utils.getIntencjeOgolne();
        let range = sheet.getRange("A2:C");
        let values = range.getValues().slice(1) as string[][];
        return values;
    }

    Logger.log(text);

    try {
        let [ss, sheet] = Utils.getActiveSheetByName("Intencje");
        PropertiesService.getDocumentProperties().setProperty("last_text", text);
        let range = sheet.getRange("A2:I");
        let values = SheetOperations.getFilteredValues(ss, sheet).slice(1);

        let names = values.map((value) => value[3]);
        let intentions = values.map((value) => value[5]);
        let emails = values.map((value) => value[7]).map((value) => value.split(" "));

        let groupedData = groupedDataInitialize(emails);
        let defaultIntention = getDefaultIntention();

        for (let i = 0; i < emails.length; i++) {
            let emailList = emails[i] as string[];
            if (
                emailList.length === 0 || (emailList.length <= 1 && emailList[0] === "" )
            ) {
                throw new Error(
                    `Przydziel wszystkie intencje do uczestników. Wiersz ${i + 2
                    } nie ma przypisanego uczestnika.`,
                );
            }
            for (let email of emailList) {
                groupedData[email].names.push(names[i]);
                groupedData[email].intentions.push(intentions[i]);
            }
        }

        for (let ogólnaIntencja of defaultIntention) {
            for (let email in groupedData) {
                Logger.log(`email: ${email} ogólnaIntencja: ${ogólnaIntencja}`)
                groupedData[email].names.push(ogólnaIntencja[1]);
                groupedData[email].intentions.push(ogólnaIntencja[2]);
            }
        }

        let date = new Date();

        let dateString = Utilities.formatDate(date, "GMT+1", "yyyy-MM-dd HH:mm:ss");

        // Use groupedData for further processing
        for (let email in groupedData) {
            let dateRange = SheetOperations.getRangeArray(SheetOperations.getRange()) as readonly string[];
            dateRange = dateRange.map((date) => date.replace(/-/g, "."));
            let names = groupedData[email].names;
            let intentions = groupedData[email].intentions;
            let subject = "[Modlitwa wstawiennicza MOST] Intencje " + dateRange[0] +
                " - " + dateRange[1]
            EmailOperations.sendEmail({
                to: email,
                subject: subject,
                text: text,
                intentions: intentions,
                names: names,
            });
            try {
                FirebaseInit.firestore.createDocument(`intentions/${email}/dates/${dateString}`, {names : names, intentions : intentions, dateRange: dateRange});
            } catch (e) {
                Logger.log(e);
            }
        }
        const payload = JSON.stringify({
            "data": {
                "title": "Przydzielono nowe intencje",
                "body": "Przydzielono dla Ciebie nowe intencje. Sprawdź aplikację."
            },
            "topic": "default-topic",
            "webpush": {
                "fcmOptions": {
                    "link": "https://modlitwa-wstawiennicza-23992.web.app/intencje"
                }
            }
        })
        Logger.log("test:", payload)
        const request = UrlFetchApp.getRequest("https://ttrxyguhmmkd6gzd3bu4evucbe0acjlx.lambda-url.eu-north-1.on.aws/sendMessage", {
            method: "post",
            payload: payload,
            contentType: "application/json",
            muteHttpExceptions: true
        });
        Logger.log(request);
        const response = UrlFetchApp.fetch("https://ttrxyguhmmkd6gzd3bu4evucbe0acjlx.lambda-url.eu-north-1.on.aws/sendMessage", {
            method: "post",
            payload: payload,
            contentType: "application/json",
            muteHttpExceptions: true
        });
        Logger.log(response);
        UIOperations.showDialog("Wysłano maile", null, null, "Wysłano maile.");
        
    } catch (e: any) {
        Utils.handleError(e);
    }
}

function generateIntentionsDoc(): void {
    function getDefaultIntention() : string[][] {
        let [ss, sheet] = Utils.getIntencjeOgolne();
        let range = sheet.getRange("A2:C");
        let values = range.getValues().slice(1) as string[][];
        return values;
    }
    UIOperations.showLoading();
    try {
        let [ss, sheet] = Utils.getActiveSheetByName("Intencje");
        let rangeDate = SheetOperations.getRange();
        let defaultIntention = getDefaultIntention();
        let doc = DocumentApp.create(`Intencje_${rangeDate}`);
        if (sheet === null) {
            throw new Error("Sheet not found");
        }
        ss.getEditors().forEach((editor) => doc.addEditor(editor));
        let filteredBody = SheetOperations.getFilteredValues(
            SpreadsheetApp.getActive(),
            sheet,
        ).slice(1);
        let mappedBody = filteredBody.map((value) => [value[3], value[5]]);
        mappedBody = mappedBody.concat(defaultIntention.map((value) => [value[1], value[2]]));

        let header = doc.addHeader().setText(`Intencje z zakresu: ${rangeDate}`);
        interface DocumentStyle {
            [key: string]: string | number | boolean;
        }
        let style: DocumentStyle = {};
        style[DocumentApp.Attribute.FONT_FAMILY] = "Calibri";
        style[DocumentApp.Attribute.FONT_SIZE] = 14;
        style[DocumentApp.Attribute.BOLD] = false;
        style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
            DocumentApp.HorizontalAlignment.CENTER;
        header.setAttributes(style);
        doc.getBody().setText(
            "Módlmy się w intencjach przesłanych grupie modlitwy wstawienniczej:",
        );
        let nextParagraph = doc.getBody().appendParagraph("");
        for (let value of mappedBody) {
            let style: DocumentStyle = {};
            style[DocumentApp.Attribute.BOLD] = false;
            let bullet = doc.getBody().appendListItem(`${value[0]}: `).setAttributes(
                style,
            );
            style[DocumentApp.Attribute.BOLD] = false;
            bullet.appendText(`${value[1]}`).setAttributes(style);
        }
        // let listItems = mappedBody.map(value => doc.getBody().appendListItem())
        UIOperations.openUrl(doc.getUrl());
    } catch (e: any) {
        Utils.handleError(e);
    }
}
