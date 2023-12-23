import { Utils } from "utils";

function updateDateRange(range: string) {
    try {
        let sheet = SpreadsheetApp.getActive().getSheetByName('Intencje')
        if (sheet === null) {
            throw new Error("Sheet 'Intencje' not found")
        }
        const [start, end] = Utils.getRangeArray()
        const start_day = Utils.dayWeek[new Date(start).getDay()]
        const end_day = Utils.dayWeek[new Date(end).getDay()]
        sheet?.getRange("A1:H1").setValue(`Intencje z zakresu: ${range} [${start_day} - ${end_day}]`)
        sheet?.getRange("A1:H1").protect()
        Utils.refreshFilter()
        Utils.insertZaParafian(sheet)
        
    } catch (e: any) {
        if (e instanceof Error) {
            Utils.showDialog("Błąd", null, "Wystąpił błąd: " + e.message)
        } else {
            Utils.showDialog("Błąd", null, "Wystąpił błąd: " + String(e))
        }
    }
}

function showDateRangePicker() {
    Utils.showPickerDialog(Utils.getRange(), 400, 650)
}

function refresh() {
    let error: string | null = null
    let unread: any[] = []
    try{
        let START = Date.now();
        let unstored = GmailApp.search('subject:"[Skrzynka intencji] Nowa intencja" -label:stored_sheet')
        let intentions_all_sheet = SpreadsheetApp.getActive().getSheetByName('Intencje')
        if (intentions_all_sheet === null) {
            throw new Error("Sheet not found")
        }
        unstored.reverse()
        
        for (let unstored_thread of unstored) {
            if (!Utils.isTimeLeft(START)) {
                throw new Error("Przekroczono czas wykonania skryptu. Nie wszystkie intencje zostały zapisane. Spróbuj ponownie.")
            }
            for (let unstored_message of unstored_thread.getMessages()) {
                let [date, name, intention] = Utils.parseIntentionGmail(intentions_all_sheet, unstored_message);
                Utils.insertIntention(date, name, intention, intentions_all_sheet);
            }
            unstored_thread.addLabel(GmailApp.getUserLabelByName('stored_sheet'))
        }
        unread = GmailApp.search('-subject:"[Skrzynka intencji] Nowa intencja" is:unread')
        Utils.refreshFilter()
        Utils.insertZaParafian(intentions_all_sheet);
        
    } catch (e: any) {
        if (e instanceof Error) {
            error = e.message
        } else {
            error = String(e)
        }
    }
    let nieprzyczytane_word = Utils.getNieprzeczytanychWord(unread.length)
    let warning = unread.length > 0 ? `W poczcie znajdują się ${unread.length} ${nieprzyczytane_word}. Otwórz Gmail i przeczytaj je.` : null
    let alert = error !== null ? "Wystąpił błąd: " + error : null
    let info = error === null ? "Odświeżono intencje." : null
    let DIALOG_TITLE = "Odświeżanie intencji"
    Utils.showDialog(DIALOG_TITLE, warning, alert, info)
}



function insertFromDialog(date: string, name: string, intention: string) {
    try{
        Utils.insertIntention(date, name, intention);
        Utils.refreshFilter()
    } catch (e: any) {
        if (e instanceof Error) {
            Utils.showDialog("Błąd", null, "Wystąpił błąd: " + e.message)
        } else {
            Utils.showDialog("Błąd", null, "Wystąpił błąd: " + String(e))
        }
    }    
}

function insertDialog() {
    let html = HtmlService.createTemplateFromFile('src/AddIntention').evaluate().setWidth(400).setHeight(650)
    SpreadsheetApp.getUi().showModalDialog(html, 'Dodaj intencję')
}


function onOpen() {
    SpreadsheetApp
        .getUi()
        .createMenu('Intencje')
        .addItem('Otwórz panel', 'showAdminSidebar')
        .addToUi();
}

function showAdminSidebar() {
    var widget = HtmlService.createHtmlOutputFromFile('src/AdminSidebar')
    SpreadsheetApp.getUi().showSidebar(widget);
}

function deleteIntention(){
    try{
        //not deleting intention, just hiding it by setting "deleted" column to true
        let sheet = SpreadsheetApp.getActive().getSheetByName('Intencje')
        if (sheet === null) {
            throw new Error("Sheet not found")
        }
        let range = sheet.getActiveRange()
        if (range === null) {
            throw new Error("Nie wybrano zakresu. Kliknij na interesujący Cię wiersz.")
        }
        let UUID = sheet.getRange(`A${range.getRow()}`).getValue()
        let row = sheet.getRange(`A3:A`).createTextFinder(UUID).findNext()?.getRow()
        if (row === undefined) {
            throw new Error("UUID not found")
        }
        SpreadsheetApp.getActive().getSheetByName('Intencje')?.getRange(`G${row}`).setValue("TRUE")
        Utils.refreshFilter()
        Utils.showDialog("Usunięto intencję", null, null, "Usunięto intencję.")
    } catch (e: any) {
        if (e instanceof Error) {
            Utils.showDialog("Błąd", null, "Wystąpił błąd: " + e.message)
        } else {
            Utils.showDialog("Błąd", null, "Wystąpił błąd: " + String(e))
        }
    }
}

function assignIntentions(){
    try{
        let ss = SpreadsheetApp.getActive()
        let usersSheet = ss.getSheetByName('Uczestnicy')
        if (usersSheet === null) {
            throw new Error("Sheet not found")
        }
        let intentionsSheet = ss.getSheetByName('Intencje')
        if (intentionsSheet === null) {
            throw new Error("Sheet not found")
        }
        let usersRange = usersSheet.getRange("B3:B")
        let usersValues = usersRange.getValues()
        let usersLength = usersValues.length
        let intentionsRange = Utils.getFilteredValues(ss, intentionsSheet).slice(1)
        let intentionsLength = intentionsRange.length
        let intentionsUUID = intentionsRange.map(value => value[0])
        let usersShuffled = usersValues
            .map(value => ({ value, sort: Math.random() }))
            .sort((a, b) => a.sort - b.sort)
            .map(({ value }) => value)
            .map((value) => value[0])
        let intentionsUUIDShuffled = intentionsUUID
            .map(value => ({ value, sort: Math.random() }))
            .sort((a, b) => a.sort - b.sort)
            .map(({ value }) => value)
        let assigned = Math.floor(intentionsLength / usersLength)
        let unassigned = intentionsLength % usersLength
        let last = 0
        for (let user of usersShuffled) {
            let intentions : any[]
            if (unassigned > 0) {
                intentions = intentionsUUIDShuffled.slice(last, last + assigned + 1)
                last += assigned + 1
                unassigned--
            } else {
                intentions = intentionsUUIDShuffled.slice(last, last + assigned)
                last += assigned
            }
            for (let UUID of intentions) {
                let row = intentionsSheet.getRange(`A3:A`).createTextFinder(UUID).findNext()?.getRow()
                if (row !== undefined) {
                    intentionsSheet.getRange(`H${row}`).setValue(user)
                } else {
                    throw new Error("UUID not found")
                }
            }
        }
        Utils.showDialog("Przypisano intencje", null, null, "Przypisano intencje.")
    } catch (e: any) {
        if (e instanceof Error) {
            Utils.showDialog("Błąd", null, "Wystąpił błąd: " + e.message)
        } else {
            Utils.showDialog("Błąd", null, "Wystąpił błąd: " + String(e))
        }
    }
}

function sendEmails(){
    let template = HtmlService.createTemplateFromFile('src/SendEmails')
    template.last_text = Utils.getVariable('last_text')
    let html = template.evaluate().setWidth(400).setHeight(650)
    SpreadsheetApp.getUi().showModalDialog(html, 'Wyślij maile')
}

function sendEmailsCallback(text: string) {
    try {
        Utils.saveVariable('last_text', text);
        let ss = SpreadsheetApp.getActive();
        let sheet = ss.getSheetByName('Intencje');
        if (sheet === null) {
            throw new Error("Sheet not found");
        }
        let range = sheet.getRange("A2:I");
        let values = Utils.getFilteredValues(ss, sheet).slice(1);

        let names = values.map(value => value[3]);
        let intentions = values.map(value => value[5]);
        let emails = values.map(value => value[7]);

        let groupedData: { [email: string]: { names: string[], intentions: string[] } } = {};

        for (let i = 0; i < emails.length; i++) {
            let email = emails[i] as string;
            if ((typeof email == "string" && email.trim() === "") || typeof email != "string") {
                throw new Error(`Przydziel wszystkie intencje do uczestników. Wiersz ${i + 2} nie ma przypisanego uczestnika.`);
            }
            if (!groupedData[email]) {
                groupedData[email] = { names: [], intentions: [] };
            }
            groupedData[email].names.push(names[i]);
            groupedData[email].intentions.push(intentions[i]);
        }

        // Use groupedData for further processing
        for (let email in groupedData) {
            let dateRange = Utils.getRangeArray() as string[];
            dateRange = dateRange.map(date => date.replace(/-/g, "."));
            let names = groupedData[email].names;
            let intentions = groupedData[email].intentions;
            let mailTemplate = HtmlService.createTemplateFromFile('src/EmailTemplate');
            mailTemplate.names = names;
            mailTemplate.intentions = intentions;
            mailTemplate.text = text;
            let html = mailTemplate.evaluate().getContent();

            MailApp.sendEmail({
                to: email,
                subject: "[Modlitwa wstawiennicza MOST] Intencje " + dateRange[0] + " - " + dateRange[1],
                htmlBody: html,
                name: "Modlitwa wstawiennicza MOST"
            })
        }
        Utils.showDialog("Wysłano maile", null, null, "Wysłano maile.");

    } catch (e: any) {
        if (e instanceof Error) {
            Utils.showDialog("Błąd", null, "Wystąpił błąd: " + e.message)
        } else {
            Utils.showDialog("Błąd", null, "Wystąpił błąd: " + String(e))
        }
    }
}

function generateIntentionsDoc(){
    try{
        let rangeDate = Utils.getRange()
        let doc = DocumentApp.create(`Intencje_${rangeDate}`)
        let ss = SpreadsheetApp.getActive()
        let sheet = ss.getSheetByName('Intencje')
        if (sheet === null) {
            throw new Error("Sheet not found")
        }
        ss.getEditors().forEach(editor => doc.addEditor(editor))
        let filteredBody = Utils.getFilteredValues(SpreadsheetApp.getActive(), sheet).slice(1)
        let mappedBody = filteredBody.map(value => [value[3], value[5]])
        
        let header = doc.addHeader().setText(`Intencje z zakresu: ${rangeDate}`)
        interface DocumentStyle {
            [key: string]: string | number | boolean
        }
        let style : DocumentStyle = {};
        style[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
        style[DocumentApp.Attribute.FONT_SIZE] = 14;
        style[DocumentApp.Attribute.BOLD] = true;
        style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
        header.setAttributes(style);
        doc.getBody().setText("Módlmy się w intencjach przesłanych grupie modlitwy wstawienniczej:")
        let nextParagraph = doc.getBody().appendParagraph("")
        for (let value of mappedBody) {
            let style : DocumentStyle = {};
            style[DocumentApp.Attribute.BOLD] = true;
            let bullet = doc.getBody().appendListItem(`${value[0]}: `).setAttributes(style)
            style[DocumentApp.Attribute.BOLD] = false;
            bullet.appendText(`${value[1]}`).setAttributes(style)
        }
        // let listItems = mappedBody.map(value => doc.getBody().appendListItem())
        Utils.openUrl(doc.getUrl())
    } catch (e: any) {
        if (e instanceof Error) {
            Utils.showDialog("Błąd", null, "Wystąpił błąd: " + e.message)
        } else {
            Utils.showDialog("Błąd", null, "Wystąpił błąd: " + String(e))
        }
    }

}

