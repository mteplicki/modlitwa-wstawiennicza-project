namespace EmailOperations {
    export interface EmailType {
        to: string,
        subject: string,
        name?: string,
        text: string,
        intentions?: string[],
        names?: string[]
    }

    export function sendEmail({ to, subject, name = undefined, text, intentions = undefined, names=undefined}: EmailType) {
        let template = HtmlService.createTemplateFromFile("src/templates/EmailTemplate");
        template.text = text;
        template.name = name;

        if (!intentions || !names) {
            template.notTable = true
            template.intentions = [""]
            template.names = [""]
        } else {
            if (names.length != intentions.length) throw new Error("Names and intentions length mismatch")
            template.notTable = false
            template.names = names
            template.intentions = intentions
        }
        let html = template.evaluate().getContent();
        MailApp.sendEmail({
            to: to,
            subject: subject,
            htmlBody: html,
            name: "Modlitwa wstawiennicza MOST",
        });
    }

}

export default EmailOperations