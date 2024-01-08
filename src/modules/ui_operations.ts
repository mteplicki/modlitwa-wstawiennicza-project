namespace UIOperations {
      interface DayWeek {
        [index: number]: string;
      }

      export function getNieprzeczytanychWord(number: number) : string {
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

      export const dayWeek: DayWeek = {
        0: "Niedziela",
        1: "Poniedziałek",
        2: "Wtorek",
        3: "Środa",
        4: "Czwartek",
        5: "Piątek",
        6: "Sobota",
      };

      export function showLoading(){
        let ui = HtmlService.createTemplateFromFile('src/templates/Loading')
        let evaluate_ui = ui
          .evaluate()
          .setWidth(350)
          .setHeight(250);
        SpreadsheetApp.getUi().showModalDialog(evaluate_ui, "Ładowanie...");
      }
    
      export function openUrl( url : string ) : void{
        const html = `<html>
      <a id='url' href="${url}">${url}</a>
        <script>
           var winRef = window.open("${url}");
           winRef ? google.script.host.close() : window.alert('Skonfiguruj przeglądarkę, aby pozwalała na przekierowanie do ${url}') ;
           </script>
      </html>`; 
        var htmlOutput = HtmlService.createHtmlOutput(html).setWidth( 250 ).setHeight( 300 );
        SpreadsheetApp.getUi().showModalDialog( htmlOutput, `Otwieranie linku...` ); // https://developers.google.com/apps-script/reference/base/ui#showModalDialog(Object,String)  Requires authorization with this scope: https://www.googleapis.com/auth/script.container.ui  See https://developers.google.com/apps-script/concepts/scopes#setting_explicit_scopes
      }
      
    export function showDialog(title: string, warning: string | null = null, alert: string | null = null, info: string | null = null, height: number = 200, width: number = 450) : void {
      let ui = HtmlService.createTemplateFromFile('src/templates/RefreshWindow')
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
  
    export function showPickerDialog(range: string, height: number = 200, width: number = 450) : void {
      let ui = HtmlService.createTemplateFromFile('src/templates/DateRangePicker')
      ui.range = range;
      let title = "Wybierz zakres dat"
      let evaluate_ui = ui
        .evaluate()
        .setWidth(width)
        .setHeight(height);
      SpreadsheetApp.getUi().showModalDialog(evaluate_ui, title);
    }
}

export default UIOperations;