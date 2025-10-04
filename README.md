# Aplikacja modlitwy wstawienniczej

Aplikacja - panel administratora do zarządzania i wysyłania intencji przesyłanych do grupy modlitwy wstawienniczej Salezjańskiego Duszpasterstwa Akademickiego MOST.
Do stworzenia wykorzystano arkusz kalkulacyjny Google Sheets wraz z pakietem Google Apps Script. Arkusz jest powiązany z aplikacją internetową modlitwa-wstawiennicza-app.

### Idea

Moim celem było stworzenie aplikacji do zarządzania przesyłanych intencji do omodlenia w szybki i bezkosztowy sposób. Wybrałem przez to pakiet Google Apps Script, który choć jest bardzo ograniczony, to funkcjonalność miał wystarczającą dla tego projektu. Panel administratora to jest arkusz Google Sheets, zautomatyzowany za pomocą skryptów, panelu bocznego oraz reguł ograniczających modyfikowanie arkusza. Logika panelu administratora jest napisana w języku TypeScript.

### Funkcjonalności

Panel administratora ma następujące funkcjonalności:
 - baza intencji (w tym celu aplikacja pobiera wiadomości ze skrzynki mailowej, na które one przychodzą, i odpowiednio je parsuje tak, by dodać je do arkusza)
 - zarządzanie bazą osób omadlających
 - przydzielanie intencji w danym tygodniu do omadlającego
 - wysyłanie maili, które informują uczestników inicjatywy o przydzielonych im intencjach

Poza tym, projekt jest zintegrowany z aplikacją PWA modlitwy wstawienniczej (modlitwa-wstawiennicza-app):
- synchronizowanie intencji z bazą Firestore
- powiadamianie systemu Amazon Simple Notification Service w celu wysłania powiadomienia push przy przydzielaniu nowych intencji
