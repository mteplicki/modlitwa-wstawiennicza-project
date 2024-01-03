function onOpen(): void {
    SpreadsheetApp
        .getUi()
        .createMenu("Modlitwa wstawiennicza")
        .addItem("Otwórz panel intencji", "showIntentionSidebar")
        .addItem("Otwórz panel uczestników", "showUsersSidebar")
        .addItem("Otwórz panel domyślnych intencji", "showDefaultIntentionsSidebar")
        .addToUi();
}