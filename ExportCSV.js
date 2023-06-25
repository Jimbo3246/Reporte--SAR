function onOpen() {
  //Codigo para crear un menu
  SpreadsheetApp.getUi()
                .createMenu("Exportar")
                .addItem("Exportar Excel", "obtener_Excel")
                .addItem("Exportar Csv", "obtener_Csv")
                .addToUi()
}

