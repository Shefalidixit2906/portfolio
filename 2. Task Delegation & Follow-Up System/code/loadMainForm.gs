function doGet(e) {
  const htmlServ = HtmlService.createTemplateFromFile("main");
  const html = htmlServ.evaluate();
  return html
}


function loadInside(){
  const htmlServ = HtmlService.createTemplateFromFile("main");
  const html = htmlServ.evaluate();
  html.setWidth(950).setHeight(700)
  const ui = SpreadsheetApp.getUi()
  ui.showModalDialog(html,"BCI Delegation Sheet")
}