function openSideBar() {
  var html = HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Helper')
      .setWidth(300);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(html);
}