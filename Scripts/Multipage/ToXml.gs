function toXmlDocument() {
  var sheetsArray = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var sheetsMap = this.getSheetsMap(sheetsArray);

  var sheet = sheetsArray[0];
  var name = sheet.getName();
  var root = XmlService.createElement(name);

  var range = sheet.getDataRange();
  var width = range.getWidth();
  var height = range.getHeight();
  var data = range.getValues();
  
  for (var r = 1; r < height; ++r) {
    var elementName = data[r][0];
    var elementValue = data[r][1];
    
    var element = this.createElement(elementName, elementValue, sheetsMap);
    if (element != null) {
      root.addContent(element);
    } else {
      root.setAttribute(elementName, elementValue);
    }
  }
  
  var document = XmlService.createDocument(root);
  return document;
}

function createElement(elementName, elementValue, sheetsMap) {
  var sheet = sheetsMap[elementName];
  if (sheet != null) {
    var element = XmlService.createElement(elementName);

    var range = sheet.getDataRange();
    var width = range.getWidth();
    var height = range.getHeight();
    var data = range.getValues();
    
    for (var c = 1; c < width; ++c) {
      var columnValue = data[0][c];

      if (columnValue == elementValue) {
        for (var r = 1; r < height; ++r) {
          var subelementName = data[r][0];
          var subelementValue = data[r][c];

          if (subelementValue != "") {
            var subelement = this.createElement(subelementName, subelementValue, sheetsMap);
            if (subelement != null) {
              element.addContent(subelement);
            } else {
              element.setAttribute(subelementName, subelementValue);
            }
          }
        }
        
        break;
      }
    }
    
    return element;
  }
  
  return null;
}
