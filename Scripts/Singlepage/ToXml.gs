function getWidth(data, rowIndex, columnIndex, maxWidth) {
  var result = columnIndex;
  var value = data[rowIndex][result++];
  while (value != "" && result < maxWidth) {
    value = data[rowIndex][result++];
  }
  return result - columnIndex;
}

function getHeight(data, rowIndex, columnIndex, maxHeight) {
  var result = rowIndex;
  var value = data[result++][columnIndex];
  while (value != "" && result < maxHeight) {
    value = data[result++][columnIndex];
  }
  return result - rowIndex;
}

function readVerticalValues(data, rowIndex, columnIndex, height) {
  var values = new Array(height);
  for (var r = 0; r < height; ++r) {
    values[r] = data[rowIndex + r][columnIndex];
  }
  return values;
}

function getElementMap(data, rowIndex, width, height) {
  var elementName = data[rowIndex][0];
  var elementMap = {};

  var names = this.readVerticalValues(data, rowIndex + 1, 0, height - 1);
  elementMap[elementName] = names;

  for (var c = 1; c < width; ++c) {
    var elementRef = data[rowIndex][c];
    var values = this.readVerticalValues(data, rowIndex + 1, c, height - 1);
    elementMap[elementRef] = values;
  }

  return elementMap;
}

function getElementsMap(data, width, height) {
  var elementsMap = {};
  
  var rowIndex = 0;
  while (rowIndex < height) {
    var elementWidth = this.getWidth(data, rowIndex, 0, width);
    var elementHeight = this.getHeight(data, rowIndex, 0, height);
    if (rowIndex == 0) { // Temp hack...
      elementHeight -= 1;
    }
    var elementName = data[rowIndex][0];
    elementsMap[elementName] = this.getElementMap(data, rowIndex, elementWidth, elementHeight);
    rowIndex += elementHeight;
  }

  return elementsMap;
}

function readXmlElement(elementsMap, elementName, elementValue) {
  var elementMap = elementsMap[elementName];
  if (elementMap != null) {
    var xmlElement = XmlService.createElement(elementName);
    var names = elementMap[elementName];
    var values = elementMap[elementValue];
    
    // assert(values != null)
    if (values == null) { // Happens when there is both element and attribute with same name
      return null;//throw "Values for '" + elementName + "' and '" + elementValue + "' are null";
    }

    for (var i = 0; i < names.length; ++i) {
      var name = names[i];
      var value = values[i];
      if (value != null && value != "") {
        var subXmlElement = this.readXmlElement(elementsMap, name, value);
        if (subXmlElement != null) {
          xmlElement.addContent(subXmlElement);
        } else {
          xmlElement.setAttribute(name, value);
        }
      }
    }
    return xmlElement;
  } else {
    return null;
  }
}

function readXmlElements(elementsMap, rootName, rootValue) {
  var xmlsMap = {};
  for (var elementName in elementsMap) {
    var elementMap = elementsMap[elementName];
    var xmlMap = {};
    for (var elementValue in elementMap) {
      if (elementValue == elementName) {
        continue;
      }
      xmlMap[elementValue] = new XmlElement(elementName);
    }
    xmlsMap[elementName] = xmlMap;
  }

  for (var elementName in elementsMap) {
    var elementMap = elementsMap[elementName];
    var xmlMap = xmlsMap[elementName];

    var names = elementMap[elementName];
    for (var elementValue in elementMap) {
      if (elementValue == elementName) {
        continue;
      }
      var values = elementMap[elementValue];

      var xml = xmlMap[elementValue];
      for (var i = 0; i < names.length; ++i) {
        var name = names[i];
        var value = values[i];

        if (value != "") {
          var subXmlMap = xmlsMap[name];
          if (subXmlMap != null) {
            var subXml = subXmlMap[value];
            if (subXml != null) {
              xml.addContent(subXml);
            } else {
              xml.setAttribute(name, value);
            }
          } else {
            xml.setAttribute(name, value);
          }
        }
      }
    }
  }

  return xmlsMap[rootName][rootValue];
}

class XmlElement {
  constructor(name) {
    this.name = name;
    this.subelements = [];
    this.attributes = [];
  }

  setAttribute(name, value) {
    this.attributes.push([ name, value ]);
  }

  addContent(subelement) {
    this.subelements.push(subelement);
  }

  toString() {
    var result = "<" + this.name;
    for (var i = 0; i < this.attributes.length; ++i) {
      var attribute = this.attributes[i];
      var name = attribute[0];
      var value = attribute[1];
      result += " " + name + "=" + value;
    }
    if (this.subelements.length == 0) {
      result += "/>\n";
    } else {
      result += ">\n";
      for (var i = 0; i < this.subelements.length; ++i) {
        var subelement = this.subelements[i];
        result += "\t" + subelement.toString();
      }
      result += "</" + this.name + ">\n";
    }
    return result;
  }
}

function toXmlDocument() {
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var activeRange = activeSheet.getActiveRange();

  var range = activeSheet.getDataRange();
  var width = range.getWidth();
  var height = range.getHeight();
  var data = range.getValues();

  var elementsMap = this.getElementsMap(data, width, height);

  var rowIndex = activeRange.getRowIndex();
  var elementName = activeSheet.getRange(rowIndex, 1).getValue();
  var elementValue = activeRange.getValue();

  var xmlRoot = this.readXmlElements(elementsMap, elementName, elementValue);
  return xmlRoot.toString();
  var xmlDocument = XmlService.createDocument(xmlRoot);
  return xmlDocument;
}
