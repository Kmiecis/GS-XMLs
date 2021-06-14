class StringRef {
  constructor(value) {
    this.value = value;
  }

  getValue() {
    return this.value;
  }

  replace(substr, newSubstr) {
    this.value.replace(substr, newSubstr);
  }
}

class XmlElement {
  constructor(name) {
    this.name = name;
    this.attributes = [];
    this.subelements = [];
  }

  addAttribute(name, value) {
    this.attributes.push([ name, value ]);
  }

  addSubelement(subelement) {
    this.subelements.push(subelement);
  }

  getName() {
    return this.name;
  }

  getAttributes() {
    return this.attributes;
  }

  getSubelements() {
    return this.subelements;
  }

  toString() {
    var result = "";
    var stack = [];
    stack.push([ this.name, this.attributes, this.subelements ]);
    while (stack.length > 0) {
      var item = stack.pop();

      var name = item[0];
      var attributes = item[1];
      var subelements = item[2];

      result += "<" + name;
      for (var i = 0; i < attributes.length; ++i) {
        var attribute = attributes[i];
        var attributeName = attribute[0];
        var attributeValue = attribute[1];
        result += " " + attributeName + "=\"" + attributeValue + "\"";
      }
      if (subelements.length > 0) {
        result += ">\n";
        for (var i = subelements.length - 1; i > -1; --i) {
          var subelement = subelements[i];
          stack.push([ subelement.getName(), subelement.getAttributes(), subelement.getSubelements() ]);
        }
        var endelement = stack[stack.length - subelements.length];
        for (var i = 3; i < item.length; ++i) {
          var endname = item[i];
          endelement.push(endname);
        }
        endelement.push(name);
      } else {
        result += "/>\n";

        for (var i = item.length - 1; i > 2; --i) {
          var endname = item[i];
          result += "</" + endname + ">\n";
        }
      }
    }
    return result;
  }
}

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

function readElementMap(data, rowIndex, width, height) {
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

function readElementsMap(data, width, height) {
  var elementsMap = {};
  
  var rowIndex = 0;
  while (rowIndex < height) {
    var elementWidth = this.getWidth(data, rowIndex, 0, width);
    var elementHeight = this.getHeight(data, rowIndex, 0, height);
    if (rowIndex == 0) { // Temp hack...
      elementHeight -= 1;
    }
    var elementName = data[rowIndex][0];
    elementsMap[elementName] = this.readElementMap(data, rowIndex, elementWidth, elementHeight);
    rowIndex += elementHeight;
  }

  return elementsMap;
}

function readXmlsMap(elementsMap) {
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
              xml.addSubelement(subXml);
            } else {
              xml.addAttribute(name, value);
            }
          } else {
            xml.addAttribute(name, value);
          }
        }
      }
    }
  }

  return xmlsMap;
}

function toXmlDocument() {
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var activeRange = activeSheet.getActiveRange();

  var range = activeSheet.getDataRange();
  var width = range.getWidth();
  var height = range.getHeight();
  var data = range.getValues();

  var elementsMap = this.readElementsMap(data, width, height);

  var rowIndex = activeRange.getRowIndex();
  var elementName = activeSheet.getRange(rowIndex, 1).getValue();
  var elementValue = activeRange.getValue();

  var xmlsMap = this.readXmlsMap(elementsMap);
  var xmlRoot = xmlsMap[elementName][elementValue];
  var result = xmlRoot.toString();
  return result;
  var xmlDocument = XmlService.createDocument(xmlRoot);
  return xmlDocument;
}
