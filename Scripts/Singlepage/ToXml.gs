class StringRef {
  constructor(value) {
    this.value = value;
  }

  getValue() {
    return this.value;
  }

  setValue(value) {
    this.value = value;
  }

  replace(substr, newSubstr) {
    this.value = this.value.replace(substr, newSubstr);
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

  addSubelement(name, value) {
    this.subelements.push([ name, value ]);
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
  while (result < maxWidth && data[rowIndex][result] != "") {
    result++;
  }
  return result - columnIndex;
}

function getHeight(data, rowIndex, columnIndex, maxHeight) {
  var result = rowIndex;
  while (result < maxHeight && data[result][columnIndex] != "") {
    result++;
  }
  return result - rowIndex;
}
// TODO: read only values != "" ?
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
    var elementValue = data[rowIndex][c];
    var values = this.readVerticalValues(data, rowIndex + 1, c, height - 1);
    elementMap[elementValue] = values;
  }

  return elementMap;
}

function readElementsMap(data, width, height) {
  var elementsMap = {};
  
  var rowIndex = 0;
  while (rowIndex < height) {
    var elementWidth = this.getWidth(data, rowIndex, 0, width);
    var elementHeight = this.getHeight(data, rowIndex, 0, height);
    var elementName = data[rowIndex][0];
    elementsMap[elementName] = this.readElementMap(data, rowIndex, elementWidth, elementHeight);
    rowIndex += (1 + elementHeight);
  }

  return elementsMap;
}

function getXmlsMap(elementsMap) {
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

  return xmlsMap;
}

function fillXmlsMap(xmlsMap, elementsMap) {
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
          if (name in elementsMap && value in elementsMap[name]) {
            xml.addSubelement(name, value);
          } else {
            xml.addAttribute(name, value);
          }
        }
      }
    }
  }
}

function getStringsMap(xmlsMap) {
  var stringsMap = {};

  for (var elementName in xmlsMap) {
    var xmlMap = xmlsMap[elementName];
    var stringMap = {};
    for (var elementValue in xmlMap) {
      var xml = xmlMap[elementValue];

      var name = xml.getName();
      var attributes = xml.getAttributes();
      var subelements = xml.getSubelements();

      var stringValue = "";
      stringValue += "<" + name;
      for (var i = 0; i < attributes.length; ++i) {
        var attribute = attributes[i];
        stringValue += " " + attribute[0] + "=\"" + attribute[1] + "\"";
      }
      if (subelements.length > 0) {
        stringValue += ">\n";
        for (var i = 0; i < subelements.length; ++i) {
          stringValue += "{" + i + "}\n";
        }
        stringValue += "</" + name + ">";
      } else {
        stringValue += "/>"
      }

      var stringRef = new StringRef(stringValue);
      stringMap[elementValue] = stringRef;
    }
    stringsMap[elementName] = stringMap;
  }

  return stringsMap;
}
// TODO: remove duplicated values && implement reordering method if already set order is higher than it should be
function getDepthStack(elementsMap) {
  var nameToOrderMap = {};
  var stack = [];

  for (var elementName in elementsMap) {
    var elementMap = elementsMap[elementName];
    var names = elementMap[elementName];

    for (var elementValue in elementMap) {
      if (elementValue == elementName) {
        continue;
      }
      var values = elementMap[elementValue];

      if (!(elementName in nameToOrderMap)) {
        nameToOrderMap[elementName] = 0;
      }
      var order = nameToOrderMap[elementName];
      
      while (stack.length <= order) {
        stack.push([]);
      }
      stack[order].push([ elementName, elementValue ]);

      for (var i = 0; i < names.length; ++i) {
        var name = names[i];
        var value = values[i];

        if (name in elementsMap && value in elementsMap[name]) {

          if (!(name in nameToOrderMap)) {
            nameToOrderMap[name] = order + 1;
          }
          var suborder = nameToOrderMap[name];

          while (stack.length <= suborder) {
            stack.push([]);
          }
          stack[suborder].push([ name, value ]);
        }
      }
    }
  }

  return stack;
}

function fillStringsMap(xmlsMap, stringsMap, depthStack) {
  for (var i = depthStack.length - 1; i > -1; --i) {
    var pairs = depthStack[i];
    for (var j = 0; j < pairs.length; ++j) {
      var pair = pairs[j];
      var name = pair[0];
      var value = pair[1];

      var stringRef = stringsMap[name][value];
      var xml = xmlsMap[name][value];

      var stringValue = stringRef.getValue();
      var subelements = xml.getSubelements();
      for (var k = 0; k < subelements.length; ++k) {
        var subelement = subelements[k];
        var subname = subelement[0];
        var subvalue = subelement[1];

        var subStringRef = stringsMap[subname][subvalue];
        stringValue = stringValue.replace("{" + k + "}", subStringRef.getValue());
      }
      stringRef.setValue(stringValue);
    }
  }
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

  var xmlsMap = this.getXmlsMap(elementsMap);
  this.fillXmlsMap(xmlsMap, elementsMap);
  var stringsMap = this.getStringsMap(xmlsMap);
  var depthStack = this.getDepthStack(elementsMap);
  this.fillStringsMap(xmlsMap, stringsMap, depthStack);

  var result = stringsMap[elementName][elementValue].getValue();
  return result;
  var xmlDocument = XmlService.createDocument(xmlRoot);
  return xmlDocument;
}
