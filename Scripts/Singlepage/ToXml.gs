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

function getDepthStack(elementsMap, rootName, rootValue) {
  var stack = [[[rootName, rootValue]]];
  var order = 0;

  while (order < stack.length) {
    var prev = stack[order];
    var next = [];
    for (var i = 0; i < prev.length; ++i) {
      var pair = prev[i];
      var elementName = pair[0];
      var elementValue = pair[1];

      var elementMap = elementsMap[elementName];
      var names = elementMap[elementName];
      var values = elementMap[elementValue];
      for (var j = 0; j < names.length; ++j) {
        var name = names[j];
        var value = values[j];

        if (name in elementsMap && value in elementsMap[name]) {
          next.push([name, value]);
        }
      }
    }
    if (next.length > 0) {
      stack.push(next);
    }
    order++;
  }

  return stack;
}

// TODO: remove duplicated values && implement reordering method if already set order is higher than it should be
/*function getDepthStack(elementsMap) {
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
      var suborder = order + 1;
      
      while (stack.length <= suborder) {
        stack.push({});
      }
      var ordered = stack[order];
      if (!(elementName in ordered)) {
        ordered[elementName] = new Set();
      }
      ordered[elementName].add(elementValue);

      for (var i = 0; i < names.length; ++i) {
        var name = names[i];
        var value = values[i];

        if (name in elementsMap && value in elementsMap[name]) {
          var subordered = stack[suborder];
          if (!(name in subordered)) {
            subordered[name] = new Set();
          }
          subordered[name].add(value);
        }
      }
    }
  }

  return stack;
}*/

function fillStringsMap(xmlsMap, stringsMap, depthStack) {
  for (var i = depthStack.length - 1; i > -1; --i) {
    var ordered = depthStack[i];
    for (var j = 0; j < ordered.length; ++j) {
      var pair = ordered[j];
      var name = pair[0];
      var value = pair[1];
    //for (var name in ordered) {
      //var values = ordered[name];
      //for (var value of values) {
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
      //}
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
  var elementName = "model";// activeSheet.getRange(rowIndex, 1).getValue();
  var elementValue = "ref1";// activeRange.getValue();

  var xmlsMap = this.getXmlsMap(elementsMap);
  this.fillXmlsMap(xmlsMap, elementsMap);
  var stringsMap = this.getStringsMap(xmlsMap);
  var depthStack = this.getDepthStack(elementsMap, elementName, elementValue);
  this.fillStringsMap(xmlsMap, stringsMap, depthStack);

  var result = stringsMap[elementName][elementValue].getValue();
  return result;
  var xmlDocument = XmlService.createDocument(xmlRoot);
  return xmlDocument;
}
