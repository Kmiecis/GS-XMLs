class XmlFromSheet {
  static asString(sheet) {
    var rootName = sheet.getRange(1, 1).getValue();
    var rootValue = sheet.getRange(2, 1).getValue();

    var elementsMap = this.createElementsMap(sheet);  
    var xmlsMap = this.createXmlsMap(elementsMap);
    var depthStack = this.createDepthStack(xmlsMap, rootName, rootValue);
    var stringsMap = this.createStringsMap(xmlsMap, depthStack);

    const XML_HEADER = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n";
    return XML_HEADER + stringsMap[rootName][rootValue].getValue();
  }

  static createElementsMap(sheet) {
    var elementsMap = {};
    this.fillElementsMap(elementsMap, sheet);
    return elementsMap;
  }

  static fillElementsMap(elementsMap, sheet) {
    var data = sheet.getDataRange().getValues();
    var height = data.length;
    var width = data[0].length;

    var rowIndex = 0;
    while (rowIndex < height) {
      var elementWidth = this.getDataWidth(data, rowIndex, 0, width);
      var elementHeight = this.getDataHeight(data, rowIndex, 0, height);
      var elementName = data[rowIndex][0];
      elementsMap[elementName] = this.createElementMap(data, rowIndex, elementWidth, elementHeight);
      rowIndex += (1 + elementHeight);
    }
  }

  static getDataWidth(data, rowIndex, columnIndex, maxWidth) {
    var result = columnIndex;
    while (result < maxWidth && data[rowIndex][result] != "") {
      result++;
    }
    return result - columnIndex;
  }

  static getDataHeight(data, rowIndex, columnIndex, maxHeight) {
    var result = rowIndex;
    while (result < maxHeight && data[result][columnIndex] != "") {
      result++;
    }
    return result - rowIndex;
  }
  
  static createElementMap(data, rowIndex, width, height) {
    var elementMap = {};
    this.fillElementMapHorizontally(elementMap, data, rowIndex, width, height);
    return elementMap;
  }

  static fillElementMapVertically(elementMap, data, rowIndex, width, height) {
    var elementName = data[rowIndex][0];

    var names = this.readVerticalValues(data, rowIndex + 1, 0, height - 1);
    elementMap[elementName] = names;

    for (var c = 1; c < width; ++c) {
      var elementValue = data[rowIndex][c];
      var values = this.readVerticalValues(data, rowIndex + 1, c, height - 1);
      elementMap[elementValue] = values;
    }
  }

  static readVerticalValues(data, rowIndex, columnIndex, height) {
    var values = new Array(height);
    for (var r = 0; r < height; ++r) {
      values[r] = data[rowIndex + r][columnIndex];
    }
    return values;
  }

  static fillElementMapHorizontally(elementMap, data, rowIndex, width, height) {
    var elementName = data[rowIndex][0];

    var names = this.readHorizontalValues(data, rowIndex, 1, width - 1);
    elementMap[elementName] = names;

    for (var r = 1; r < height; ++r) {
      var elementValue = data[rowIndex + r][0];
      var values = this.readHorizontalValues(data, rowIndex + r, 1, width - 1);
      elementMap[elementValue] = values;
    }
  }

  static readHorizontalValues(data, rowIndex, columnIndex, width) {
    var values = new Array(width);
    for (var c = 0; c < width; ++c) {
      values[c] = data[rowIndex][columnIndex + c];
    }
    return values;
  }
  
  static createXmlsMap(elementsMap) {
    var xmlsMap = {};
    this.preloadXmlsMap(xmlsMap, elementsMap);
    this.fillXmlsMap(xmlsMap, elementsMap);
    return xmlsMap;
  }

  static preloadXmlsMap(xmlsMap, elementsMap) {
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
  }

  static fillXmlsMap(xmlsMap, elementsMap) {
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
              xml.addChild(name, value);
            } else {
              xml.addAttribute(name, value);
            }
          }
        }
      }
    }
  }

  static createDepthStack(xmlsMap, rootName, rootValue) {
    var stack = [[[rootName, rootValue]]];
    var order = 0;

    while (order < stack.length) {
      var prev = stack[order];
      var next = [];
      for (var i = 0; i < prev.length; ++i) {
        var pair = prev[i];
        var name = pair[0];
        var value = pair[1];

        var xmlMap = xmlsMap[name];
        var xml = xmlMap[value];

        var children = xml.getChildren();
        for (var c = 0; c < children.length; ++c) {
          var child = children[c];
          var childName = child[0];
          var childValue = child[1];

          next.push([childName, childValue]);
        }
      }
      if (next.length > 0) {
        stack.push(next);
      }
      order++;
    }

    return stack;
  }

  static createStringsMap(xmlsMap, depthStack) {
    var stringsMap = {};
    this.preloadStringsMap(stringsMap, xmlsMap);
    this.fillStringsMap(stringsMap, xmlsMap, depthStack);
    return stringsMap;
  }

  static preloadStringsMap(stringsMap, xmlsMap) {
    for (var elementName in xmlsMap) {
      var xmlMap = xmlsMap[elementName];
      var stringMap = {};
      for (var elementValue in xmlMap) {
        var xml = xmlMap[elementValue];

        var name = xml.getName();
        var attributes = xml.getAttributes();
        var children = xml.getChildren();

        var stringValue = "";
        stringValue += "<" + name;
        for (var i = 0; i < attributes.length; ++i) {
          var attribute = attributes[i];
          stringValue += " " + attribute[0] + "=\"" + attribute[1] + "\"";
        }
        if (children.length > 0) {
          stringValue += ">\n";
          for (var i = 0; i < children.length; ++i) {
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
  }

  static fillStringsMap(stringsMap, xmlsMap, depthStack) {
    for (var i = depthStack.length - 1; i > -1; --i) {
      var stack = depthStack[i];
      for (var j = 0; j < stack.length; ++j) {
        var pair = stack[j];
        var name = pair[0];
        var value = pair[1];

        var stringRef = stringsMap[name][value];
        var xml = xmlsMap[name][value];

        var stringValue = stringRef.getValue();
        var children = xml.getChildren();
        for (var k = 0; k < children.length; ++k) {
          var child = children[k];
          var childName = child[0];
          var childValue = child[1];

          var subStringRef = stringsMap[childName][childValue];
          stringValue = stringValue.replace("{" + k + "}", subStringRef.getValue());
        }
        stringRef.setValue(stringValue);
      }
    }
  }
}
