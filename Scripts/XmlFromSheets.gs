class XmlFromSheets {
  static asString(sheets) {
    var rootSheet = sheets[0];
    var rootName = rootSheet.getName();
    var rootValue = "";

    var sheetsMap = this.createSheetsMap(sheets);
    var elementsMap = this.createElementsMap(sheetsMap);
    var xmlsMap = this.createXmlsMap(elementsMap);
    var depthStack = this.createDepthStack(xmlsMap, rootName);
    var stringsMap = this.createStringsMap(xmlsMap, depthStack);

    const XML_HEADER = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n";
    return XML_HEADER + stringsMap[rootName][rootValue].getValue();
  }

  static createSheetsMap(sheets) {
    var sheetsMap = {};
    this.fillSheetsMap(sheetsMap, sheets);
    return sheetsMap;
  }

  static fillSheetsMap(sheetsMap, sheets) {
    for (var i = 0; i < sheets.length; ++i) {
      var sheet = sheets[i];
      var name = sheet.getName();
      var data = sheet.getDataRange().getValues();
      sheetsMap[name] = data;
    }
  }

  static createElementsMap(sheetsMap) {
    var elementsMap = {};
    this.fillElementsMap(elementsMap, sheetsMap);
    return elementsMap;
  }

  static fillElementsMap(elementsMap, sheetsMap) {
    for (var name in sheetsMap) {
      var data = sheetsMap[name];
      elementsMap[name] = this.createElementMap(name, data);
    }
  }

  static createElementMap(name, data) {
    var elementMap = {};
    this.fillElementMap(elementMap, name, data);
    return elementMap;
  }

  static fillElementMap(elementMap, name, data) {
    if (data[0][0].includes(':')) {
      this.fillElementMapVertically(elementMap, name, data);
    } else {
      this.fillElementMapHorizontally(elementMap, name, data);
    }
  }

  static fillElementMapVertically(elementMap, name, data) {
    var width = data[0].length;
    var names = this.readValuesVertically(data, 0);
    elementMap[name] = names;
    for (var c = 1; c < width; ++c) {
      var value = data[0][c];
      var values = this.readValuesVertically(data, c);
      elementMap[value] = values;
    }
  }

  static readValuesVertically(data, column) {
    var values = new Array(data.length - 1);
    for (var i = 0; i < values.length; ++i) {
      values[i] = data[i + 1][column];
    }
    return values;
  }

  static fillElementMapHorizontally(elementMap, name, data) {
    var height = data.length;
    var names = this.readValuesHorizontally(data, 0);
    elementMap[name] = names;
    for (var r = 1; r < height; ++r) {
      var value = data[r][0];
      var values = this.readValuesHorizontally(data, r);
      elementMap[value] = values;
    }
  }

  static readValuesHorizontally(data, row) {
    var values = new Array(data[0].length - 1);
    for (var i = 0; i < values.length; ++i) {
      values[i] = data[row][i + 1];
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
    for (var elementName in xmlsMap) {
      var xmlMap = xmlsMap[elementName];
      var elementMap = elementsMap[elementName];

      var names = elementMap[elementName];
      for (var elementValue in xmlMap) {
        var xml = xmlMap[elementValue];
        var values = elementMap[elementValue];

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

  static createDepthStack(xmlsMap, rootName) {
    var stack = [[[rootName, ""]]];
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
