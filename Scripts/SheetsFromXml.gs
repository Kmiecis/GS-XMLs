class SheetsFromXml {
  static fillSpreadsheet(spreadsheet, document) {
    var rootXml = document.getRootElement();
    var rootName = rootXml.getName();
    var rootValue = "ref0";

    var xmlsMap = this.createXmlsMap(rootXml, rootValue);
    var invXmlsMap = this.createInvXmlsMap(xmlsMap);
    var depthStack = this.createDepthStack(xmlsMap, rootName, rootValue);

    this.trimXmlsMap(xmlsMap, invXmlsMap, depthStack);

    var elementsMap = this.createElementsMap(xmlsMap);
    this.parseToSpreadsheet(spreadsheet, elementsMap);
  }

  static createXmlsMap(rootXml, rootValue) {
    var xmlsMap = {};
    this.fillXmlsMap(xmlsMap, rootXml, rootValue);
    return xmlsMap;
  }

  static fillXmlsMap(xmlsMap, rootXml, rootValue) {
    var count = 1;

    var stack = [[ rootXml, rootValue ]];
    while (stack.length > 0) {
      var pair = stack.pop();

      var xml = pair[0];
      var name = xml.getName();
      var value = pair[1];

      var element = new XmlElement(name);

      var attributes = xml.getAttributes();
      for (var a = 0; a < attributes.length; ++a) {
        var attribute = attributes[a];
        var attributeName = attribute.getName();
        var attributeValue = attribute.getValue();
        element.addAttribute(attributeName, attributeValue);
      }

      var children = xml.getChildren();
      for (var c = 0; c < children.length; ++c) {
        var child = children[c];
        var childName = child.getName();
        var childValue = "ref" + count++;
        element.addChild(childName, childValue);

        stack.push([ child, childValue ]);
      }

      if (!(name in xmlsMap)) {
        xmlsMap[name] = {};
      }
      var xmlMap = xmlsMap[name];

      xmlMap[value] = element;
    }
    return xmlsMap;
  }

  static createInvXmlsMap(xmlsMap) {
    var invXmlsMap = {};
    this.fillInvXmlsMap(invXmlsMap, xmlsMap);
    return invXmlsMap;
  }

  static fillInvXmlsMap(invXmlsMap, xmlsMap) {
    for (var name in xmlsMap) {
      var xmlMap = xmlsMap[name];
      for (var value in xmlMap) {
        var xml = xmlMap[value];

        var children = xml.getChildren();
        for (var c = 0; c < children.length; ++c) {
          var child = children[c];
          var childName = child[0];
          var childValue = child[1];

          if (!(childName in invXmlsMap)) {
            invXmlsMap[childName] = {};
          }
          var invXmlMap = invXmlsMap[childName];

          if (!(childValue in invXmlMap)) {
            invXmlMap[childValue] = [];
          }
          var references = invXmlMap[childValue];
          
          references.push([ name, value ]);
        }
      }
    }
  }

  static createDepthStack(xmlsMap, rootName, rootValue) {
    var stack = [[[ rootName, rootValue ]]];
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

          next.push([ childName, childValue ]);
        }
      }

      if (next.length > 0) {
        stack.push(next);
      }

      order++;
    }

    return stack;
  }

  static trimXmlsMap(xmlsMap, invXmlsMap, depthStack) {
    for (var i = depthStack.length - 1; i > -1; --i) {
      var stack = depthStack[i];
      for (var s = 0; s < stack.length; ++s) {
        var pair = stack[s];
        var name = pair[0];
        var value = pair[1];

        var xmlMap = xmlsMap[name];
        var xml = xmlMap[value];
        if (xml == null) {
          // Happens when we remove value from xmlMap but it is still on stack.
          // We don't care about depth stack being correct later on so we can ignore this issue.
          continue;
        }
        var invXmlMap = invXmlsMap[name];

        for (var otherValue in xmlMap) {
          if (otherValue == value) {
            continue;
          }
          var otherXml = xmlMap[otherValue];

          if (xml.equals(otherXml)) {
            var references = invXmlMap[otherValue];
            var x = xml.equals(otherXml);
            for (var r = 0; r < references.length; ++r) {
              var reference = references[r];
              var refName = reference[0];
              var refValue = reference[1];

              var refXml = xmlsMap[refName][refValue];
              refXml.replaceChild(name, otherValue, value);
            }

            delete xmlMap[otherValue];
            delete invXmlMap[otherValue];
          }
        }
      }
    }
  }

  static createElementsMap(xmlsMap) {
    var elementsMap = {};
    this.fillElementsMap(elementsMap, xmlsMap);
    this.fillElementsMapWithNames(elementsMap, xmlsMap);
    this.fillElementsMapWithValues(elementsMap, xmlsMap);
    return elementsMap;
  }

  static fillElementsMap(elementsMap, xmlsMap) {
    for (var name in xmlsMap) {
      var xmlMap = xmlsMap[name];

      var elementMap = {};
      for (var value in xmlMap) {
        elementMap[value] = [];
      }
      elementMap[name] = [];

      elementsMap[name] = elementMap;
    }
  }

  static fillElementsMapWithNames(elementsMap, xmlsMap) {
    for (var name in xmlsMap) {
      var xmlMap = xmlsMap[name];

      var attributeNames = new Set();
      var childNames = {};

      for (var value in xmlMap) {
        var xml = xmlMap[value];
        var attributes = xml.getAttributes();
        var children = xml.getChildren();

        for (var a = 0; a < attributes.length; ++a) {
          var attribute = attributes[a];
          var attributeName = attribute[0];
          attributeNames.add(attributeName);
        }

        var xmlChildNames = {};
        for (var c = 0; c < children.length; ++c) {
          var child = children[c];
          var childName = child[0];

          if (!(childName in xmlChildNames)) {
            xmlChildNames[childName] = 0;
          }
          xmlChildNames[childName] += 1;
        }
        for (var childName in xmlChildNames) {
          var newCount = xmlChildNames[childName];
          if (!(childName in childNames)) {
            childNames[childName] = 0;
          }
          var oldCount = childNames[childName];
          if (oldCount < newCount) {
            childNames[childName] = newCount;
          }
        }
      }

      var names = [];
      for (var attributeName of attributeNames) {
        names.push(attributeName);
      }

      for (var childName in childNames) {
        var count = childNames[childName];
        for (var i = 0; i < count; ++i) {
          names.push(childName);
        }
      }
      
      elementsMap[name][name] = names;
    }
  }

  static fillElementsMapWithValues(elementsMap, xmlsMap) {
    for (var name in xmlsMap) {
      var xmlMap = xmlsMap[name];
      var elementMap = elementsMap[name];
      var names = elementMap[name];

      for (var value in xmlMap) {
        var xml = xmlMap[value];
        var attributes = xml.getAttributes();
        var children = xml.getChildren();

        var values = new Array(names.length);

        for (var a = 0; a < attributes.length; ++a) {
          var attribute = attributes[a];
          var attributeName = attribute[0];
          var attributeValue = attribute[1];

          var nameIndex = names.indexOf(attributeName);
          values[nameIndex] = attributeValue;
        }

        var childValuesMap = {};
        for (var c = 0; c < children.length; ++c) {
          var child = children[c];
          var childName = child[0];
          var childValue = child[1];

          if (!(childName in childValuesMap)) {
            childValuesMap[childName] = [];
          }
          childValuesMap[childName].push(childValue);
        }
        for (var childName in childValuesMap) {
          var childValues = childValuesMap[childName];
          var nameIndex = names.indexOf(childName);
          for (var cv = 0; cv < childValues.length; ++cv) {
            var childValue = childValues[cv];
            values[nameIndex + cv] = childValue;
          }
        }

        elementMap[value] = values;
      }
    }
  }

  static createSheetsMap(sheets) {
    var sheetsMap = {};
    this.fillSheetsMap(sheetsMap, sheets);
    return sheetsMap;
  }

  static fillSheetsMap(sheetsMap, sheets) {
    for (var i = 0; i < sheets.length; ++i) {
      var sheet = sheets[i];
      sheetsMap[sheet.getName()] = sheet;
    }
  }

  static parseToSpreadsheet(spreadsheet, elementsMap) {
    var sheets = spreadsheet.getSheets();
    var sheetsMap = this.createSheetsMap(sheets);
    for (var elementName in elementsMap) {
      var elementMap = elementsMap[elementName];
      if (!(elementName in sheetsMap)) {
        sheetsMap[elementName] = spreadsheet.insertSheet(elementName);
      }
      var sheet = sheetsMap[elementName];

      this.parseToSheetHorizontally(sheet, elementName, elementMap);
    }
  }

  static parseToSheetVertically(sheet, elementName, elementMap) {
    var size = this.getElementMapSize(elementMap);
    var rows = size[1];
    var columns = size[0];

    var data = new Array(rows);
    for (var r = 0; r < rows; ++r) {
      data[r] = new Array(columns);
    }

    var names = elementMap[elementName];
    for (var n = 0; n < names.length; ++n) {
      data[n + 1][0] = names[n];
    }
    var column = 1;
    for (var elementValue in elementMap) {
      if (elementValue == elementName) {
        continue;
      }
      data[0][column] = elementValue;
      var values = elementMap[elementValue];
      for (var v = 0; v < values.length; ++v) {
        data[v + 1][column] = values[v];
      }
      column += 1;
    }

    sheet.getRange(1, 1, data.length, data[0].length)
      .setValues(data);
  }

  static parseToSheetHorizontally(sheet, elementName, elementMap) {
    var size = this.getElementMapSize(elementMap);
    var rows = size[0];
    var columns = size[1];

    var data = new Array(rows);
    for (var r = 0; r < rows; ++r) {
      data[r] = new Array(columns);
    }

    var names = elementMap[elementName];
    for (var n = 0; n < names.length; ++n) {
      data[0][n + 1] = names[n];
    }
    var row = 1;
    for (var elementValue in elementMap) {
      if (elementValue == elementName) {
        continue;
      }
      data[row][0] = elementValue;
      var values = elementMap[elementValue];
      for (var v = 0; v < values.length; ++v) {
        data[row][v + 1] = values[v];
      }
      row += 1;
    }

    sheet.getRange(1, 1, data.length, data[0].length)
      .setValues(data);
  }

  static getElementMapSize(elementMap) {
    var width = 0;
    var height = 0;

    for (var value in elementMap) {
      var values = elementMap[value];
      width += 1;
      height = Math.max(height, 1 + values.length);
    }

    return [ width, height ];
  }
}
