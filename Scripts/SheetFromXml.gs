class SheetFromXml {
  static fillSheet(sheet, document) {
    var rootXml = document.getRootElement();
    var rootName = rootXml.getName();
    var rootValue = "ref0";

    var xmlsMap = this.createXmlsMap(rootXml, rootValue);
    var invXmlsMap = this.createInvXmlsMap(xmlsMap);
    var depthStack = this.createDepthStack(xmlsMap, rootName, rootValue);

    this.trimXmlsMap(xmlsMap, invXmlsMap, depthStack);

    var elementsMap = this.createElementsMap(xmlsMap);
    this.parseToSheet(sheet, elementsMap);
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

  static parseToSheet(sheet, elementsMap) {
    this.parseToSheetVertically(sheet, elementsMap);
  }

  static parseToSheetVertically(sheet, elementsMap) {
    var size = this.getElementsMapVerticalSize(elementsMap);
    var rows = size[0];
    var columns = size[1];

    var data = new Array(rows);
    for (var r = 0; r < rows; ++r) {
      data[r] = new Array(columns);
    }

    var rowsOffset = 0;
    for (var elementName in elementsMap) {
      var elementMap = elementsMap[elementName]
      var names = elementMap[elementName];

      var columnsOffset = 1;
      for (var elementValue in elementMap) {
        var values = elementMap[elementValue];

        var currentColumnsOffset = (elementValue == elementName) ? 0 : columnsOffset;

        data[rowsOffset][currentColumnsOffset] = elementValue;
        for (var i = 0; i < values.length; ++i) {
          var value = values[i];
          data[1 + rowsOffset + i][currentColumnsOffset] = value;
        }

        columnsOffset += 1;
      }

      rowsOffset += (1 + names.length + 1);
    }

    sheet.getRange(1, 1, data.length, data[0].length)
      .setValues(data);
  }

  static getElementsMapVerticalSize(elementsMap) {
    var rows = 0;
    var columns = 0;

    for (var name in elementsMap) {
      var elementMap = elementsMap[name];

      var values = 0;
      for (var value in elementMap) {
        values += 1;
      }
      columns = Math.max(columns, values);

      var names = elementMap[name];
      rows += 1 + names.length + 1;
    }
    rows -= 1;

    return [ rows, columns ];
  }

  static parseToSheetHorizontally(sheet, elementsMap) {
    var size = this.getElementsMapHorizontalSize(elementsMap);
    var rows = size[0];
    var columns = size[1];

    var data = new Array(rows);
    for (var r = 0; r < rows; ++r) {
      data[r] = new Array(columns);
    }

    var namesRowsOffset = 0;
    var rowsOffset = 1;
    for (var elementName in elementsMap) {
      var elementMap = elementsMap[elementName]

      for (var elementValue in elementMap) {
        var values = elementMap[elementValue];

        var currentRowsOffset = (elementValue == elementName) ? namesRowsOffset : rowsOffset;

        data[currentRowsOffset][0] = elementValue;
        for (var i = 0; i < values.length; ++i) {
          var value = values[i];
          data[currentRowsOffset][i + 1] = value;
        }

        rowsOffset += 1;
      }
      
      namesRowsOffset = rowsOffset;
      rowsOffset += 1;
    }

    sheet.getRange(1, 1, data.length, data[0].length)
      .setValues(data);
  }

  static getElementsMapHorizontalSize(elementsMap) {
    var rows = 0;
    var columns = 0;

    for (var name in elementsMap) {
      var elementMap = elementsMap[name];

      var names = elementMap[name];
      columns = Math.max(columns, 1 + names.length);

      rows += 1;
      for (var value in elementMap) {
        rows += 1;
      }
    }
    rows -= 1;

    return [ rows, columns ];
  }
}
