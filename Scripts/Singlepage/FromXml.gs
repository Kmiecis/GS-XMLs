function fromXmlDocument(document) {
  /*var raw = "" +
    "<tasks>" +
    "<task id='adventure.ce.1.t1' type='hunt' guide='animalhunt'>" +
    "<translation value='adventure.ce.1.t1.name'/>" +
    "<scoring>" +
    "<criterion value='count'/>" +
    "<repetitions value='1'/>" +
    "<pointsPerCompletion value='200'/>" +
    "<limit value='6'/>" +
    "</scoring>" +
    "<animal>" +
    "<speciesId value='vulpes.vulpes'/>" +
    "<speciesId value='capreolus.capreolus'/>" +
    "</animal>" +
    "<game mode='freehunt'/>" +
    "</task>" +
    "<task id='adventure.ce.1.t2' type='hunt' guide='animalhunt'>" +
    "<translation value='adventure.ce.1.t2.name'/>" +
    "<scoring>" +
    "<criterion value='count'/>" +
    "<repetitions value='1'/>" +
    "<pointsPerCompletion value='200'/>" +
    "<limit value='6'/>" +
    "</scoring>" +
    "<animal>" +
    "<gender value='f'/>" +
    "</animal>" +
    "<game mode='freehunt'/>" +
    "</task>" +
    "<task id='prechallenge.ce.1.fh.6' type='hunt' guide='animalhunt'>" +
    "<translation value='prechallenge.ce.1.fh.6.name'/>" +
    "<scoring>" +
    "<criterion value='count'/>" +
    "<repetitions value='1'/>" +
    "<pointsPerCompletion value='200'/>" +
    "<limit value='9'/>" +
    "</scoring>" +
    "<animal>" +
    "<stars min='1' max='3'/>" +
    "</animal>" +
    "</task>" +
    "</tasks>" +
  "";*/
  //var document = XmlService.parse(raw);
  var root = document.getRootElement();
  var map = {};
  this.fillMapWithNames(root, map);
  this.fillMapWithValues(root, map);
  this.parseToSheetVertically(map);
}

function fillMapWithNames(element, map) {
  var elementName = element.getName();

  var elementMap = map[elementName];
  if (elementMap == null) {
    elementMap = map[elementName] = {};
  }
  var names = elementMap[elementName];
  if (names == null) {
    names = elementMap[elementName] = [];
  }

  var attributes = element.getAttributes();
  var children = element.getChildren();

  for (var i = 0; i < attributes.length; ++i) {
    var attribute = attributes[i];
    var attributeName = attribute.getName();
    // Add unique
    if (!names.includes(attributeName)) {
      names.push(attributeName);
    }
  }

  var lastChildName = null;
  var nameOffset = 0;
  for (var i = 0; i < children.length; ++i) {
    var child = children[i];
    var childName = child.getName();
    if (lastChildName != childName) {
      lastChildName = childName;
      nameOffset = 0;
    }
    // Add if out of order
    var nameIndex = names.indexOf(childName);
    if (nameIndex == -1) {
      names.push(childName);
    } else {
      if (names[nameIndex + nameOffset] != childName) {
        names.splice(nameIndex + nameOffset, 0, childName);
      }
    }
    nameOffset++;
  }

  for (var i = 0; i < children.length; ++i) {
    var child = children[i];
    this.fillMapWithNames(child, map);
  }
}

function fillMapWithValues(element, map) {
  var elementName = element.getName();
  var elementMap = map[elementName];

  var names = elementMap[elementName];

  var attributes = element.getAttributes();
  var children = element.getChildren();

  var ref = null;
  var elementMapLength = 0;
  for (var elementRef in elementMap) {
    elementMapLength++;
    if (elementRef == elementName) { // Skip first
      continue;
    }
    ref = elementRef;

    var values = elementMap[elementRef];
    for (var i = 0; i < attributes.length && ref != null; ++i) {
      var attribute = attributes[i];
      var attributeName = attribute.getName();
      var attributeValue = attribute.getValue();
      var nameIndex = names.indexOf(attributeName);
      var currentValue = values[nameIndex];
      if (currentValue != attributeValue) {
        ref = null;
      }
    }
    for (var i = 0; i < children.length && ref != null; ++i) {
      var child = children[i];
      var childName = child.getName();
      var childValue = this.fillMapWithValues(child, map);
      var nameIndex = names.indexOf(childName);
      while (nameIndex < names.length && names[nameIndex] == childName && values[nameIndex] != childValue) {
        nameIndex++;
      }
      if (nameIndex >= names.length || names[nameIndex] != childName || values[nameIndex] != childValue) {
        ref = null;
      }
    }

    if (ref != null) {
      break;
    }
  }

  if (ref == null) {
    ref = "ref" + elementMapLength;
    var values = new Array(names.length);
    for (var i = 0; i < attributes.length; ++i) {
      var attribute = attributes[i];
      var attributeName = attribute.getName();
      var nameIndex = names.indexOf(attributeName);
      if (nameIndex == -1) {
        throw "Unable to find attribute '" + attributeName + "' in element names for '" + elementName + "'";
      }
      if (values[nameIndex] != null) {
        throw "Unable to assign attribute '" + attributeName + "' for element '" + elementName + "' because value already exist";
      }
      var attributeValue = attribute.getValue();
      values[nameIndex] = attributeValue;
    }
    for (var i = 0; i < children.length; ++i) {
      var child = children[i];
      var childName = child.getName();
      var nameIndex = names.indexOf(childName);
      if (nameIndex == -1) {
        throw "Unable to find child '" + childName + "' in element names for '" + elementName + "'";
      }
      while (nameIndex < names.length && values[nameIndex] != null) {
        nameIndex++;
      }
      if (nameIndex >= names.length) {
        throw "Unable to assign child value for '" + childName + "' in element '" + elementName + "' because there is not enough space";
      }
      if (names[nameIndex] != childName) {
        throw "Unable to assign child value for '" + childName + "' in element '" + elementName + "' because there is not enough children: " + names;
      }
      var childValue = this.fillMapWithValues(child, map);
      values[nameIndex] = childValue;
    }
    elementMap[ref] = values;
  }
  return ref;
}

function parseToSheetVertically(map) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  var rowsOffset = 0;
  for (var elementName in map) {
    var elementMap = map[elementName]
    var names = elementMap[elementName];

    var columnsOffset = 0;
    for (var elementRef in elementMap) {
      var values = elementMap[elementRef];

      var rows = new Array(1 + values.length);
      rows[0] = [ elementRef ];
      for (var i = 0; i < values.length; ++i) {
        var value = values[i];
        rows[1 + i] = [ value ];
      }

      sheet.getRange(1 + rowsOffset, 1 + columnsOffset, rows.length)
        .setValues(rows);
      
      columnsOffset += 1;
    }

    rowsOffset += (2 + names.length);
  }
}

function parseToSheetHorizontally(map) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  var rowsOffset = 0;
  for (var elementName in map) {
    var elementMap = map[elementName]
    var names = elementMap[elementName];

    var columnsOffset = 0;
    for (var elementRef in elementMap) {
      var values = elementMap[elementRef];

      var rows = new Array(1 + values.length);
      rows[0] = elementRef;
      for (var i = 0; i < values.length; ++i) {
        var value = values[i];
        rows[1 + i] = value;
      }

      sheet.getRange(1 + rowsOffset, 1 + columnsOffset, 1, rows.length)
        .setValues([ rows ]);
      
      rowsOffset += 1;
    }

    rowsOffset += 1;
  }
}
