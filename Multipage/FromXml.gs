function fromXmlDocument(document) {
  var root = document.getRootElement();
  var sheetsArray = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var sheetsMap = this.getSheetsMap(sheetsArray);
  this.ensureElement(root, sheetsMap);
  this.ensureElementId(root, sheetsMap);
}

function ensureElement(element, sheetsMap) {
  var elementName = element.getName();

  var sheet = sheetsMap[elementName];
  if (sheet == null) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(elementName);
    sheetsMap[elementName] = sheet;
  }

  var attributes = element.getAttributes();
  var children = element.getChildren();
  
  var rows = new Array(attributes.length + children.length);
  for (var i = 0; i < attributes.length; ++i) {
    var attribute = attributes[i];
    var attributeName = attribute.getName();

    rows[i] = new Array(1);
    rows[i][0] = attributeName;
  }
  for (var i = 0; i < children.length; ++i) {
    var child = children[i];
    var childName = child.getName();

    rows[attributes.length + i] = new Array(1);
    rows[attributes.length + i][0] = childName;
  }

  sheet.getRange(2, 1, rows.length).setValues(rows);

  for (var i = 0; i < children.length; ++i) {
    var child = children[i];
    this.ensureElement(child, sheetsMap);
  }
}

function ensureElementId(element, sheetsMap) {
  var elementName = element.getName();
  var sheet = sheetsMap[elementName];

  var attributes = element.getAttributes();
  var children = element.getChildren();

  var range = sheet.getDataRange();
  var width = range.getWidth();
  var height = range.getHeight();
  var data = range.getValues();

  var c = 1;
  var match = false;
  if (attributes.length > 0) {
    var attribute = attributes[0];
    var attributeValue = attribute.getValue();

    for (; c < width && !match; ++c) {
      var value = data[1][c];
      match = value == attributeValue;
    }
  } else { // children.length > 0
    var child = children[0];
    var childValue = this.ensureElementId(child, sheetsMap);

    for (; c < width && !match; ++c) {
      var value = data[1][c];
      match = value == childValue;
    }
  }
  if (match) {
    --c;
    for (var ai = 0; ai < attributes.length && match; ++ai) {
      var attribute = attributes[ai];
      var attributeValue = attribute.getValue();

      var r = 1 + ai;
      var value = data[r][c];

      match = value == attributeValue;
    }
    for (var ci = 0; ci < children.length && match; ++ci) {
      var child = children[ci];
      var childValue = this.ensureElementId(child, sheetsMap);

      var r = 1 + attributes.length + ci;
      var value = data[r][c];

      match = value == childValue;
    }
  }

  if (match) {
    return data[0][c];
  } else {
    var id = "ref" + c;
    var values = new Array(1 + attributes.length + children.length);
    values[0] = new Array(1);
    values[0][0] = id;
    for (var ai = 0; ai < attributes.length; ++ai) {
      var attribute = attributes[ai];

      var i = 1 + ai;
      values[i] = new Array(1);
      values[i][0] = attribute.getValue();
    }
    for (var ci = 0; ci < children.length; ++ci) {
      var child = children[ci];

      var i = 1 + attributes.length + ci;
      values[i] = new Array(1);
      values[i][0] = this.ensureElementId(child, sheetsMap);
    }
    sheet.getRange(1, 1 + c, values.length).setValues(values);
    return id;
  }
}
