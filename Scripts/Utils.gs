function createArray(width, height, depth) {
  var widthArray = new Array(width);
  for (var x = 0; x < width; ++x) {
    var heightArray = new Array(height);
    for (var y = 0; y < height; ++y) {
      var depthArray = new Array(depth);
      heightArray[y] = depthArray;
    }
    widthArray[x] = heightArray;
  }
  return widthArray;
}

function createArray(width, height) {
  var widthArray = new Array(width);
  for (var x = 0; x < width; ++x) {
    var heightArray = new Array(height);
    widthArray[x] = heightArray;
  }
  return widthArray;
}

function createArray(length) {
  var lengthArray = new Array(length);
  return lengthArray;
}

function createSheetsMap(sheetsArray) {
  var result = {};
  for (var i = 0; i < sheetsArray.length; ++i) {
    var sheet = sheetsArray[i];
    result[sheet.getName()] = sheet;
  }
  return result;
}
