class ArrayUtils {
  static areEqualsUnordered(leftArray, rightArray, comparator) {
    var matches = leftArray.length == rightArray.length;

    if (matches) {
      var checkArray = new Array(rightArray.length);
      for (var i = 0; i < checkArray.length; ++i) {
        checkArray[i] = false;
      }

      for (var i = 0; i < leftArray.length && matches; ++i) {
        var leftItem = leftArray[i];
        var found = false;
        for (var j = 0; j < rightArray.length && !found; ++j) {
          var rightItem = rightArray[j];
          found = comparator(leftItem, rightItem) && !checkArray[j];
          if (found) {
            checkArray[j] = true;
          }
        }
        matches = found;
      }
    }

    return matches;
  }
}
