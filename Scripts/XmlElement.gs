class XmlElement {
  constructor(name) {
    this.name = name;
    this.attributes = [];
    this.children = [];
  }

  addAttribute(name, value) {
    this.attributes.push([ name, value ]);
  }

  addChild(name, value) {
    this.children.push([ name, value ]);
  }

  getName() {
    return this.name;
  }

  getAttributes() {
    return this.attributes;
  }

  getChildren() {
    return this.children;
  }

  replaceAttribute(name, value) {
    for (var i = 0; i < this.attributes.length; ++i) {
      var attribute = this.attributes[i];
      if (attribute[0] == name) {
        attribute[1] = value;
        break;
      }
    }
  }

  replaceChild(name, oldValue, newValue) {
    for (var i = 0; i < this.children.length; ++i) {
      var child = this.children[i];
      if (child[0] == name && child[1] == oldValue) {
        child[1] = newValue;
        break;
      }
    }
  }

  equals(other) {
    return this.areAttributesEqual(other.getAttributes()) && this.areChildrenEqual(other.getChildren());
  }

  areAttributesEqual(otherAttributes) {
    var thisAttributes = this.attributes;

    if (thisAttributes.length != otherAttributes.length) {
      return false;
    }
    var length = thisAttributes.length;

    var matches = true;
    for (var i = 0; i < length && matches; ++i) {
      var thisAttribute = thisAttributes[i];
      var thisAttributeName = thisAttribute[0];
      var thisAttributeValue = thisAttribute[1];

      var found = false;
      for (var j = 0; j < length && !found; ++j) {
        var otherAttribute = otherAttributes[j];
        var otherAttributeName = otherAttribute[0];
        var otherAttributeValue = otherAttribute[1];

        found = thisAttributeName == otherAttributeName && thisAttributeValue == otherAttributeValue;
      }
      matches = found;
    }

    return matches;
  }

  areChildrenEqual(otherChildren) {
    var thisChildren = this.children;

    if (thisChildren.length != otherChildren.length) {
      return false;
    }
    var length = thisChildren.length;

    var matches = true;
    for (var i = 0; i < length && matches; ++i) {
      var thisChild = thisChildren[i];
      var thisChildName = thisChild[0];
      var thisChildValue = thisChild[1];

      var found = false;
      for (var j = 0; j < length && !found; ++j) {
        var otherChild = otherChildren[j];
        var otherChildName = otherChild[0];
        var otherChildValue = otherChild[1];

        found = thisChildName === otherChildName && thisChildValue === otherChildValue;
      }
      matches = found;
    }

    return matches;
  }
}
