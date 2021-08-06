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

  areElementsEqual(leftElement, rightElement) {
    return (
      leftElement[0] === rightElement[0] &&
      leftElement[1] === rightElement[1]
    );
  }

  equals(other) {
    return (
      ArrayUtils.areEqualsUnordered(this.getAttributes(), other.getAttributes(), this.areElementsEqual) &&
      ArrayUtils.areEqualsUnordered(this.getChildren(), other.getChildren(), this.areElementsEqual)
    );
  }
}
