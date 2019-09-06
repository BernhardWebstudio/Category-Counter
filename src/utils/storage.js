class Storage {
  constructor() {
    this.standalone = true;
  }

  setStandalone(standalone) {
    this.standalone = standalone;
  }

  setItem(key, value) {
    if (this.standalone) {
      localStorage.setItem(key, value);
    } else {
      Office.context.document.settings.set(key, value);
    }
  }

  getItem(key) {
    if (this.standalone) {
      localStorage.getItem(key);
    } else {
      Office.context.document.settings.get(key);
    }
  }
}

export default new Storage();
