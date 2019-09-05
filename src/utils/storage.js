class Storage {
  standalone = false

  constructor() {
    this.checkStandalone();
  }

  checkStandalone() {
    try {
      this.standalone = info.host === Office.HostType.Outlook;
    } catch (e) {
      this.standalone = false;
    }
  }

  setItem(key, value) {
    this.checkStandalone() // check again because we could have been instantiated too early
    if (this.standalone) {
      localStorage.setItem(key, value);
    } else {
      Office.context.document.settings.set(key, value);
    }
  }

  getItem(key) {
    this.checkStandalone() // check again because we could have been instantiated too early
    if (this.standalone) {
      localStorage.getItem(key);
    } else {
      Office.context.document.settings.get(key);
    }
  }
}


