const {
  app,
  Menu,
  MenuItem,
  clipboard,
  dialog,
  nativeImage,
  shell,
} = require("electron");
const fs = require("node:fs"),
  path = require("node:path");
const appMenu = require("./appMenu");
const {
  collectPartitionsToClear,
  clearStorageForPartitions,
} = require("../utils/storagePartitions");
const Tray = require("./tray");
const TrayIconChooser = require("../browser/tools/trayIconChooser");
const { SpellCheckProvider } = require("../spellCheckProvider");
const GpuInfoWindow = require("../gpuInfoWindow");
const autoUpdaterModule = require("../autoUpdater");
const {
  writeMigratedConfig,
  MIGRATED_FILE,
} = require("../config/migrateFile");

// Project documentation lives in the repository README.
const DOCUMENTATION_URL = "https://github.com/Brumhilde/outlook-for-linux#readme";

let _Menus_onSpellCheckerLanguageChanged = new WeakMap();
class Menus {
  constructor(window, configGroup, iconPath, connectionManager) {
    this.window = window;
    this.iconPath = iconPath;
    this.configGroup = configGroup;
    this.connectionManager = connectionManager;
    this.allowQuit = false;
    this.gpuInfoWindow = new GpuInfoWindow();
    this.initialize();
  }

  get onSpellCheckerLanguageChanged() {
    return _Menus_onSpellCheckerLanguageChanged.get(this);
  }

  set onSpellCheckerLanguageChanged(value) {
    if (typeof value === "function") {
      _Menus_onSpellCheckerLanguageChanged.set(this, value);
    }
  }

  async quit(clearStorage = false) {
    this.allowQuit = true;

    clearStorage =
      clearStorage &&
      dialog.showMessageBoxSync(this.window, {
        buttons: ["Yes", "No"],
        title: "Quit",
        normalizeAccessKeys: true,
        defaultId: 1,
        cancelId: 1,
        message:
          "Are you sure you want to clear the storage before quitting? If you have clearStorageData set in the config, it will use that configuration.",
        type: "question",
      }) === 0;

    if (clearStorage) {
      // startupConfig, not configGroup: AppConfiguration keeps the parsed
      // config behind a getter, so reading an option straight off the instance
      // is always undefined and silently clears everything (#2860).
      const clearOptions = this.configGroup.startupConfig.clearStorageData;
      await clearStorageForPartitions(
        collectPartitionsToClear(this.configGroup.startupConfig.partition),
        clearOptions,
        "on quit"
      );
    }

    this.window.close();
  }

  open() {
    if (!this.window.isVisible()) {
      this.window.show();
    }

    this.window.focus();
  }

  about() {
    const appInfo = [];
    appInfo.push(`outlook-for-linux@${app.getVersion()}\n`);
    for (const prop in process.versions) {
      if (
        prop === "node" ||
        prop === "v8" ||
        prop === "electron" ||
        prop === "chrome"
      ) {
        appInfo.push(`${prop}: ${process.versions[prop]}`);
      }
    }
    dialog.showMessageBoxSync(this.window, {
      buttons: ["OK"],
      title: "About",
      normalizeAccessKeys: true,
      defaultId: 0,
      cancelId: 0,
      message: appInfo.join("\n"),
      type: "info",
    });
  }

  reload(show = true) {
    if (show) {
      this.window.show();
    }

    this.connectionManager.refresh(true);
  }

  debug() {
    this.window.openDevTools();
  }

  hide() {
    this.window.hide();
  }

  initialize() {
    const menu = appMenu(this);

    if (this.configGroup.startupConfig.menubar === "hidden") {
      this.window.removeMenu();
    } else {
      this.window.setMenu(Menu.buildFromTemplate([menu]));
    }

    this.initializeEventHandlers();

    if (this.configGroup.startupConfig.trayIconEnabled) {
      this.tray = new Tray(
        this.window,
        menu.submenu,
        this.iconPath,
        this.configGroup.startupConfig
      );
      this.tray.initialize();
    }
    this.spellCheckProvider = new SpellCheckProvider(this.window);
  }

  initializeEventHandlers() {
    app.on("before-quit", () => this.onBeforeQuit());
    this.window.on("close", (event) => this.onClose(event));
    this.window.webContents.on("context-menu", assignContextMenuHandler(this));
  }

  onBeforeQuit() {
    console.debug("before-quit");
    this.allowQuit = true;
  }

  onClose(event) {
    console.debug("window close");
    if (!this.allowQuit && !this.configGroup.startupConfig.closeAppOnCross) {
      event.preventDefault();
      if (this.configGroup.startupConfig.minimizeOnClose) {
        this.window.minimize();
      } else {
        this.hide();
      }
    } else {
      this.tray?.close();
      this.window.webContents.session.flushStorageData();
    }
  }

  chooseAppIcon() {
    const result = dialog.showOpenDialogSync(this.window, {
      title: "Choose App Icon",
      filters: [{ name: "Images", extensions: ["png"] }],
      properties: ["openFile"],
    });
    if (result && result.length > 0) {
      const selectedPath = result[0];
      if (nativeImage.createFromPath(selectedPath).isEmpty()) {
        dialog.showMessageBoxSync(this.window, {
          type: "error",
          title: "Choose App Icon",
          message: "That file could not be read as an image.",
          detail: selectedPath,
        });
        return;
      }
      this.configGroup.startupConfig.appIcon = selectedPath;
      this.configGroup.legacyConfigStore.set("appIcon", selectedPath);
      this.tray?.setBaseIconPath(selectedPath);
      this.#updateWindowIcon(selectedPath);
      this.updateMenu();
    }
  }

  resetAppIcon() {
    this.configGroup.startupConfig.appIcon = "";
    this.configGroup.legacyConfigStore.set("appIcon", "");
    const iconChooser = new TrayIconChooser(this.configGroup.startupConfig);
    const iconPath = iconChooser.getFile();
    this.tray?.setBaseIconPath(iconPath);
    this.#updateWindowIcon(iconPath);
    this.updateMenu();
  }

  #updateWindowIcon(iconPath) {
    this.window.setIcon(nativeImage.createFromPath(iconPath));
    if (!app.dock) return;
    // The tray asset is 16px on macOS but the dock needs >=128px, so the
    // default is resolved separately here, exactly as startup does it.
    const custom = this.configGroup.startupConfig.appIcon?.trim();
    const dockIconPath = custom
      ? custom
      : path.join(this.configGroup.startupConfig.appPath, "assets/icons/icon-256x256.png");
    const dockIcon = nativeImage.createFromPath(dockIconPath);
    app.dock.setIcon(
      dockIcon.getSize().width < 128
        ? dockIcon.resize({ width: 128, height: 128 })
        : dockIcon,
    );
  }

  async showMigratedConfig() {
    const result = writeMigratedConfig(this.configGroup.configPath);
    const TITLE = "Updated Config";

    // Everything below prints option NAMES only. A config file carries broker
    // URLs, certificate paths and service URLs, and this text goes on screen
    // and into the log.
    const plain = {
      "no-config": "There is no config.json yet, so there is nothing to update.",
      "nothing-to-migrate": "Your config already uses the current option names.",
      "invalid-json":
        "config.json is not a readable JSON object, so it cannot be updated. Fix the file and try again.",
    };
    if (plain[result.status]) {
      await dialog.showMessageBox(this.window, {
        type: "info",
        title: TITLE,
        message: plain[result.status],
      });
      return;
    }

    if (result.status === "blocked") {
      await dialog.showMessageBox(this.window, {
        type: "warning",
        title: TITLE,
        message: "Nothing could be updated automatically.",
        detail: [
          "These options have new names, but the namespace they move into is",
          "already set to something that is not a group of settings:",
          ...result.skipped.map((name) => `  • ${name}`),
        ].join("\n"),
      });
      return;
    }

    if (result.status === "write-failed") {
      await dialog.showMessageBox(this.window, {
        type: "error",
        title: TITLE,
        message: "The updated config could not be written.",
        detail: result.error,
      });
      return;
    }

    const count = result.renamed.length;
    const detail = [
      `${count} ${count === 1 ? "option has" : "options have"} a new name:`,
      ...result.renamed.map((name) => `  • ${name}`),
      "",
      "Your config.json is untouched. Review the copy, then rename it over",
      "config.json when you are happy with it.",
      ...(result.skipped.length
        ? ["", "Left alone, because their namespace is already set to something else:",
           ...result.skipped.map((name) => `  • ${name}`)]
        : []),
      ...(result.warnings.length
        ? ["", "The copy needs a look first:", ...result.warnings.map((w) => `  • ${w}`)]
        : []),
    ].join("\n");

    const { response } = await dialog.showMessageBox(this.window, {
      type: result.warnings.length ? "warning" : "info",
      title: TITLE,
      // The directory too, so someone whose desktop has no handler for .json
      // can still find the file after "Open It" does nothing.
      message: `Written to ${MIGRATED_FILE} in ${result.dir}`,
      detail,
      buttons: ["Open It", "Close"],
      defaultId: 0,
      cancelId: 1,
    });
    if (response !== 0) return;

    // Electron resolves with a non-empty error string when the OS has no
    // handler; matching app/downloadManager, log only that it failed, since
    // the string can carry the path.
    const openError = await shell.openPath(result.file);
    if (openError) {
      console.warn("[Config] Could not open the updated config", { failed: true });
    }
  }

  // Opens config.json in the user's default editor. The file usually does not
  // exist — startup logs "No config file found ... using default values" — and
  // shell.openPath on a missing path just returns an error string, so an empty
  // stub is written first. That also gives the user something to edit rather
  // than an empty buffer they have to save into the right place themselves.
  async openConfigFile() {
    const configFile = this.configGroup.configFilePath;
    try {
      if (!fs.existsSync(configFile)) {
        fs.mkdirSync(path.dirname(configFile), { recursive: true });
        fs.writeFileSync(configFile, "{}\n");
      }
    } catch (err) {
      this.#reportOpenFailure("config file", configFile, err.message);
      return;
    }
    await this.#openPath("config file", configFile);
  }

  async openConfigFolder() {
    await this.#openPath("config folder", this.configGroup.configPath);
  }

  // shell.openPath resolves to an empty string on success and to the reason as
  // a string on failure — it does not reject — so the result has to be checked.
  async #openPath(what, target) {
    let reason;
    try {
      reason = await shell.openPath(target);
    } catch (err) {
      reason = err.message;
    }
    if (reason) {
      this.#reportOpenFailure(what, target, reason);
    }
  }

  #reportOpenFailure(what, target, reason) {
    console.error(`Failed to open ${what} at ${target}: ${reason}`);
    dialog.showMessageBoxSync(this.window, {
      message: `Could not open the ${what}.\n\n${target}\n\n${reason}`,
      title: "Open config",
      type: "error",
    });
  }

  updateMenu() {
    const menu = appMenu(this);
    if (this.configGroup.startupConfig.menubar !== "hidden") {
      this.window.setMenu(Menu.buildFromTemplate([menu]));
    }
    this.tray?.setContextMenu(menu.submenu);

    // Notify renderer process of config changes that affect renderer behavior
    // This allows menu toggles to take effect immediately without restart
    this.window.webContents.send("config-changed", {
      disableNotifications: this.configGroup.startupConfig.disableNotifications,
      disableNotificationSound: this.configGroup.startupConfig.disableNotificationSound,
      disableNotificationWindowFlash: this.configGroup.startupConfig.disableNotificationWindowFlash,
      disableBadgeCount: this.configGroup.startupConfig.disableBadgeCount,
      defaultNotificationUrgency: this.configGroup.startupConfig.defaultNotificationUrgency,
      appIcon: this.configGroup.startupConfig.appIcon,
    });
  }

  toggleDisableNotifications() {
    this.configGroup.startupConfig.disableNotifications =
      !this.configGroup.startupConfig.disableNotifications;
    this.configGroup.legacyConfigStore.set(
      "disableNotifications",
      this.configGroup.startupConfig.disableNotifications
    );
    this.updateMenu();
  }

  toggleDisableNotificationSound() {
    this.configGroup.startupConfig.disableNotificationSound =
      !this.configGroup.startupConfig.disableNotificationSound;
    this.configGroup.legacyConfigStore.set(
      "disableNotificationSound",
      this.configGroup.startupConfig.disableNotificationSound
    );
    this.updateMenu();
  }

  toggleDisableNotificationWindowFlash() {
    this.configGroup.startupConfig.disableNotificationWindowFlash =
      !this.configGroup.startupConfig.disableNotificationWindowFlash;
    this.configGroup.legacyConfigStore.set(
      "disableNotificationWindowFlash",
      this.configGroup.startupConfig.disableNotificationWindowFlash
    );
    this.updateMenu();
  }

  toggleDisableBadgeCount() {
    this.configGroup.startupConfig.disableBadgeCount =
      !this.configGroup.startupConfig.disableBadgeCount;
    this.configGroup.legacyConfigStore.set(
      "disableBadgeCount",
      this.configGroup.startupConfig.disableBadgeCount
    );
    this.updateMenu();
  }

  setNotificationUrgency(value) {
    this.configGroup.startupConfig.defaultNotificationUrgency = value;
    this.configGroup.legacyConfigStore.set("defaultNotificationUrgency", value);
    this.updateMenu();
  }

  showDocumentation() {
    shell.openExternal(DOCUMENTATION_URL);
  }

  showGpuInfo() {
    this.gpuInfoWindow.show();
  }

  checkForUpdates() {
    autoUpdaterModule.checkForUpdates();
  }
}

function assignContextMenuHandler(menus) {
  return (_event, params) => {
    const menu = new Menu();

    assignReplaceWordHandler(params, menu, menus);
    assignAddToDictionaryHandler(params, menu, menus);

    if (menu.items.length > 0) {
      menu.popup();
    }
  };
}

function assignReplaceWordHandler(params, menu, menus) {
  for (const suggestion of params.dictionarySuggestions) {
    menu.append(
      new MenuItem({
        label: suggestion,
        click: () => menus.window.webContents.replaceMisspelling(suggestion),
      })
    );
  }
}

function assignAddToDictionaryHandler(params, menu, menus) {
  if (params.misspelledWord) {
    menu.append(
      new MenuItem({
        label: "Add to dictionary",
        click: () =>
          menus.window.webContents.session.addWordToSpellCheckerDictionary(
            params.misspelledWord
          ),
      })
    );

    menu.append(
      new MenuItem({
        type: "separator",
      })
    );
  }

  addTextEditMenuItems(params, menu, menus);
}

function addTextEditMenuItems(params, menu, menus) {
  if (params.isEditable) {
    buildEditContextMenu(menu, menus);
  } else if (params.linkURL !== "") {
    menu.append(
      new MenuItem({
        label: "Copy",
        click: () => clipboard.writeText(params.linkURL),
      })
    );
  }
}

function buildEditContextMenu(menu, menus) {
  menu.append(
    new MenuItem({
      role: "cut",
    })
  );

  menu.append(
    new MenuItem({
      role: "copy",
    })
  );

  menu.append(
    new MenuItem({
      role: "paste",
    })
  );

  addSpellCheckMenuItems(menu, menus);
}

function addSpellCheckMenuItems(menu, menus) {
  menu.append(
    new MenuItem({
      type: "separator",
    })
  );

  menu.append(
    new MenuItem({
      label: "Writing Languages",
      submenu: createSpellCheckLanguagesMenu(menus),
    })
  );
}

function createSpellCheckLanguagesMenu(menus) {
  const activeLanguages =
    menus.window.webContents.session.getSpellCheckerLanguages();
  const splChkMenu = new Menu();
  for (const group of menus.spellCheckProvider.supportedListByGroup) {
    const subMenu = new Menu();
    splChkMenu.append(
      new MenuItem({
        label: group.key,
        submenu: subMenu,
      })
    );
    for (const language of group.list) {
      subMenu.append(createLanguageMenuItem(language, activeLanguages, menus));
    }
  }

  createSpellCheckLanguagesNoneMenuEntry(splChkMenu, menus);

  return splChkMenu;
}

function createSpellCheckLanguagesNoneMenuEntry(menu, menus) {
  menu.append(
    new MenuItem({
      type: "separator",
    })
  );
  menu.append(
    new MenuItem({
      label: "None",
      click: () => chooseLanguage(null, menus),
    })
  );
}

function createLanguageMenuItem(language, activeLanguages, menus) {
  return new MenuItem({
    label: language.language,
    type: "checkbox",
    id: language.code,
    checked: activeLanguages.includes(language.code),
    click: (menuItem) => chooseLanguage(menuItem, menus),
  });
}

function chooseLanguage(item, menus) {
  const activeLanguages =
    menus.window.webContents.session.getSpellCheckerLanguages();
  if (item) {
    if (item.checked) {
      addToList(activeLanguages, item.id);
    } else {
      removeFromList(activeLanguages, item.id);
    }
  }

  const changes = menus.spellCheckProvider.setLanguages(
    item ? activeLanguages : []
  );

  if (menus.onSpellCheckerLanguageChanged) {
    menus.onSpellCheckerLanguageChanged(changes);
  }
}

function removeFromList(list, item) {
  const itemIndex = list.findIndex((l) => l === item);
  if (itemIndex >= 0) {
    list.splice(itemIndex, 1);
  }

  return list;
}

function addToList(list, item) {
  const itemIndex = list.findIndex((l) => l === item);
  if (itemIndex < 0) {
    list.push(item);
  }

  return list;
}

exports = module.exports = Menus;
