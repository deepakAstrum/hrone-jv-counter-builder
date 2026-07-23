import path from "node:path";
import {
  app,
  BrowserWindow,
  dialog,
  session,
} from "electron";

let mainWindow;
const smokeTest = process.argv.includes("--smoke-test");

function createWindow() {
  mainWindow = new BrowserWindow({
    title: "HROne Counter Builder",
    width: 1500,
    height: 980,
    minWidth: 1050,
    minHeight: 720,
    backgroundColor: "#f6f6f6",
    autoHideMenuBar: true,
    show: false,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      devTools: false,
    },
  });

  mainWindow.removeMenu();
  if (!smokeTest) {
    mainWindow.once("ready-to-show", () => mainWindow.show());
  }
  mainWindow.webContents.once(
    "did-fail-load",
    (_event, errorCode, errorDescription) => {
      console.error(`Renderer failed to load (${errorCode}): ${errorDescription}`);
      app.exit(1);
    },
  );
  if (smokeTest) {
    mainWindow.webContents.once("did-finish-load", () => {
      setTimeout(() => app.quit(), 100);
    });
  }
  mainWindow.webContents.setWindowOpenHandler(() => ({ action: "deny" }));
  mainWindow.webContents.on("will-navigate", (event, url) => {
    if (!url.startsWith("file://")) event.preventDefault();
  });

  const renderer = path.join(
    app.getAppPath(),
    "dist",
    "renderer",
    "index.html",
  );
  void mainWindow.loadFile(renderer);
}

app.whenReady().then(() => {
  session.defaultSession.setPermissionRequestHandler(
    (_webContents, _permission, callback) => callback(false),
  );
  session.defaultSession.webRequest.onBeforeRequest(
    { urls: ["http://*/*", "https://*/*"] },
    (_details, callback) => callback({ cancel: true }),
  );
  session.defaultSession.on("will-download", (_event, item) => {
    const destination = dialog.showSaveDialogSync(mainWindow, {
      title: "Save countered workbook",
      defaultPath: path.join(app.getPath("documents"), item.getFilename()),
      filters: [
        { name: "Excel workbook", extensions: ["xlsx"] },
        { name: "All files", extensions: ["*"] },
      ],
    });

    if (destination) item.setSavePath(destination);
    else item.cancel();
  });

  createWindow();

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on("window-all-closed", () => {
  app.quit();
});
