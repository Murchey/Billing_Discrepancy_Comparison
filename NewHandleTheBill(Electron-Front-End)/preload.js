const { contextBridge, ipcRenderer } = require('electron')

window.addEventListener('DOMContentLoaded', () => {
  const replaceText = (selector, text) => {
    const element = document.getElementById(selector)
    if (element) element.innerText = text
  }

  for (const type of ['chrome', 'node', 'electron']) {
    replaceText(`${type}-version`, process.versions[type])
  }
})

contextBridge.exposeInMainWorld('electronAPI', {
  readDirectory: (dirPath) => ipcRenderer.invoke('read-directory', dirPath),
  selectDirectory: (title) => ipcRenderer.invoke('select-directory', title),
  openDirectory: (dirPath) => ipcRenderer.invoke('open-directory', dirPath)
})