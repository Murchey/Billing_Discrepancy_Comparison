const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron')
const path = require('path')
const fs = require('fs')
const { spawn } = require('child_process')

let backendProcess = null

function startBackend() {
  let appPath
  
  if (app.isPackaged) {
    appPath = path.join(process.resourcesPath, 'app.asar.unpacked')
    console.log('Running in packaged mode, appPath:', appPath)
  } else {
    appPath = __dirname
    console.log('Running in development mode, appPath:', appPath)
  }
  
  const jarPath = path.join(appPath, 'backEnd', 'NewHandleTheBill.jar')
  const javaPath = path.join(appPath, 'backEnd', 'jre', 'bin', 'java.exe')
  
  console.log('Looking for JAR at:', jarPath)
  console.log('Looking for Java at:', javaPath)
  
  try {
    if (!fs.existsSync(jarPath)) {
      console.error('JAR file not found at:', jarPath)
      dialog.showErrorBox('后端启动失败', '无法找到后端JAR文件。\n\n路径：' + jarPath)
      return
    }
    
    console.log('JAR file found, starting backend initialization')
    
    if (fs.existsSync(javaPath)) {
      console.log('Bundled JRE found, trying to start backend')
      try {
        const normalizedJavaPath = path.normalize(javaPath)
        console.log('Normalized Java path:', normalizedJavaPath)
        
        backendProcess = spawn(normalizedJavaPath, ['-jar', jarPath], {
          cwd: path.join(appPath, 'backEnd'),
          stdio: ['pipe', 'pipe', 'pipe'],
          windowsHide: true
        })
        
        backendProcess.stdout.on('data', (data) => {
          console.log('Backend stdout:', data.toString())
        })
        
        backendProcess.stderr.on('data', (data) => {
          console.error('Backend stderr:', data.toString())
        })
        
        backendProcess.on('close', (code) => {
          console.log(`Backend process exited with code ${code}`)
          if (code !== 0) {
            console.error('Backend process exited with non-zero code')
          }
        })
        
        backendProcess.on('error', (err) => {
          console.error('Backend process error:', err)
          dialog.showErrorBox('后端启动失败', '无法启动后端服务。\n\n错误：' + err.message)
        })
        
        console.log('Backend started using bundled JRE')
      } catch (jreError) {
        console.error('Error starting backend with bundled JRE:', jreError)
        dialog.showErrorBox('后端启动失败', '无法启动后端服务。\n\n可能的原因：\n1. 内置JRE损坏\n2. Java运行时环境错误\n\n请检查Java运行时环境。')
      }
    } else {
      console.log('Bundled JRE not found, trying system Java')
      try {
        backendProcess = spawn('java', ['-jar', jarPath], {
          cwd: path.join(appPath, 'backEnd'),
          stdio: ['pipe', 'pipe', 'pipe'],
          windowsHide: true
        })
        
        backendProcess.stdout.on('data', (data) => {
          console.log('Backend stdout:', data.toString())
        })
        
        backendProcess.stderr.on('data', (data) => {
          console.error('Backend stderr:', data.toString())
        })
        
        backendProcess.on('close', (code) => {
          console.log(`Backend process exited with code ${code}`)
          if (code !== 0) {
            console.error('Backend process exited with non-zero code')
          }
        })
        
        backendProcess.on('error', (err) => {
          console.error('Backend process error:', err)
          dialog.showErrorBox('后端启动失败', '无法启动后端服务。\n\n错误：' + err.message)
        })
        
        console.log('Backend started using system Java')
      } catch (systemError) {
        console.error('Error starting backend with system Java:', systemError)
        dialog.showErrorBox('后端启动失败', '无法启动后端服务。\n\n可能的原因：\n1. 系统未安装Java运行时环境\n2. Java未添加到系统PATH\n\n请安装Java 8或更高版本后重试。')
      }
    }
  } catch (error) {
    console.error('Error starting backend:', error)
    dialog.showErrorBox('后端启动失败', '启动后端服务时发生错误：' + error.message)
  }
}

function stopBackend() {
  if (backendProcess) {
    backendProcess.kill()
    console.log('Backend stopped successfully')
  }
}

function createWindow () {
  const mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    title: '账单比对系统',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true
    },
    autoHideMenuBar: true
  })

  mainWindow.loadFile('src/index.html')
}

app.whenReady().then(() => {
  startBackend()
  createWindow()

  app.on('activate', function () {
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})

app.on('window-all-closed', function () {
  stopBackend()
  if (process.platform !== 'darwin') app.quit()
})

app.on('quit', function () {
  stopBackend()
})

ipcMain.handle('read-directory', async (event, dirPath) => {
  try {
    const files = fs.readdirSync(dirPath)
    return { success: true, files: files }
  } catch (error) {
    return { success: false, error: error.message }
  }
})

ipcMain.handle('select-directory', async (event, title) => {
  const result = await dialog.showOpenDialog({
    properties: ['openDirectory'],
    title: title || '选择文件夹'
  })
  
  if (result.canceled) {
    return { success: false, canceled: true }
  }
  
  return { success: true, path: result.filePaths[0] }
})

ipcMain.handle('open-directory', async (event, dirPath) => {
  try {
    await shell.openPath(dirPath)
    return { success: true }
  } catch (error) {
    return { success: false, error: error.message }
  }
})
