const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron')
const path = require('path')
const fs = require('fs')
const { spawn } = require('child_process')

let backendProcess = null

function startBackend() {
  let appPath
  
  // 检测是否在打包环境中运行
  if (app.isPackaged) {
    // 打包环境：使用process.resourcesPath来获取资源目录
    appPath = path.join(process.resourcesPath, 'app.asar.unpacked')
    console.log('Running in packaged mode, appPath:', appPath)
  } else {
    // 开发环境：__dirname就是项目根目录
    appPath = __dirname
    console.log('Running in development mode, appPath:', appPath)
  }
  
  const jarPath = path.join(appPath, 'backEnd', 'NewHandleTheBill.jar')
  const javaPath = path.join(appPath, 'backEnd', 'jre', 'bin', 'java.exe')
  
  console.log('Looking for JAR at:', jarPath)
  console.log('Looking for Java at:', javaPath)
  
  try {
    // 验证路径是否存在
    if (!fs.existsSync(jarPath)) {
      console.error('JAR file not found at:', jarPath)
      dialog.showErrorBox('后端启动失败', '无法找到后端JAR文件。\n\n路径：' + jarPath)
      return
    }
    
    console.log('JAR file found, starting backend initialization')
    
    // 首先尝试使用系统Java，因为它更可靠
    console.log('Trying system Java first')
    try {
      backendProcess = spawn('java', ['-jar', jarPath], {
        cwd: path.join(appPath, 'backEnd'),
        stdio: 'inherit'
      })
      console.log('Backend started using system Java')
    } catch (systemError) {
      console.error('Error starting backend with system Java:', systemError)
      
      // 尝试使用bundled JRE
      if (fs.existsSync(javaPath)) {
        console.log('System Java failed, trying bundled JRE')
        try {
          // 确保路径是绝对路径且格式正确
          const normalizedJavaPath = path.normalize(javaPath)
          console.log('Normalized Java path:', normalizedJavaPath)
          
          backendProcess = spawn(normalizedJavaPath, ['-jar', jarPath], {
            cwd: path.join(appPath, 'backEnd'),
            stdio: 'inherit'
          })
          console.log('Backend started using bundled JRE')
        } catch (jreError) {
          console.error('Error starting backend with bundled JRE:', jreError)
          // 显示详细的错误信息
          dialog.showErrorBox('后端启动失败', '无法启动后端服务。\n\n可能的原因：\n1. 系统未安装Java运行时环境\n2. Java未添加到系统PATH\n3. 内置JRE损坏\n\n请安装Java 8或更高版本后重试。')
        }
      } else {
        console.error('Bundled JRE not found at:', javaPath)
        dialog.showErrorBox('后端启动失败', '无法启动后端服务。\n\n可能的原因：\n1. 系统未安装Java运行时环境\n2. Java未添加到系统PATH\n3. 内置JRE目录不存在\n\n请安装Java 8或更高版本后重试。')
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