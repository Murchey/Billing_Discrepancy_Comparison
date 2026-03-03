function toggleTheme() {
    const themeNow = document.documentElement.getAttribute('data-bs-theme');
    const newTheme = themeNow === 'light' ? 'dark' : 'light';
    document.documentElement.setAttribute('data-bs-theme', newTheme);
}

async function refreshFileLists() {
    const handleDirPath = document.getElementById('handleDirPath').value;
    const standardDirPath = document.getElementById('standardDirPath').value;
    
    if (!handleDirPath || !standardDirPath) {
        alert('请先设置待核对账单和标准账单的文件夹路径！');
        return;
    }
    
    try {
        const handleResult = await window.electronAPI.readDirectory(handleDirPath);
        const standardResult = await window.electronAPI.readDirectory(standardDirPath);
        
        if (handleResult.success && standardResult.success) {
            updateFileList('handle', handleResult.files);
            updateFileList('standard', standardResult.files);
            console.log('文件列表刷新成功！');
        } else {
            console.error('读取文件夹失败：', handleResult.error || standardResult.error);
            alert('读取文件夹失败：' + (handleResult.error || standardResult.error));
        }
    } catch (error) {
        console.error('读取文件夹时发生错误：', error);
        alert('读取文件夹时发生错误：' + error.message);
    }
}

function updateFileList(type, files) {
    const listGroup = type === 'handle' 
        ? document.querySelector('#cmpOperationPage .row .col:first-child .list-group')
        : document.querySelector('#cmpOperationPage .row .col:last-child .list-group');
    
    if (listGroup) {
        listGroup.innerHTML = '';
        files.forEach(file => {
            const listItem = document.createElement('li');
            listItem.className = 'list-group-item';
            listItem.textContent = file;
            listGroup.appendChild(listItem);
        });
    }
}

async function selectHandleDirectory() {
    try {
        const result = await window.electronAPI.selectDirectory();
        if (result.success) {
            const handleDirPathInput = document.getElementById('handleDirPath');
            if (handleDirPathInput) {
                handleDirPathInput.value = result.path;
                console.log('待核对账单文件夹路径已设置：', result.path);
            }
        } else if (!result.canceled) {
            alert('选择文件夹失败！');
        }
    } catch (error) {
        console.error('选择文件夹时发生错误：', error);
        alert('选择文件夹时发生错误：' + error.message);
    }
}

async function selectStandardDirectory() {
    try {
        const result = await window.electronAPI.selectDirectory();
        if (result.success) {
            const standardDirPathInput = document.getElementById('standardDirPath');
            if (standardDirPathInput) {
                standardDirPathInput.value = result.path;
                console.log('标准账单文件夹路径已设置：', result.path);
            }
        } else if (!result.canceled) {
            alert('选择文件夹失败！');
        }
    } catch (error) {
        console.error('选择文件夹时发生错误：', error);
        alert('选择文件夹时发生错误：' + error.message);
    }
}

async function openResultDirectory() {
    try {
        const resultDirPathInput = document.getElementById('resultDirPath');
        if (!resultDirPathInput || !resultDirPathInput.value) {
            alert('请先设置对比结果存放路径！');
            return;
        }
        
        const result = await window.electronAPI.openDirectory(resultDirPathInput.value);
        if (!result.success) {
            console.error('打开文件夹失败：', result.error);
            alert('打开文件夹失败：' + result.error);
        }
    } catch (error) {
        console.error('打开文件夹时发生错误：', error);
        alert('打开文件夹时发生错误：' + error.message);
    }
}

async function selectResultDirectory() {
    try {
        const result = await window.electronAPI.selectDirectory();
        if (result.success) {
            const resultDirPathInput = document.getElementById('resultDirPath');
            if (resultDirPathInput) {
                resultDirPathInput.value = result.path;
                console.log('对比结果存放路径已设置：', result.path);
            }
        } else if (!result.canceled) {
            alert('选择文件夹失败！');
        }
    } catch (error) {
        console.error('选择文件夹时发生错误：', error);
        alert('选择文件夹时发生错误：' + error.message);
    }
}

let progressInterval = null;

async function startComparison() {
    try {
        // 获取所有必要的参数
        const handleDirPath = document.getElementById('handleDirPath').value;
        const standardDirPath = document.getElementById('standardDirPath').value;
        const handleHeadRow = document.getElementById('handleHeadRow').value;
        const handleNameCol = document.getElementById('handleNameCol').value;
        const handleMoneyCol = document.getElementById('handleMoneyCol').value;
        const standardHeadRow = document.getElementById('standardHeadRow').value;
        const standardNameCol = document.getElementById('standardNameCol').value;
        const standardMoneyCol = document.getElementById('standardMoneyCol').value;
        let resultDirPath = document.getElementById('resultDirPath').value;
        
        // 验证参数
        if (!handleDirPath || !standardDirPath || !handleHeadRow || !handleNameCol || !handleMoneyCol || 
            !standardHeadRow || !standardNameCol || !standardMoneyCol) {
            alert('请完成所有设置项！');
            return;
        }
        
        // 如果没有设置结果存放路径，提示用户选择
        if (!resultDirPath) {
            const result = await window.electronAPI.selectDirectory('选择存放比对结果的文件夹');
            if (result.success) {
                resultDirPath = result.path;
                const resultDirPathInput = document.getElementById('resultDirPath');
                if (resultDirPathInput) {
                    resultDirPathInput.value = result.path;
                    console.log('对比结果存放路径已设置：', result.path);
                }
            } else {
                // 用户取消选择
                return;
            }
        }
        
        // 重置进度条
        updateProgressBar(0);
        
        // 构建API请求URL
        const apiUrl = `http://localhost:8090/api/startCmp?needHandleFilePath=${encodeURIComponent(handleDirPath)}&standardFilePath=${encodeURIComponent(standardDirPath)}&handleHeadRow=${handleHeadRow}&handleNameCol=${handleNameCol}&handleMoneyCol=${handleMoneyCol}&standardHeadRow=${standardHeadRow}&standardNameCol=${standardNameCol}&standardMoneyCol=${standardMoneyCol}&resultPath=${encodeURIComponent(resultDirPath)}`;
        
        // 开始轮询进度
        startProgressPolling();
        
        // 调用后端API
        const response = await fetch(apiUrl);
        const success = await response.json();
        
        // 停止轮询
        stopProgressPolling();
        
        // 确保进度条显示100%
        updateProgressBar(100);
        
        // 延迟显示提示，确保用户看到进度条显示100%
        setTimeout(() => {
            if (success) {
                alert('比对完成！');
            } else {
                alert('比对失败，请检查设置！');
            }
        }, 500);
    } catch (error) {
        console.error('开始比对时发生错误：', error);
        alert('开始比对时发生错误：' + error.message);
        stopProgressPolling();
    }
}

function startProgressPolling() {
    // 清除之前的轮询
    stopProgressPolling();
    
    // 每200毫秒获取一次进度，提高更新频率
    progressInterval = setInterval(async () => {
        try {
            const response = await fetch('http://localhost:8090/api/getProgress');
            const progress = await response.json();
            updateProgressBar(progress * 100);
        } catch (error) {
            console.error('获取进度时发生错误：', error);
        }
    }, 200);
}

function stopProgressPolling() {
    if (progressInterval) {
        clearInterval(progressInterval);
        progressInterval = null;
    }
}

function updateProgressBar(percentage) {
    const progressBar = document.getElementById('progressBar');
    if (progressBar) {
        const percent = Math.min(100, Math.max(0, percentage));
        progressBar.style.width = `${percent}%`;
        progressBar.setAttribute('aria-valuenow', percent);
        progressBar.textContent = `${Math.round(percent)}%`;
    }
}
 
// 确保DOM加载完成后执行
window.addEventListener('DOMContentLoaded', function() {
    // 检查按钮是否存在
    const showModeChangeBtn = document.getElementById("showModeChangeBtn");
    if (showModeChangeBtn) {
        console.log("按钮找到成功！");
        showModeChangeBtn.addEventListener("click", toggleTheme);
    } else {
        console.log("按钮未找到！");
    }
    
    // 检查刷新列表按钮是否存在
    const refreshListBtn = document.getElementById("refreshListBtn");
    if (refreshListBtn) {
        console.log("刷新列表按钮找到成功！");
        refreshListBtn.addEventListener("click", refreshFileLists);
    } else {
        console.log("刷新列表按钮未找到！");
    }
    
    // 检查选择待核对账单文件夹按钮是否存在
    const selectHandleDirBtn = document.getElementById("selectHandleDirBtn");
    if (selectHandleDirBtn) {
        console.log("选择待核对账单文件夹按钮找到成功！");
        selectHandleDirBtn.addEventListener("click", selectHandleDirectory);
    } else {
        console.log("选择待核对账单文件夹按钮未找到！");
    }
    
    // 检查选择标准账单文件夹按钮是否存在
    const selectStandardDirBtn = document.getElementById("selectStandardDirBtn");
    if (selectStandardDirBtn) {
        console.log("选择标准账单文件夹按钮找到成功！");
        selectStandardDirBtn.addEventListener("click", selectStandardDirectory);
    } else {
        console.log("选择标准账单文件夹按钮未找到！");
    }
    
    // 检查打开对比结果文件夹按钮是否存在
    const openResultDirBtn = document.getElementById("openResultDirBtn");
    if (openResultDirBtn) {
        console.log("打开对比结果文件夹按钮找到成功！");
        openResultDirBtn.addEventListener("click", async function() {
            const resultDirPathInput = document.getElementById('resultDirPath');
            if (resultDirPathInput && resultDirPathInput.value) {
                // 如果已有路径，打开文件夹
                await openResultDirectory();
            } else {
                // 如果没有路径，选择文件夹并设置
                await selectResultDirectory();
            }
        });
    } else {
        console.log("打开对比结果文件夹按钮未找到！");
    }
    
    // 检查开始比对按钮是否存在
    const startCmpBtn = document.getElementById("startCmpBtn");
    if (startCmpBtn) {
        console.log("开始比对按钮找到成功！");
        startCmpBtn.addEventListener("click", startComparison);
    } else {
        console.log("开始比对按钮未找到！");
    }
});