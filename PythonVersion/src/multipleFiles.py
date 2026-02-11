# _*_coding : UTF-8 _*_
# @Time : 2026/2/6 20:53
# @Author : Murchey
# @File : multipleFiles
# @Project : python

import sys
from pathlib import Path

def get_application_dir():
    """获取应用程序的根目录，无论是脚本还是打包后的exe"""
    if getattr(sys, 'frozen', False):
        # 打包后的exe模式
        return Path(sys.executable).resolve().parent
    else:
        # 脚本模式
        return Path(__file__).resolve().parent

def newFolders():
    app_dir = get_application_dir()
    standardFileFolder = app_dir / '标准表格'
    standardFileFolder.mkdir(parents=True, exist_ok=True)
    handleFileFolder = app_dir / '需核对表格'
    handleFileFolder.mkdir(parents=True, exist_ok=True)
    dataFileFolder = app_dir / '已提取数据'
    dataFileFolder.mkdir(parents=True, exist_ok=True)
    resFileFolder = app_dir / '比对结果'
    resFileFolder.mkdir(parents=True, exist_ok=True)

def getFilesNames(subFolderName):
    app_dir = get_application_dir()
    subFolderPath = app_dir / f"{subFolderName}"
    # 确保文件夹存在
    subFolderPath.mkdir(parents=True, exist_ok=True)
    # 打印调试信息
    print(f"获取文件夹: {subFolderPath}")
    print(f"文件夹存在: {subFolderPath.exists()}")
    # 转换为列表并返回
    file_names = list(p for p in subFolderPath.glob('*.xlsx'))
    print(f"找到文件: {[f.name for f in file_names]}")
    return file_names