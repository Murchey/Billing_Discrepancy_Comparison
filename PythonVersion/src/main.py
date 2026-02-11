# _*_coding : UTF-8 _*_
# @Time : 2026/2/6 21:00
# @Author : Murchey
# @File : main
# @Project : python
from random import choice

import multipleFiles as mf
import getData as gd
import compare as cp
import getStandardData as gsd
mf.newFolders()
print("已在软件目录下创建工作目录，请按照说明存放文件。\n-------------------------")
print('''.\已提取数据      存放软件工作时产生的文件
.\标准表格        存放标准收费表格，是判断差异时的标准表格
.\需核对表格      存放实际收费单据
.\比对结果        存放差异比对结果
-------------------------''')
while True:
    print('''
选择功能以执行：
1.数据提取
2.标准表格处理
3.数据比对
------------------------------
    ''')
    choice = input()
    if choice == '1':
        print('进入数据提取环节……')
        handle_file_paths = mf.getFilesNames("需核对表格")
        print("准备处理的文件：", handle_file_paths)
        print('工作目录内容如下：')
        # 获取列参数
        print("请输入姓名所在列（如A、B、C等），默认值为C:")
        name_col = input().strip() or "C"
        print("请输入金额所在列（如A、B、C等），默认值为J:")
        money_col = input().strip() or "J"
        
        for item in handle_file_paths:
            print(item)
            gd.mainFunc(item, "已提取数据", name_col, money_col)
    elif choice == '2':
        print('进入标准表格处理环节……')
        # 获取标准表格文件
        std_files = mf.getFilesNames("标准表格")
        if not std_files:
            print("错误：标准表格目录下没有文件！")
        else:
            print("找到的标准表格文件：")
            for i, file in enumerate(std_files, 1):
                print(f"{i}. {file.name}")
            
            # 选择处理模式
            print("请选择处理模式：")
            print("1. 处理所有文件")
            print("2. 处理单个文件")
            mode_choice = input().strip()
            
            # 获取参数（对两种模式都适用）
            print("请输入表头所在行（数字），默认值为3:")
            begin_row_str = input().strip() or "3"
            print("请输入姓名所在列（如A、B、C等），默认值为B:")
            name_col = input().strip().upper() or "B"
            print("请输入金额所在列（如A、B、C等），默认值为F:")
            money_col = input().strip().upper() or "F"
            
            # 验证参数
            try:
                begin_row = int(begin_row_str)
                if begin_row < 1:
                    print("错误：表头所在行必须大于0")
                    continue
            except ValueError:
                print("错误：表头所在行必须是数字")
                continue
            
            if not name_col or not money_col:
                print("错误：请填写所有参数")
                continue
            
            # 构建保存路径
            import sys
            from pathlib import Path
            if getattr(sys, 'frozen', False):
                # 打包后的exe模式
                app_dir = Path(sys.executable).resolve().parent
            else:
                # 脚本模式
                app_dir = Path(__file__).resolve().parent
            
            save_dir = app_dir / "已提取数据" / "标准表格数据"
            save_dir.mkdir(parents=True, exist_ok=True)
            
            if mode_choice == "1":
                # 批量处理所有文件
                print("开始批量处理所有标准表格文件...")
                processed_count = 0
                failed_count = 0
                
                for file_path in std_files:
                    try:
                        print(f"处理文件: {file_path.name}")
                        
                        # 提取数据
                        data = gsd.getValuableData(file_path, begin_row, name_col, money_col)
                        print(f"成功提取 {len(data)} 条数据")
                        
                        # 获取文件名
                        file_name = file_path.stem
                        save_path = save_dir / f"{file_name}_处理结果.xlsx"
                        
                        # 保存数据
                        success = gsd.saveFile(data, save_path)
                        
                        if success:
                            print(f"数据保存成功！")
                            print(f"保存路径: {save_path}")
                            processed_count += 1
                        else:
                            print("数据保存失败！")
                            failed_count += 1
                    except Exception as e:
                        print(f"处理过程中出错: {e}")
                        failed_count += 1
                    print()
                
                # 显示处理结果
                print(f"批量处理完成！")
                print(f"成功处理: {processed_count} 个文件")
                print(f"失败: {failed_count} 个文件")
            else:
                # 处理单个文件
                print("请选择要处理的文件序号：")
                try:
                    file_index = int(input().strip()) - 1
                    if 0 <= file_index < len(std_files):
                        selected_file = std_files[file_index]
                        print(f"已选择文件：{selected_file.name}")
                        
                        # 提取数据
                        print("正在提取数据...")
                        data = gsd.getValuableData(selected_file, begin_row, name_col, money_col)
                        print(f"成功提取 {len(data)} 条数据")
                        
                        # 获取文件名
                        file_name = selected_file.stem
                        save_path = save_dir / f"{file_name}_处理结果.xlsx"
                        
                        # 保存数据
                        print("正在保存数据...")
                        success = gsd.saveFile(data, save_path)
                        
                        if success:
                            print(f"数据保存成功！")
                            print(f"保存路径: {save_path}")
                        else:
                            print("数据保存失败！")
                    else:
                        print("错误：无效的文件序号")
                except ValueError:
                    print("错误：请输入有效的数字")
    elif choice == '3':
        print('进入数据比对环节……')
        # 获取标准表格（从已提取数据/标准表格数据文件夹）和已提取数据表格
        import sys
        from pathlib import Path
        if getattr(sys, 'frozen', False):
            # 打包后的exe模式
            app_dir = Path(sys.executable).resolve().parent
        else:
            # 脚本模式
            app_dir = Path(__file__).resolve().parent
        
        # 获取标准表格数据
        std_data_dir = app_dir / "已提取数据" / "标准表格数据"
        std_files = [p for p in std_data_dir.glob('*.xlsx')]
        
        # 获取已提取数据
        extracted_files = mf.getFilesNames("已提取数据")
        
        if not std_files:
            print("错误：已提取数据/标准表格数据目录下没有文件！")
            print("请先运行标准表格处理功能")
        elif not extracted_files:
            print("错误：已提取数据目录下没有文件！")
        else:
            # 创建标准表格字典，以文件名（不含扩展名）为键
            std_file_dict = {}
            for file in std_files:
                # 获取文件名（不含扩展名），去掉末尾的 "_处理结果"
                file_name = file.stem
                if file_name.endswith("_处理结果"):
                    file_name = file_name[:-5]  # 去掉 "_处理结果"
                std_file_dict[file_name] = file
            
            print(f"找到 {len(std_files)} 个标准表格数据文件")
            print(f"找到 {len(extracted_files)} 个待比对文件")
            print()
            
            # 处理每个待比对文件
            matched_count = 0
            unmatched_count = 0
            
            for file in extracted_files:
                # 获取文件名（不含扩展名），去掉末尾的 "_new"
                file_name = file.stem
                if file_name.endswith("_new"):
                    file_name = file_name[:-4]  # 去掉 "_new"
                
                print(f"处理文件: {file.name}")
                
                # 查找对应的标准表格
                if file_name in std_file_dict:
                    std_file = std_file_dict[file_name]
                    print(f"匹配标准表格: {std_file.name}")
                    
                    try:
                        # 调用比对函数
                        cp.compare_and_save(std_file, file)
                        print(f"文件 {file.name} 比对完成")
                        print(f"结果已保存: 比对结果/{file.stem}_比对结果.txt")
                        matched_count += 1
                    except Exception as e:
                        print(f"比对过程中出错: {e}")
                else:
                    print(f"未找到对应标准表格，跳过比对")
                    unmatched_count += 1
                
                print()
            
            print(f"比对完成！")
            print(f"成功比对: {matched_count} 个文件")
            print(f"未匹配: {unmatched_count} 个文件")
            print("比对结果已保存到 '比对结果' 文件夹")
    else:
        print("不是有效输入，请重新选择。")