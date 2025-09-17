# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
import winreg
import winshell
import shutil
import threading
import queue

# textwrap 不再需要，我们用更可靠的方式处理字符串

# --- 环境设置 ---
if sys.platform == "win32":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception as e:
        print(f"设置DPI感知失败: {e}")
# --- 新增：导入 pywin32 用于读取文件版本信息 ---
try:
    from win32api import GetFileVersionInfo, LOWORD, HIWORD
except ImportError:
    # 如果没有安装 pywin32，创建一个伪函数以避免程序在后续调用时崩溃
    print("警告：未找到 pywin32 库。无法自动提取版本信息。请运行 'pip install pywin32'")
    def GetFileVersionInfo(*args, **kwargs): return None
# --- 新增结束 ---
# --- 辅助函数 (无变化) ---

def format_size(size_bytes):
    """将字节大小格式化为KB, MB, GB"""
    if size_bytes is None:
        return "N/A"
    try:
        size_bytes = int(size_bytes)
        if size_bytes < 1024:
            return f"{size_bytes} B"
        kb = size_bytes / 1024
        if kb < 1024:
            return f"{kb:.1f} KB"
        mb = kb / 1024
        if mb < 1024:
            return f"{mb:.1f} MB"
        gb = mb / 1024
        return f"{gb:.2f} GB"
    except (ValueError, TypeError):
        return "N/A"
    
def get_folder_size_kb(folder_path):
    total_size_bytes = 0
    try:
        for dirpath, dirnames, filenames in os.walk(folder_path):
            for f in filenames:
                fp = os.path.join(dirpath, f)
                if os.path.exists(fp):
                    try:
                        total_size_bytes += os.path.getsize(fp)
                    except OSError:
                        pass
    except Exception as e:
        print(f"计算文件夹大小时出错: {e}")
        return 0
    return total_size_bytes // 1024

def is_application_registered(registry_key_name):
    uninstall_key_path = r"Software\Microsoft\Windows\CurrentVersion\Uninstall"
    try:
        key_path = os.path.join(uninstall_key_path, registry_key_name)
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ):
            return True
    except FileNotFoundError:
        return False
    except Exception as e:
        print(f"检查注册表时发生未知错误: {e}")
        return False

# --- 卸载脚本生成 (终极修复版) ---

def create_uninstall_script(install_location, registry_key_name, safe_app_name):
    """
    在应用目录下创建一个 GreenUninstall.bat 脚本，用于执行完整的卸载操作。
    返回创建的脚本的完整路径。
    """
    # 使用一个简单的列表来构建批处理命令，避免所有缩进和格式化问题
    lines = [
        "@echo off",
        "setlocal EnableDelayedExpansion",
        "",
        ":: Define a newline character",
        "set NL=^",
        "",
        "",
        ":: Build the multi-line message string safely",
        f'set "CONFIRM_TITLE=确认卸载"',
        f'set "CONFIRM_MSG=您确定要彻底卸载 \'{safe_app_name}\' 吗？"',
        'set "CONFIRM_MSG=!CONFIRM_MSG!!NL!!NL!此操作将："',
        'set "CONFIRM_MSG=!CONFIRM_MSG!!NL!1. 删除开始菜单快捷方式"',
        'set "CONFIRM_MSG=!CONFIRM_MSG!!NL!2. 删除注册表信息"',
        f'set "CONFIRM_MSG=!CONFIRM_MSG!!NL!3. 删除整个应用程序文件夹："',
        f'set "CONFIRM_MSG=!CONFIRM_MSG!!NL!   {install_location}"',
        'set "CONFIRM_MSG=!CONFIRM_MSG!!NL!!NL!此操作不可恢复！"',
        "",
        'set "TEMP_VBS=%TEMP%\\confirm.vbs"',
        'echo On Error Resume Next > "%TEMP_VBS%"',
        'echo Dim title, msg, result >> "%TEMP_VBS%"',
        'echo title = WScript.Arguments(0) >> "%TEMP_VBS%"',
        'echo msg = WScript.Arguments(1) >> "%TEMP_VBS%"',
        'echo result = MsgBox(msg, vbYesNo + vbExclamation + vbSystemModal, title) >> "%TEMP_VBS%"',
        'echo WScript.Quit(result) >> "%TEMP_VBS%"',
        "",
        'cscript //nologo "%TEMP_VBS%" "!CONFIRM_TITLE!" "!CONFIRM_MSG!"',
        'set "USER_CHOICE=%ERRORLEVEL%"',
        'del "%TEMP_VBS%" >nul 2>nul',
        "",
        'if "%USER_CHOICE%" NEQ "6" (',
        "    echo.",
        "    echo 用户取消了操作。卸载已中止。",
        "    pause",
        "    exit /b",
        ")",
        "",
        "echo.",
        "echo 用户已确认，开始卸载...",
        "",
        f'set "START_MENU_FOLDER=%APPDATA%\\Microsoft\\Windows\\Start Menu\\Programs\\{safe_app_name}"',
        'if exist "%START_MENU_FOLDER%" (',
        "    echo  - 正在删除开始菜单快捷方式...",
        '    rmdir /s /q "%START_MENU_FOLDER%"',
        ")",
        "",
        f'set "REG_KEY=HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{registry_key_name}"',
        "echo  - 正在删除注册表信息...",
        'reg delete "%REG_KEY%" /f >nul 2>nul',
        "",
        "echo.",
        "echo  - 核心文件将在后台静默删除...",
        "",
        f'set "APP_FOLDER={install_location}"',
        'set "DELETER_SCRIPT=%TEMP%\\deleter_%RANDOM%.bat"',
        "",
        "(",
        "    echo @echo off",
        "    echo title 正在完成清理...",
        "    echo timeout /t 2 /nobreak > nul",
        '    echo rmdir /s /q "!APP_FOLDER!"',
        '    echo del "%~f0"',
        ') > "%DELETER_SCRIPT%"',
        "",
        ":: Crucial fix: Start the deleter script IN the temp directory to avoid locking the app folder.",
        'start "Final Cleanup" /d "%TEMP%" cmd /c "%DELETER_SCRIPT%"',
        "",
        "exit /b"
    ]
    batch_script_content = "\r\n".join(lines)

    try:
        uninstaller_path = os.path.join(install_location, "GreenUninstall.bat")
        with open(uninstaller_path, "w", encoding="gbk") as f:
            f.write(batch_script_content)
        return uninstaller_path
    except Exception as e:
        messagebox.showerror("错误", f"创建卸载脚本失败：\n{e}")
        return None

# --- 核心注册逻辑 (无变化) ---

def register_application(exe_path, app_name, app_version, publisher):
    if not all([exe_path, app_name, app_version, publisher]):
        messagebox.showerror("错误", "所有字段都不能为空！")
        return

    exe_path = os.path.normpath(exe_path)

    allowed_extensions = ('.exe', '.bat', '.cmd', '.vbs')
    if not os.path.exists(exe_path) or not exe_path.lower().endswith(allowed_extensions):
        messagebox.showerror("错误", f"无效的文件路径或文件类型！\n支持的类型: {', '.join(allowed_extensions)}")
        return

    install_location = os.path.normpath(os.path.dirname(exe_path))
    
    safe_app_name = "".join(c for c in app_name if c.isalnum() or c in (' ', '_')).rstrip()
    shortcut_name = f"{safe_app_name}.lnk"
    registry_key_name = safe_app_name

    if is_application_registered(registry_key_name):
        response = messagebox.askyesno(
            "警告：应用已注册",
            f"应用 '{app_name}' 似乎已经被注册。\n\n"
            f"您想覆盖/更新现有的注册信息吗？\n（这会重新生成卸载脚本）",
            icon='warning'
        )
        if not response:
            messagebox.showinfo("操作取消", "注册操作已取消。")
            return
        else:
            print("用户选择覆盖现有注册。")

    try:
        file_ext = os.path.splitext(exe_path)[1].lower()
        if file_ext == '.exe':
            icon_path_for_shortcut = (exe_path, 0)
            icon_path_for_registry = f"{exe_path},0"
        elif file_ext in ('.bat', '.cmd'):
            cmd_icon_path = os.path.normpath(os.path.expandvars(r'%windir%\system32\cmd.exe'))
            icon_path_for_shortcut = (cmd_icon_path, 0)
            icon_path_for_registry = f"{cmd_icon_path},0"
        elif file_ext == '.vbs':
            wscript_icon_path = os.path.normpath(os.path.expandvars(r'%windir%\system32\wscript.exe'))
            icon_path_for_shortcut = (wscript_icon_path, 0)
            icon_path_for_registry = f"{wscript_icon_path},0"
        else:
            icon_path_for_shortcut = (exe_path, 0)
            icon_path_for_registry = f"{exe_path},0"

        start_menu_path = os.path.join(winshell.programs(), safe_app_name)
        if not os.path.exists(start_menu_path):
            os.makedirs(start_menu_path)
        
        shortcut_path = os.path.join(start_menu_path, shortcut_name)
        with winshell.shortcut(shortcut_path) as shortcut:
            shortcut.path = exe_path
            shortcut.working_directory = install_location
            shortcut.description = f"启动 {app_name}"
            shortcut.icon_location = icon_path_for_shortcut

        uninstaller_path = create_uninstall_script(install_location, registry_key_name, safe_app_name)
        if not uninstaller_path:
            return

        uninstall_key_path = r"Software\Microsoft\Windows\CurrentVersion\Uninstall"
        
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, uninstall_key_path, 0, winreg.KEY_WRITE) as uninstall_key:
            with winreg.CreateKey(uninstall_key, registry_key_name) as app_key:
                winreg.SetValueEx(app_key, "DisplayName", 0, winreg.REG_SZ, app_name)
                winreg.SetValueEx(app_key, "DisplayVersion", 0, winreg.REG_SZ, app_version)
                winreg.SetValueEx(app_key, "Publisher", 0, winreg.REG_SZ, publisher)
                winreg.SetValueEx(app_key, "InstallLocation", 0, winreg.REG_SZ, install_location)
                winreg.SetValueEx(app_key, "DisplayIcon", 0, winreg.REG_SZ, icon_path_for_registry)
                
                uninstall_command = f'"{uninstaller_path}"'
                winreg.SetValueEx(app_key, "UninstallString", 0, winreg.REG_SZ, uninstall_command)
                
                estimated_size_kb = get_folder_size_kb(install_location)
                if estimated_size_kb > 0:
                    winreg.SetValueEx(app_key, "EstimatedSize", 0, winreg.REG_DWORD, estimated_size_kb)

                winreg.SetValueEx(app_key, "NoModify", 0, winreg.REG_DWORD, 1)
                winreg.SetValueEx(app_key, "NoRepair", 0, winreg.REG_DWORD, 1)

        messagebox.showinfo("成功", f"'{app_name}' 已成功注册/更新！\n卸载脚本 'GreenUninstall.bat' 已创建。")

    except Exception as e:
        messagebox.showerror("发生错误", f"注册失败：\n{e}")

# --- 新增：从可执行文件中提取版本信息的函数 ---
def extract_version_info(file_path):
    """
    从给定的 .exe 或 .dll 文件中提取版本和发布者信息。

    Args:
        file_path (str): 文件的完整路径。

    Returns:
        dict: 包含 'name', 'version', 'publisher' 的字典。如果失败则值为空字符串。
    """
    # 默认使用文件名（不含扩展名）作为应用名
    info = {
        'name': os.path.splitext(os.path.basename(file_path))[0],
        'version': '',
        'publisher': ''
    }

    # 如果 pywin32 未安装，GetFileVersionInfo 会是 None，直接返回默认值
    if not GetFileVersionInfo or not file_path.lower().endswith('.exe'):
        return info

    try:
        # 1. 获取固定的文件版本信息
        fixed_info = GetFileVersionInfo(file_path, '\\')
        if not fixed_info:
            return info # 文件没有版本信息

        # 2. 提取产品版本号 (比文件版本号更适合展示)
        ms = fixed_info['ProductVersionMS']
        ls = fixed_info['ProductVersionLS']
        version = f"{HIWORD(ms)}.{LOWORD(ms)}.{HIWORD(ls)}.{LOWORD(ls)}"
        info['version'] = version

        # 3. 获取语言和代码页，用于查找字符串信息
        lang, codepage = GetFileVersionInfo(file_path, '\\VarFileInfo\\Translation')[0]

        # 4. 构造字符串信息的路径并提取发布者 ('CompanyName')
        str_info_path = f'\\StringFileInfo\\{lang:04x}{codepage:04x}\\CompanyName'
        publisher = GetFileVersionInfo(file_path, str_info_path)
        info['publisher'] = publisher
        
        # 5. 尝试获取产品名作为更佳的应用名
        str_product_path = f'\\StringFileInfo\\{lang:04x}{codepage:04x}\\ProductName'
        product_name = GetFileVersionInfo(file_path, str_product_path)
        if product_name:
            info['name'] = product_name

    except Exception as e:
        # 如果文件没有版本信息或发生其他错误，静默失败，返回已有的信息
        print(f"读取 '{os.path.basename(file_path)}' 的版本信息时出错: {e}")
        pass

    return info
# --- 新增结束 ---

def unregister_application(registry_key_name):
    """
    反注册一个应用：删除其快捷方式、注册表项和卸载脚本，但保留应用文件。
    """
    uninstall_key_path = r"Software\Microsoft\Windows\CurrentVersion\Uninstall"
    app_key_path = os.path.join(uninstall_key_path, registry_key_name)

    try:
        # 1. 从注册表读取必要信息 (在删除它之前！)
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, app_key_path, 0, winreg.KEY_READ) as app_key:
            display_name = winreg.QueryValueEx(app_key, "DisplayName")[0]
            install_location = winreg.QueryValueEx(app_key, "InstallLocation")[0]

        # 2. 删除 GreenUninstall.bat 脚本
        uninstaller_path = os.path.join(install_location, "GreenUninstall.bat")
        if os.path.exists(uninstaller_path):
            os.remove(uninstaller_path)
            print(f"已删除卸载脚本: {uninstaller_path}")

        # 3. 删除开始菜单快捷方式文件夹
        # 注意：这里的 safe_app_name 逻辑必须和注册时完全一致！
        safe_app_name = "".join(c for c in display_name if c.isalnum() or c in (' ', '_')).rstrip()
        start_menu_folder = os.path.join(winshell.programs(), safe_app_name)
        if os.path.exists(start_menu_folder):
            shutil.rmtree(start_menu_folder, ignore_errors=True)
            print(f"已删除开始菜单文件夹: {start_menu_folder}")

        # 4. 删除注册表项 (这是最后一步)
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, uninstall_key_path, 0, winreg.KEY_WRITE) as uninstall_key:
            winreg.DeleteKey(uninstall_key, registry_key_name)
            print(f"已删除注册表项: {registry_key_name}")
        
        return True, f"'{display_name}' 已成功反注册。"

    except FileNotFoundError:
        return False, f"找不到注册信息 '{registry_key_name}'，可能已被移除。"
    except Exception as e:
        return False, f"反注册时发生错误：\n{e}"

def find_registered_apps_optimized():
    """
    高效扫描注册表，仅查找由本工具注册的应用。
    判断依据：UninstallString 的值包含 'GreenUninstall.bat'。
    """
    app_data = []
    # 我们需要检查的注册表位置
    # 1. 64位系统的64位应用 和 32位系统的32位应用
    # 2. 64位系统的32位应用 (Wow6432Node)
    # 3. 当前用户的应用
    uninstall_paths = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
    ]

    for hkey, path in uninstall_paths:
        try:
            # 打开注册表项
            with winreg.OpenKey(hkey, path) as key:
                # 遍历该项下的所有子项（每个子项代表一个应用）
                for i in range(winreg.QueryInfoKey(key)[0]):
                    app_key_name = winreg.EnumKey(key, i)
                    try:
                        with winreg.OpenKey(key, app_key_name) as app_key:
                            # --- 核心优化点 ---
                            # 首先，只读取 UninstallString
                            uninstall_string, _ = winreg.QueryValueEx(app_key, "UninstallString")
                            
                            # 如果不包含我们的目标脚本，立即跳过
                            if "GreenUninstall.bat" not in uninstall_string:
                                continue

                            # --- 如果匹配，才读取其他信息 ---
                            # 读取应用名称 (DisplayName)
                            try:
                                display_name, _ = winreg.QueryValueEx(app_key, "DisplayName")
                            except FileNotFoundError:
                                display_name = app_key_name # 备用名称

                            # 读取发布者 (Publisher)
                            try:
                                publisher, _ = winreg.QueryValueEx(app_key, "Publisher")
                            except FileNotFoundError:
                                publisher = "N/A"

                            # 读取版本号 (DisplayVersion)
                            try:
                                version, _ = winreg.QueryValueEx(app_key, "DisplayVersion")
                            except FileNotFoundError:
                                version = "N/A"
                                
                            # 读取大小 (EstimatedSize)，这是一个DWORD值，速度飞快
                            try:
                                # EstimatedSize 单位是 KB，需要乘以 1024
                                size_kb, _ = winreg.QueryValueEx(app_key, "EstimatedSize")
                                size = format_size(size_kb * 1024)
                            except FileNotFoundError:
                                size = "N/A"

                            app_info = {
                                "key_name": app_key_name,
                                "display_name": display_name,
                                "publisher": publisher,
                                "version": version,
                                "size": size,
                            }
                            app_data.append(app_info)

                    except (FileNotFoundError, OSError):
                        # 某些子项可能没有 UninstallString 或无法访问，直接跳过
                        continue
        except FileNotFoundError:
            # 如果某个根路径不存在（例如32位系统没有Wow6432Node），也直接跳过
            continue
            
    return app_data



def open_unregister_window(root):
    """打开一个新窗口，使用 Treeview 显示并管理已注册的应用。"""
    
    unregister_win = tk.Toplevel(root)
    unregister_win.title("管理已注册的应用")
    unregister_win.geometry("600x400")
    unregister_win.minsize(450, 300)

    frame = tk.Frame(unregister_win, padx=10, pady=10)
    frame.pack(fill=tk.BOTH, expand=True)

    tk.Label(frame, text="以下是已注册的绿色软件列表:").pack(anchor="w")

    # --- Treeview 设置 ---
    tree_frame = tk.Frame(frame)
    
    columns = ('name', 'publisher', 'version', 'size')
    tree = ttk.Treeview(tree_frame, columns=columns, show='headings')

    tree.heading('name', text='应用名称')
    tree.heading('publisher', text='发布者')
    tree.heading('version', text='版本号')
    tree.heading('size', text='大小')

    tree.column('name', width=200, stretch=tk.YES)
    tree.column('publisher', width=120, stretch=tk.YES)
    tree.column('version', width=80, stretch=tk.NO)
    tree.column('size', width=100, stretch=tk.NO, anchor='e')

    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    tree.config(yscrollcommand=scrollbar.set)
    
    tree_frame.pack(fill=tk.BOTH, expand=True, pady=5)

    # --- 按钮框架 ---
    button_frame = tk.Frame(frame)
    button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(5, 0))

    # 创建一个队列用于线程间通信
    app_data_queue = queue.Queue()

    def populate_tree(app_data):
        """用获取到的数据填充Treeview"""
        # 清空旧数据（包括加载提示）
        for item in tree.get_children():
            tree.delete(item)

        if not app_data:
            tree.insert('', tk.END, values=(" (没有找到已注册的应用)", "", "", ""), tags=('empty',))
            tree.tag_configure('empty', foreground='grey')
        else:
            for app in app_data:
                tree.insert(
                    '', tk.END, 
                    values=(app["display_name"], app["publisher"], app["version"], app["size"]),
                    iid=app["key_name"]
                )

    def check_queue():
        """定期检查队列，看工作线程是否已完成任务"""
        try:
            # 尝试从队列中获取数据，不阻塞
            app_data = app_data_queue.get_nowait()
            # 如果获取成功，说明任务完成，用数据填充列表
            populate_tree(app_data)
        except queue.Empty:
            # 如果队列为空，说明任务还未完成，100毫秒后再次检查
            unregister_win.after(100, check_queue)

    def worker_thread_task():
        """在工作线程中执行的耗时任务"""
        # 执行优化后的、极速的耗时操作
        app_data = find_registered_apps_optimized() # <--- 修改在这里
        # 将结果放入队列
        app_data_queue.put(app_data)

    def refresh_list():
        """启动刷新过程"""
        # 1. 清空旧数据
        for item in tree.get_children():
            tree.delete(item)
        
        # 2. 插入“正在加载”的提示
        tree.insert('', tk.END, values=("正在扫描注册表，请稍候...", "", "", ""), tags=('loading',))
        tree.tag_configure('loading', foreground='blue')
        
        # 3. 启动工作线程来执行耗时任务
        threading.Thread(target=worker_thread_task, daemon=True).start()
        
        # 4. 开始定期检查队列以获取结果
        unregister_win.after(100, check_queue)

    def do_unregister():
        # ... (这部分代码无需修改，逻辑是正确的)
        selected_iids = tree.selection()
        if not selected_iids:
            messagebox.showwarning("未选择", "请至少选择一个要反注册的应用。", parent=unregister_win)
            return
        
        if tree.item(selected_iids[0])['values'][0].strip().startswith("("):
             messagebox.showwarning("无效选择", "请不要选择提示信息行。", parent=unregister_win)
             return

        confirm = messagebox.askyesno(
            "确认操作",
            f"您确定要反注册选中的 {len(selected_iids)} 个应用吗？\n\n"
            "此操作将移除它们的快捷方式和注册信息，但不会删除软件文件。",
            icon='warning',
            parent=unregister_win
        )
        if not confirm:
            return

        # 这里也可以用多线程优化，但反注册通常很快，暂时不改
        success_count = 0
        fail_count = 0
        error_messages = []

        for key_name in selected_iids:
            success, message = unregister_application(key_name)
            if success:
                success_count += 1
            else:
                fail_count += 1
                error_messages.append(message)
        
        result_message = f"操作完成。\n\n成功: {success_count}\n失败: {fail_count}"
        if error_messages:
            result_message += "\n\n错误详情:\n" + "\n".join(error_messages)
        
        messagebox.showinfo("反注册结果", result_message, parent=unregister_win)
        refresh_list()

    tk.Button(button_frame, text="反注册选中项", command=do_unregister).pack(side=tk.LEFT, expand=True, padx=5)
    tk.Button(button_frame, text="刷新列表", command=refresh_list).pack(side=tk.LEFT, expand=True, padx=5)
    tk.Button(button_frame, text="关闭", command=unregister_win.destroy).pack(side=tk.RIGHT, expand=True, padx=5)

    # 初始加载
    refresh_list()
    
    unregister_win.transient(root)
    unregister_win.grab_set()
    root.wait_window(unregister_win)

# --- GUI 界面 和 主执行块 (针对性优化) ---
def create_gui():
    root = tk.Tk()
    root.title("绿色软件注册工具")

    frame = tk.Frame(root, padx=10, pady=10)
    frame.pack()

    tk.Label(frame, text="主程序路径:").grid(row=0, column=0, sticky="w", pady=2)
    exe_path_entry = tk.Entry(frame, width=50)
    exe_path_entry.grid(row=0, column=1, pady=2)

    # --- 找到这些输入框变量，我们将在 browse_file 中使用它们 ---
    app_name_entry = tk.Entry(frame, width=50)
    version_entry = tk.Entry(frame, width=50)
    publisher_entry = tk.Entry(frame, width=50)
    # ---

    def browse_file():
        """
        【已优化】打开文件对话框，并自动提取信息填充输入框。
        """
        filetypes = [
            ("可执行文件", "*.exe"), # 优先引导用户选择 .exe
            ("所有支持的文件", "*.exe;*.bat;*.cmd;*.vbs"),
            ("所有文件", "*.*")
        ]
        path = filedialog.askopenfilename(filetypes=filetypes)
        if not path:
            return # 用户取消选择

        # 1. 设置路径输入框
        exe_path_entry.delete(0, tk.END)
        exe_path_entry.insert(0, path)

        # 2. 调用新函数提取信息
        app_info = extract_version_info(path)

        # 3. 清空并填充其他输入框
        app_name_entry.delete(0, tk.END)
        app_name_entry.insert(0, app_info.get('name', ''))

        version_entry.delete(0, tk.END)
        version_entry.insert(0, app_info.get('version', '1.0.0')) # 如果为空，提供默认值

        publisher_entry.delete(0, tk.END)
        publisher_entry.insert(0, app_info.get('publisher', 'PortableApp')) # 如果为空，提供默认值


    tk.Button(frame, text="浏览...", command=browse_file).grid(row=0, column=2, padx=5)

    tk.Label(frame, text="应用名称:").grid(row=1, column=0, sticky="w", pady=2)
    app_name_entry.grid(row=1, column=1, pady=2)

    tk.Label(frame, text="版本号:").grid(row=2, column=0, sticky="w", pady=2)
    version_entry.grid(row=2, column=1, pady=2)
    # 不再需要预先插入默认值，因为 browse_file 会处理
    # version_entry.insert(0, "1.0.0") 

    tk.Label(frame, text="发布者:").grid(row=3, column=0, sticky="w", pady=2)
    publisher_entry.grid(row=3, column=1, pady=2)
    # 不再需要预先插入默认值
    # publisher_entry.insert(0, "PortableApp")

    # --- 底部按钮和主循环部分保持不变 ---
    bottom_frame = tk.Frame(root)
    bottom_frame.pack(pady=10, fill=tk.X, padx=10)

    register_button = tk.Button(
        bottom_frame, 
        text="注册到系统", 
        command=lambda: register_application(
            exe_path_entry.get(),
            app_name_entry.get(),
            version_entry.get(),
            publisher_entry.get()
        ),
        padx=10, pady=5
    )
    register_button.pack(side=tk.LEFT, expand=True, padx=5)

    unregister_button = tk.Button(
        bottom_frame,
        text="管理 / 反注册...",
        command=lambda: open_unregister_window(root),
        padx=10, pady=5
    )
    unregister_button.pack(side=tk.RIGHT, expand=True, padx=5)
    root.mainloop()


if __name__ == "__main__":
    print(f"脚本启动，接收到的参数: {sys.argv}")

    if len(sys.argv) > 1:
        try:
            exe_path = sys.argv[1]
            
            allowed_extensions = ('.exe', '.bat', '.cmd', '.vbs')
            if not os.path.exists(exe_path) or not exe_path.lower().endswith(allowed_extensions):
                messagebox.showerror("拖拽注册错误", f"无效的文件路径或文件类型：\n{exe_path}")
                sys.exit(1)

            app_name = os.path.splitext(os.path.basename(exe_path))[0]
            app_version = "1.0.0"
            publisher = "Auto-Registered via DragDrop"

            print(f"自动模式启动，目标文件: {exe_path}")
            register_application(exe_path, app_name, app_version, publisher)
            sys.exit(0)
        except Exception as e:
            messagebox.showerror("拖拽注册发生严重错误", f"处理失败：\n{e}")
            sys.exit(1)
    else:
        create_gui()
