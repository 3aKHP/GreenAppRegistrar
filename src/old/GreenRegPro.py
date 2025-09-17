# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys
import winreg
import winshell
# textwrap 不再需要，我们用更可靠的方式处理字符串

# --- 环境设置 ---
if sys.platform == "win32":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception as e:
        print(f"设置DPI感知失败: {e}")

# --- 辅助函数 (无变化) ---

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

# --- GUI 界面 和 主执行块 (无变化) ---
def create_gui():
    root = tk.Tk()
    root.title("绿色软件注册工具")

    frame = tk.Frame(root, padx=10, pady=10)
    frame.pack()

    tk.Label(frame, text="主程序路径:").grid(row=0, column=0, sticky="w", pady=2)
    exe_path_entry = tk.Entry(frame, width=50)
    exe_path_entry.grid(row=0, column=1, pady=2)
    def browse_file():
        filetypes = [
            ("可执行文件", "*.exe;*.bat;*.cmd;*.vbs"),
            ("所有文件", "*.*")
        ]
        path = filedialog.askopenfilename(filetypes=filetypes)
        if path:
            exe_path_entry.delete(0, tk.END)
            exe_path_entry.insert(0, path)
            app_name = os.path.splitext(os.path.basename(path))[0]
            app_name_entry.delete(0, tk.END)
            app_name_entry.insert(0, app_name)

    tk.Button(frame, text="浏览...", command=browse_file).grid(row=0, column=2, padx=5)

    tk.Label(frame, text="应用名称:").grid(row=1, column=0, sticky="w", pady=2)
    app_name_entry = tk.Entry(frame, width=50)
    app_name_entry.grid(row=1, column=1, pady=2)

    tk.Label(frame, text="版本号:").grid(row=2, column=0, sticky="w", pady=2)
    version_entry = tk.Entry(frame, width=50)
    version_entry.grid(row=2, column=1, pady=2)
    version_entry.insert(0, "1.0.0")

    tk.Label(frame, text="发布者:").grid(row=3, column=0, sticky="w", pady=2)
    publisher_entry = tk.Entry(frame, width=50)
    publisher_entry.grid(row=3, column=1, pady=2)
    publisher_entry.insert(0, "PortableApp")

    register_button = tk.Button(
        root, 
        text="注册到系统", 
        command=lambda: register_application(
            exe_path_entry.get(),
            app_name_entry.get(),
            version_entry.get(),
            publisher_entry.get()
        ),
        padx=10, pady=5
    )
    register_button.pack(pady=10)

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
