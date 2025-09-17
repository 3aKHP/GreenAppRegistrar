# core.py

import os
import winreg
import winshell
import shutil
import sys
import ctypes
import time
# import subprocess 
from functools import wraps 

CONTEXT_MENU_KEY_PATH = r"exefile\shell\GreenAppRegister"
CONTEXT_MENU_COMMAND_PATH = r"exefile\shell\GreenAppRegister\command"

# --- pywin32 导入 ---
try:
    from win32api import GetFileVersionInfo, LOWORD, HIWORD
    import pywintypes
    # 将 pywintypes.error 赋值给一个变量，方便后续捕获
    PyWinError = pywintypes.error
except ImportError:
    print("警告：未找到 pywin32 库。无法自动提取版本信息。请运行 'pip install pywin32'")
    def GetFileVersionInfo(*args, **kwargs): return None
    # 如果导入失败，将 PyWinError 设为 None，避免后续代码出错
    PyWinError = None
# --- 辅助函数 ---

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
        raise IOError(f"创建卸载脚本失败：\n{e}") from e
        return None

# --- 核心操作函数 (已重构) ---

def register_application(exe_path, app_name, app_version, publisher, force_register=False):
    """
    注册应用的核心逻辑。
    返回: (bool, str) -> (是否成功, 消息)
    """
    if not all([exe_path, app_name, app_version, publisher]):
        return False, "所有字段都不能为空！"

    exe_path = os.path.normpath(exe_path)
    allowed_extensions = ('.exe', '.bat', '.cmd', '.vbs')
    if not os.path.exists(exe_path) or not exe_path.lower().endswith(allowed_extensions):
        return False, f"无效的文件路径或文件类型！\n支持的类型: {', '.join(allowed_extensions)}"

    install_location = os.path.normpath(os.path.dirname(exe_path))
    safe_app_name = "".join(c for c in app_name if c.isalnum() or c in (' ', '_')).rstrip()
    registry_key_name = safe_app_name

    if not force_register and is_application_registered(registry_key_name):
        # 返回一个特殊状态码，让GUI去询问用户
        return "ASK_OVERWRITE", f"应用 '{app_name}' 似乎已经被注册。"

    try:
        # ... (创建快捷方式和注册表的代码，与原文件完全相同) ...
        file_ext = os.path.splitext(exe_path)[1].lower()
        if file_ext == '.exe': icon_path_for_shortcut, icon_path_for_registry = (exe_path, 0), f"{exe_path},0"
        elif file_ext in ('.bat', '.cmd'):
            cmd_icon_path = os.path.normpath(os.path.expandvars(r'%windir%\system32\cmd.exe'))
            icon_path_for_shortcut, icon_path_for_registry = (cmd_icon_path, 0), f"{cmd_icon_path},0"
        elif file_ext == '.vbs':
            wscript_icon_path = os.path.normpath(os.path.expandvars(r'%windir%\system32\wscript.exe'))
            icon_path_for_shortcut, icon_path_for_registry = (wscript_icon_path, 0), f"{wscript_icon_path},0"
        else: icon_path_for_shortcut, icon_path_for_registry = (exe_path, 0), f"{exe_path},0"
        start_menu_path = os.path.join(winshell.programs(), safe_app_name)
        if not os.path.exists(start_menu_path): os.makedirs(start_menu_path)
        shortcut_path = os.path.join(start_menu_path, f"{safe_app_name}.lnk")
        with winshell.shortcut(shortcut_path) as shortcut:
            shortcut.path, shortcut.working_directory, shortcut.description, shortcut.icon_location = exe_path, install_location, f"启动 {app_name}", icon_path_for_shortcut
        try:
            uninstaller_path = create_uninstall_script(install_location, registry_key_name, safe_app_name)
        except IOError as e:
            # 捕获创建脚本时发生的IOError，并将其转换为标准的失败返回
            return False, str(e)
        uninstall_key_path = r"Software\Microsoft\Windows\CurrentVersion\Uninstall"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, uninstall_key_path, 0, winreg.KEY_WRITE) as uninstall_key:
            with winreg.CreateKey(uninstall_key, registry_key_name) as app_key:
                winreg.SetValueEx(app_key, "DisplayName", 0, winreg.REG_SZ, app_name)
                winreg.SetValueEx(app_key, "DisplayVersion", 0, winreg.REG_SZ, app_version)
                winreg.SetValueEx(app_key, "Publisher", 0, winreg.REG_SZ, publisher)
                winreg.SetValueEx(app_key, "InstallLocation", 0, winreg.REG_SZ, install_location)
                winreg.SetValueEx(app_key, "DisplayIcon", 0, winreg.REG_SZ, icon_path_for_registry)
                winreg.SetValueEx(app_key, "UninstallString", 0, winreg.REG_SZ, f'"{uninstaller_path}"')
                estimated_size_kb = get_folder_size_kb(install_location)
                if estimated_size_kb > 0: winreg.SetValueEx(app_key, "EstimatedSize", 0, winreg.REG_DWORD, estimated_size_kb)
                winreg.SetValueEx(app_key, "NoModify", 0, winreg.REG_DWORD, 1)
                winreg.SetValueEx(app_key, "NoRepair", 0, winreg.REG_DWORD, 1)
        return True, f"'{app_name}' 已成功注册/更新！"
    except Exception as e:
        return False, f"注册失败：\n{e}"

def unregister_application(registry_key_name):
    # ... (此函数已符合要求，返回元组，无需修改) ...
    uninstall_key_path = r"Software\Microsoft\Windows\CurrentVersion\Uninstall"
    app_key_path = os.path.join(uninstall_key_path, registry_key_name)
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, app_key_path, 0, winreg.KEY_READ) as app_key:
            display_name, install_location = winreg.QueryValueEx(app_key, "DisplayName")[0], winreg.QueryValueEx(app_key, "InstallLocation")[0]
        uninstaller_path = os.path.join(install_location, "GreenUninstall.bat")
        if os.path.exists(uninstaller_path): os.remove(uninstaller_path)
        safe_app_name = "".join(c for c in display_name if c.isalnum() or c in (' ', '_')).rstrip()
        start_menu_folder = os.path.join(winshell.programs(), safe_app_name)
        if os.path.exists(start_menu_folder): shutil.rmtree(start_menu_folder, ignore_errors=True)
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, uninstall_key_path, 0, winreg.KEY_WRITE) as uninstall_key:
            winreg.DeleteKey(uninstall_key, registry_key_name)
        return True, f"'{display_name}' 已成功反注册。"
    except FileNotFoundError: return False, f"找不到注册信息 '{registry_key_name}'，可能已被移除。"
    except Exception as e: return False, f"反注册时发生错误：\n{e}"

# 在 core.py 中，替换旧的 extract_version_info 函数

def extract_version_info(file_path):
    """
    【最终修复版】从 .exe 文件中提取信息，并静默处理无版本信息的正常情况。
    """
    info = {
        'name': os.path.splitext(os.path.basename(file_path))[0],
        'version': '',
        'publisher': ''
    }

    # 如果 pywin32 未安装或文件类型不对，直接返回
    if not GetFileVersionInfo or not file_path.lower().endswith('.exe'):
        return info

    try:
        # 步骤1: 获取基础信息。如果文件没有版本资源，这一步就会抛出 PyWinError。
        fixed_info = GetFileVersionInfo(file_path, '\\')
        if not fixed_info:
            return info

        # 步骤2: 提取版本号
        try:
            ms = fixed_info['ProductVersionMS']
            ls = fixed_info['ProductVersionLS']
            info['version'] = f"{HIWORD(ms)}.{LOWORD(ms)}.{HIWORD(ls)}.{LOWORD(ls)}"
        except (KeyError, TypeError, ValueError):
            pass # 版本号部分信息缺失，静默跳过

        # 步骤3: 提取语言代码页，这是获取字符串信息的前提
        lang, codepage = GetFileVersionInfo(file_path, '\\VarFileInfo\\Translation')[0]
        
        # 步骤4: 提取发布者
        try:
            str_info_path = f'\\StringFileInfo\\{lang:04x}{codepage:04x}\\CompanyName'
            publisher = GetFileVersionInfo(file_path, str_info_path)
            if publisher:
                info['publisher'] = publisher
        except Exception:
            pass # 发布者信息缺失，静默跳过

        # 步骤5: 提取产品名
        try:
            str_product_path = f'\\StringFileInfo\\{lang:04x}{codepage:04x}\\ProductName'
            product_name = GetFileVersionInfo(file_path, str_product_path)
            if product_name:
                info['name'] = product_name
        except Exception:
            pass # 产品名信息缺失，静默跳过

    except Exception as e:
        # --- 核心修复逻辑 ---
        # 检查捕获到的异常 e 是否是我们预期的 PyWinError
        if PyWinError and isinstance(e, PyWinError):
            # 如果是，说明这只是一个没有版本信息的普通文件。
            # 这是正常情况，我们静默处理，不打印任何东西。
            pass
        else:
            # 如果是其他类型的异常 (如 PermissionError)，说明发生了意外。
            # 这种意外错误我们应该打印出来，以便调试。
            print(f"读取 '{os.path.basename(file_path)}' 的版本信息时发生意外错误: {e}")

    return info


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


# v102 - CLI支持

def find_app_key_by_name(display_name_to_find):
    """
    根据应用的显示名称 (DisplayName) 查找其注册表键名。
    返回: 找到则返回 key_name (str)，否则返回 None。
    """
    # 复用 find_registered_apps_optimized 的扫描逻辑，但目标不同
    uninstall_paths = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
    ]

    for hkey, path in uninstall_paths:
        try:
            with winreg.OpenKey(hkey, path) as key:
                for i in range(winreg.QueryInfoKey(key)[0]):
                    app_key_name = winreg.EnumKey(key, i)
                    try:
                        with winreg.OpenKey(key, app_key_name) as app_key:
                            # 仅当我们关心的两个值都存在时才继续
                            uninstall_string, _ = winreg.QueryValueEx(app_key, "UninstallString")
                            display_name, _ = winreg.QueryValueEx(app_key, "DisplayName")

                            # 确保是我们工具注册的，并且名称匹配
                            if "GreenUninstall.bat" in uninstall_string and display_name.lower() == display_name_to_find.lower():
                                return app_key_name # 找到匹配项，立即返回
                    except (FileNotFoundError, OSError):
                        continue
        except FileNotFoundError:
            continue
    
    return None # 遍历完成仍未找到

# v103 - 右键菜单集成

def is_context_menu_registered():
    """
    检查右键菜单项是否已经被注册。
    返回: bool
    """
    try:
        # 我们只需要检查主键是否存在即可
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, CONTEXT_MENU_KEY_PATH, 0, winreg.KEY_READ):
            return True
    except FileNotFoundError:
        return False
    
def is_admin():
    """检查当前进程是否以管理员权限运行"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    """
    触发UAC，请求以管理员权限重新运行当前脚本。
    无论UAC请求成功与否，当前脚本都将退出。
    """
    if sys.platform == 'win32':
        # --- 诊断代码开始 ---
        print("Attempting to relaunch with admin rights...")
        print(f"Current executable: {sys.executable}")
        
        return_code = 0
        try:
            params = "--uac-relaunch"
            return_code = ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
            print(f"ShellExecuteW returned: {return_code}")

            if return_code > 32:
                print("UAC prompt should be triggered successfully.")
            else:
                # 打印出具体的错误码，可以查表了解失败原因
                print(f"UAC prompt failed silently. Error code: {return_code}")
        except Exception as e:
            print(f"An exception occurred during ShellExecuteW call: {e}", file=sys.stderr)
        
        # --- 诊断代码结束 ---

        # 暂停5秒，让我们能看清控制台的输出
        print("Waiting 5 seconds before exiting...")
        time.sleep(0.1)
        
        os._exit(1) # 无论如何都退出

    return False




def require_admin(func):
    """
    一个装饰器，用于确保函数在管理员权限下执行。
    如果当前无管理员权限，它不会执行函数，而是返回一个特定的元组信号。
    """
    @wraps(func)
    def wrapper(*args, **kwargs):
        if not is_admin():
            # 不再执行任何UI操作或提权操作，仅仅返回一个信号
            return ("NEEDS_ADMIN", "操作需要管理员权限。")
        
        # 如果已经是管理员，则直接执行原函数
        return func(*args, **kwargs)
    return wrapper


@require_admin  # <-- 应用装饰器
def add_to_context_menu():
    """
    将“使用 GreenAppRegistrar 注册”添加到 .exe 文件的右键菜单。
    返回: (bool, str) -> (是否成功, 消息)
    """
    if not sys.executable.lower().endswith('.exe'):
        return False, "错误：此功能仅在程序被编译为 .exe 后才能使用。"

    try:
        # 1. 创建主键...
        with winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, CONTEXT_MENU_KEY_PATH) as key:
            winreg.SetValue(key, None, winreg.REG_SZ, "使用 GreenAppRegistrar 注册")
            winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, f'"{sys.executable}",0')

        # 2. 创建 command 子键...
        with winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, CONTEXT_MENU_COMMAND_PATH) as key:
            command = f'"{sys.executable}" "%1"'
            winreg.SetValue(key, None, winreg.REG_SZ, command)
        
        return True, "成功添加到右键菜单！"

    # 不再需要捕获 PermissionError，但保留捕获其他未知错误是好习惯
    except Exception as e:
        return False, f"发生未知错误：\n{e}"

@require_admin  # <-- 应用装饰器
def remove_from_context_menu():
    """
    从右键菜单中移除我们的项。
    返回: (bool, str) -> (是否成功, 消息)
    """
    try:
        winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, CONTEXT_MENU_COMMAND_PATH)
        winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, CONTEXT_MENU_KEY_PATH)
        return True, "已成功从右键菜单移除。"
    except FileNotFoundError:
        return True, "右键菜单项本就不存在，无需操作。"
    except Exception as e:
        return False, f"发生未知错误：\n{e}"
