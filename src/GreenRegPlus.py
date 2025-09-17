import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys
import winreg
import winshell

# --- 辅助函数 ---
if sys.platform == "win32":
    try:
        import ctypes
        # 设置DPI感知模式。
        # 0: Unaware, 1: System Aware, 2: Per-Monitor Aware
        # System Aware对于大多数简单应用来说足够且稳定
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception as e:
        print(f"设置DPI感知失败: {e}")

def get_folder_size_kb(folder_path):
    """计算指定文件夹的总大小（单位：KB）。"""
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
    """检查指定的应用是否已经在注册表中注册（仅限当前用户）。"""
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

# --- 核心注册逻辑 (已优化) ---

def register_application(exe_path, app_name, app_version, publisher):
    # --- 输入验证 (已优化) ---
    if not all([exe_path, app_name, app_version, publisher]):
        messagebox.showerror("错误", "所有字段都不能为空！")
        return

    # 修改: 允许更多可执行文件类型
    allowed_extensions = ('.exe', '.bat', '.cmd', '.vbs')
    if not os.path.exists(exe_path) or not exe_path.lower().endswith(allowed_extensions):
        messagebox.showerror("错误", f"无效的文件路径或文件类型！\n支持的类型: {', '.join(allowed_extensions)}")
        return

    install_location = os.path.dirname(exe_path)
    safe_app_name = "".join(c for c in app_name if c.isalnum() or c in (' ', '_')).rstrip()
    shortcut_name = f"{safe_app_name}.lnk"
    registry_key_name = safe_app_name

    if is_application_registered(registry_key_name):
        response = messagebox.askyesno(
            "警告：应用已注册",
            f"应用 '{app_name}' 似乎已经被注册。\n\n"
            f"您想覆盖/更新现有的注册信息吗？",
            icon='warning'
        )
        if not response:
            messagebox.showinfo("操作取消", "注册操作已取消。")
            return
        else:
            print("用户选择覆盖现有注册。")

    try:
        # --- 动态设置图标 (新增逻辑) ---
        file_ext = os.path.splitext(exe_path)[1].lower()
        if file_ext == '.exe':
            icon_path_for_shortcut = (exe_path, 0)
            icon_path_for_registry = f"{exe_path},0"
        elif file_ext in ('.bat', '.cmd'):
            # 使用 cmd.exe 的图标
            cmd_icon_path = os.path.expandvars(r'%windir%\system32\cmd.exe')
            icon_path_for_shortcut = (cmd_icon_path, 0)
            icon_path_for_registry = f"{cmd_icon_path},0"
        elif file_ext == '.vbs':
            # 使用 wscript.exe 的图标
            wscript_icon_path = os.path.expandvars(r'%windir%\system32\wscript.exe')
            icon_path_for_shortcut = (wscript_icon_path, 0)
            icon_path_for_registry = f"{wscript_icon_path},0"
        else:
            # 备用，理论上不会执行到这里
            icon_path_for_shortcut = (exe_path, 0)
            icon_path_for_registry = f"{exe_path},0"

        # --- 创建开始菜单快捷方式 ---
        start_menu_path = os.path.join(winshell.programs(), safe_app_name)
        if not os.path.exists(start_menu_path):
            os.makedirs(start_menu_path)
        
        shortcut_path = os.path.join(start_menu_path, shortcut_name)
        with winshell.shortcut(shortcut_path) as shortcut:
            shortcut.path = exe_path
            shortcut.working_directory = install_location
            shortcut.description = f"启动 {app_name}"
            shortcut.icon_location = icon_path_for_shortcut # 修改: 使用动态图标路径

        # --- 写入注册表 ---
        uninstall_key_path = r"Software\Microsoft\Windows\CurrentVersion\Uninstall"
        
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, uninstall_key_path, 0, winreg.KEY_WRITE) as uninstall_key:
            with winreg.CreateKey(uninstall_key, registry_key_name) as app_key:
                winreg.SetValueEx(app_key, "DisplayName", 0, winreg.REG_SZ, app_name)
                winreg.SetValueEx(app_key, "DisplayVersion", 0, winreg.REG_SZ, app_version)
                winreg.SetValueEx(app_key, "Publisher", 0, winreg.REG_SZ, publisher)
                winreg.SetValueEx(app_key, "InstallLocation", 0, winreg.REG_SZ, install_location)
                winreg.SetValueEx(app_key, "DisplayIcon", 0, winreg.REG_SZ, icon_path_for_registry) # 修改: 使用动态图标路径
                
                # 注意: 这个卸载命令会删除整个文件夹。
                uninstall_command = f'cmd.exe /c "rmdir /s /q \\"{install_location}\\""'
                winreg.SetValueEx(app_key, "UninstallString", 0, winreg.REG_SZ, uninstall_command)
                
                estimated_size_kb = get_folder_size_kb(install_location)
                if estimated_size_kb > 0:
                    winreg.SetValueEx(app_key, "EstimatedSize", 0, winreg.REG_DWORD, estimated_size_kb)

                winreg.SetValueEx(app_key, "NoModify", 0, winreg.REG_DWORD, 1)
                winreg.SetValueEx(app_key, "NoRepair", 0, winreg.REG_DWORD, 1)

        messagebox.showinfo("成功", f"'{app_name}' 已成功注册/更新！")

    except Exception as e:
        messagebox.showerror("发生错误", f"注册失败：\n{e}")


# --- GUI 界面 (已优化) ---
def create_gui():
    root = tk.Tk()
    root.title("绿色软件注册工具")

    frame = tk.Frame(root, padx=10, pady=10)
    frame.pack()

    tk.Label(frame, text="主程序路径:").grid(row=0, column=0, sticky="w", pady=2)
    exe_path_entry = tk.Entry(frame, width=50)
    exe_path_entry.grid(row=0, column=1, pady=2)
    def browse_file():
        # 修改: 允许选择多种文件类型
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

# --- 主执行块 (已优化) ---
if __name__ == "__main__":
    # 优化: 增加对 sys.argv 的调试信息，可以在打包为 console 应用时查看
    print(f"脚本启动，接收到的参数: {sys.argv}")

    if len(sys.argv) > 1:
        # --- 自动（拖拽）模式 ---
        try:
            exe_path = sys.argv[1]
            
            # 优化: 使用与核心函数一致的验证逻辑
            allowed_extensions = ('.exe', '.bat', '.cmd', '.vbs')
            if not os.path.exists(exe_path) or not exe_path.lower().endswith(allowed_extensions):
                messagebox.showerror("拖拽注册错误", f"无效的文件路径或文件类型：\n{exe_path}")
                sys.exit(1)

            app_name = os.path.splitext(os.path.basename(exe_path))[0]
            app_version = "1.0.0"
            publisher = "Auto-Registered via DragDrop"

            print(f"自动模式启动，目标文件: {exe_path}")
            # 调用核心函数，它已经包含了所有反馈和错误处理
            register_application(exe_path, app_name, app_version, publisher)
            sys.exit(0)
        except Exception as e:
            # 优化: 捕获所有异常并弹窗提示
            messagebox.showerror("拖拽注册发生严重错误", f"处理失败：\n{e}")
            sys.exit(1)
    else:
        # --- 交互（GUI）模式 ---
        create_gui()
