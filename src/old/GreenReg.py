import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys
import winreg
import winshell

# --- 核心注册逻辑 ---
def register_application(exe_path, app_name, app_version, publisher):
    if not all([exe_path, app_name, app_version, publisher]):
        messagebox.showerror("错误", "所有字段都不能为空！")
        return

    if not os.path.exists(exe_path) or not exe_path.lower().endswith('.exe'):
        messagebox.showerror("错误", "无效的 .exe 文件路径！")
        return

    install_location = os.path.dirname(exe_path)
    # 使用应用名称作为注册表键名和快捷方式名，移除特殊字符
    safe_app_name = "".join(c for c in app_name if c.isalnum() or c in (' ', '_')).rstrip()
    shortcut_name = f"{safe_app_name}.lnk"
    registry_key_name = safe_app_name

    try:
        # --- 1. 创建开始菜单快捷方式 (为当前用户) ---
        start_menu_path = os.path.join(winshell.programs(), safe_app_name)
        if not os.path.exists(start_menu_path):
            os.makedirs(start_menu_path)
        
        shortcut_path = os.path.join(start_menu_path, shortcut_name)
        with winshell.shortcut(shortcut_path) as shortcut:
            shortcut.path = exe_path
            shortcut.working_directory = install_location
            shortcut.description = f"启动 {app_name}"
            shortcut.icon_location = (exe_path, 0)

        # --- 2. 写入注册表 (为当前用户) ---
        uninstall_key_path = r"Software\Microsoft\Windows\CurrentVersion\Uninstall"
        
        # 打开Uninstall键
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, uninstall_key_path, 0, winreg.KEY_WRITE) as uninstall_key:
            # 创建新的程序键
            with winreg.CreateKey(uninstall_key, registry_key_name) as app_key:
                # 写入各种值
                winreg.SetValueEx(app_key, "DisplayName", 0, winreg.REG_SZ, app_name)
                winreg.SetValueEx(app_key, "DisplayVersion", 0, winreg.REG_SZ, app_version)
                winreg.SetValueEx(app_key, "Publisher", 0, winreg.REG_SZ, publisher)
                winreg.SetValueEx(app_key, "InstallLocation", 0, winreg.REG_SZ, install_location)
                winreg.SetValueEx(app_key, "DisplayIcon", 0, winreg.REG_SZ, f"{exe_path},0")
                
                # 创建卸载命令
                uninstall_command = f'cmd.exe /c "rmdir /s /q \\"{install_location}\\""'
                winreg.SetValueEx(app_key, "UninstallString", 0, winreg.REG_SZ, uninstall_command)
                
                # 估算大小 (这里简单设置为0，可以后续改进)
                # folder_size_kb = sum(os.path.getsize(os.path.join(dirpath, filename)) for dirpath, _, filenames in os.walk(install_location) for filename in filenames) // 1024
                # winreg.SetValueEx(app_key, "EstimatedSize", 0, winreg.REG_DWORD, folder_size_kb)

                # 禁用修改和修复按钮
                winreg.SetValueEx(app_key, "NoModify", 0, winreg.REG_DWORD, 1)
                winreg.SetValueEx(app_key, "NoRepair", 0, winreg.REG_DWORD, 1)

        messagebox.showinfo("成功", f"'{app_name}' 已成功注册！")

    except Exception as e:
        messagebox.showerror("发生错误", f"注册失败：\n{e}")


# --- GUI 界面 ---
def create_gui():
    root = tk.Tk()
    root.title("绿色软件注册工具")

    # --- 布局 ---
    frame = tk.Frame(root, padx=10, pady=10)
    frame.pack()

    # 文件路径
    tk.Label(frame, text="主程序路径 (.exe):").grid(row=0, column=0, sticky="w", pady=2)
    exe_path_entry = tk.Entry(frame, width=50)
    exe_path_entry.grid(row=0, column=1, pady=2)
    def browse_file():
        path = filedialog.askopenfilename(filetypes=[("Executable files", "*.exe")])
        if path:
            exe_path_entry.delete(0, tk.END)
            exe_path_entry.insert(0, path)
            # 尝试自动填充应用名称
            app_name = os.path.splitext(os.path.basename(path))[0]
            app_name_entry.delete(0, tk.END)
            app_name_entry.insert(0, app_name)

    tk.Button(frame, text="浏览...", command=browse_file).grid(row=0, column=2, padx=5)

    # 应用信息
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

    # 注册按钮
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
    create_gui()
