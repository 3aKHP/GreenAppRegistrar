# main.py

import sys
import os
from tkinter import messagebox

# 仅在需要时导入模块
# import gui
# import core

# --- 环境设置 ---
def initialize_environment():
    """执行程序启动时的初始化设置"""
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception as e:
            print(f"设置DPI感知失败: {e}")

def handle_drag_drop(file_path):
    """处理拖拽注册的逻辑"""
    import core # 延迟导入
    
    allowed_extensions = ('.exe', '.bat', '.cmd', '.vbs')
    if not os.path.exists(file_path) or not file_path.lower().endswith(allowed_extensions):
        messagebox.showerror("拖拽注册错误", f"无效的文件路径或文件类型：\n{file_path}")
        return

    print(f"自动模式启动，目标文件: {file_path}")
    
    # 自动填充信息
    app_info = core.extract_version_info(file_path)
    app_name = app_info.get('name') or os.path.splitext(os.path.basename(file_path))[0]
    app_version = app_info.get('version') or "1.0.0"
    publisher = app_info.get('publisher') or "Auto-Registered via DragDrop"

    # 直接强制注册
    success, message = core.register_application(file_path, app_name, app_version, publisher, force_register=True)
    
    if not success:
        messagebox.showerror("拖拽注册发生错误", f"处理失败：\n{message}")

if __name__ == "__main__":
    initialize_environment()
    
    # 检查是否有命令行参数（拖拽文件就是一种命令行参数）
    if len(sys.argv) > 1:
        try:
            handle_drag_drop(sys.argv[1])
        except Exception as e:
            messagebox.showerror("拖拽注册发生严重错误", f"处理失败：\n{e}")
        sys.exit(0)
    else:
        # 如果没有参数，启动GUI
        import gui # 延迟导入
        gui.create_main_window()
