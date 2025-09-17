# main.py

import sys
import os
import argparse
from tkinter import messagebox

# 延迟导入，仅在需要时加载
# import core
# import gui

def initialize_environment():
    """执行程序启动时的初始化设置"""
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception as e:
            print(f"警告: 设置DPI感知失败: {e}")

# --- CLI 命令处理函数 ---

def handle_cli_register(args):
    """处理 'register' 命令"""
    import core # 延迟导入
    print(f"正在注册应用: {args.path}")

    # 如果用户没有提供名称、版本等，自动提取
    app_info = core.extract_version_info(args.path)
    app_name = args.name or app_info.get('name') or os.path.splitext(os.path.basename(args.path))[0]
    app_version = args.version or app_info.get('version') or "1.0.0"
    publisher = args.publisher or app_info.get('publisher') or "CLI Registered"

    # --- 修正部分开始 ---
    status, message = core.register_application(
        args.path, app_name, app_version, publisher, force_register=args.force
    )

    if status is True:
        print(f"成功: {message}")
    elif status == "ASK_OVERWRITE":
        print(f"警告: {message}")
        print("操作被中止。如需覆盖，请使用 --force 标志。")
    else: # status is False
        print(f"错误: {message}", file=sys.stderr)
        sys.exit(1) # 以错误码退出

def handle_cli_unregister(args):
    """处理 'unregister' 命令"""
    import core # 延迟导入
    print(f"正在查找要反注册的应用: '{args.name}'...")
    
    key_name = core.find_app_key_by_name(args.name)
    
    if not key_name:
        print(f"错误: 找不到名为 '{args.name}' 的已注册应用。", file=sys.stderr)
        sys.exit(1)

    print(f"找到应用，注册表键为: {key_name}。正在执行反注册...")
    success, message = core.unregister_application(key_name)
    
    if success:
        print(f"成功: {message}")
    else:
        print(f"错误: {message}", file=sys.stderr)
        sys.exit(1)

def handle_cli_list(args):
    """处理 'list' 命令"""
    import core # 延迟导入
    apps = core.find_registered_apps_optimized()
    if not apps:
        print("没有找到由本工具注册的绿色软件。")
        return

    # 格式化输出为表格
    print(f"{'应用名称':<40} {'发布者':<25} {'版本':<15} {'大小'}")
    print("-" * 95)
    for app in sorted(apps, key=lambda x: x['display_name'].lower()):
        name = (app['display_name'][:37] + '...') if len(app['display_name']) > 40 else app['display_name']
        publisher = (app['publisher'][:22] + '...') if len(app['publisher']) > 25 else app['publisher']
        version = (app['version'][:12] + '...') if len(app['version']) > 15 else app['version']
        print(f"{name:<40} {publisher:<25} {version:<15} {app['size']}")

def handle_drag_drop(file_path):
    """处理拖拽注册的逻辑 (保留作为快捷方式)"""
    import core # 延迟导入
    
    allowed_extensions = ('.exe', '.bat', '.cmd', '.vbs')
    if not os.path.exists(file_path) or not file_path.lower().endswith(allowed_extensions):
        messagebox.showerror("拖拽注册错误", f"无效的文件路径或文件类型：\n{file_path}")
        return

    app_info = core.extract_version_info(file_path)
    app_name = app_info.get('name') or os.path.splitext(os.path.basename(file_path))[0]
    app_version = app_info.get('version') or "1.0.0"
    publisher = app_info.get('publisher') or "Auto-Registered via DragDrop"

    # 拖拽总是强制覆盖，以提供流畅体验
    success, message = core.register_application(file_path, app_name, app_version, publisher, force_register=True)
    
    if success:
        messagebox.showinfo("拖拽注册成功", f"应用 '{app_name}' 已成功注册！")
    else:
        messagebox.showerror("拖拽注册发生错误", f"处理失败：\n{message}")

def main():
    """程序主入口"""
    initialize_environment()

    # --- 命令行解析器设置 ---
    parser = argparse.ArgumentParser(
        prog="GreenAppRegistrar",
        description="一个用于注册、管理绿色软件的工具。",
        epilog="如果不带任何参数运行，将启动图形用户界面 (GUI)。"
    )
    subparsers = parser.add_subparsers(dest="command", help="可用的命令")

    # 'register' 命令
    parser_register = subparsers.add_parser("register", help="注册一个新的绿色软件。")
    parser_register.add_argument("--path", required=True, help="主程序文件的完整路径 (e.g., C:\\Apps\\MyApp\\app.exe)")
    parser_register.add_argument("--name", help="应用名称 (如果省略，将尝试从文件中提取)")
    parser_register.add_argument("--ver", "--version", dest="version", help="版本号 (e.g., 1.2.3)")
    parser_register.add_argument("--pub", "--publisher", dest="publisher", help="发布者名称")
    parser_register.add_argument("--force", action="store_true", help="如果应用已注册，则强制覆盖。")

    # 'unregister' 命令
    parser_unregister = subparsers.add_parser("unregister", help="反注册一个软件。")
    parser_unregister.add_argument("--name", required=True, help="要反注册的应用的“显示名称”。")

    # 'list' 命令
    parser_list = subparsers.add_parser("list", help="列出所有已注册的绿色软件。")

    # --- 逻辑分发 ---
    # 如果参数数量大于1，说明有命令行操作或拖拽
    if len(sys.argv) > 1:
        # 检查是否是已定义的命令
        if sys.argv[1] in ['register', 'unregister', 'list']:
            args = parser.parse_args()
            if args.command == "register":
                handle_cli_register(args)
            elif args.command == "unregister":
                handle_cli_unregister(args)
            elif args.command == "list":
                handle_cli_list(args)
        # 如果不是命令，但只有一个参数，我们假定它是拖拽文件
        elif len(sys.argv) == 2 and sys.argv[1] not in ['-h', '--help']:
             try:
                handle_drag_drop(sys.argv[1])
             except Exception as e:
                messagebox.showerror("拖拽注册发生严重错误", f"处理失败：\n{e}")
        # 其他情况（如 -h, --help），让 argparse 自己处理
        else:
            parser.parse_args()
    else:
        # 没有参数，启动GUI
        print("未提供命令行参数，启动 GUI 模式...")
        import gui # 延迟导入
        gui.create_main_window()

if __name__ == "__main__":
    main()
