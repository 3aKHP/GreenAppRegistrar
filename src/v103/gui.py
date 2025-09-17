# gui.py

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import queue
import os

# 从我们的核心逻辑模块导入所有需要的函数
import core

def open_unregister_window(root):
    # ... (此函数代码几乎不变，只需将 unregister_application 调用改为 core.unregister_application) ...
    unregister_win = tk.Toplevel(root)
    unregister_win.title("管理已注册的应用")
    unregister_win.geometry("600x400")
    unregister_win.minsize(450, 300)
    frame = tk.Frame(unregister_win, padx=10, pady=10)
    frame.pack(fill=tk.BOTH, expand=True)
    tk.Label(frame, text="以下是已注册的绿色软件列表:").pack(anchor="w")
    tree_frame = tk.Frame(frame)
    columns = ('name', 'publisher', 'version', 'size')
    tree = ttk.Treeview(tree_frame, columns=columns, show='headings')
    tree.heading('name', text='应用名称'); tree.heading('publisher', text='发布者')
    tree.heading('version', text='版本号'); tree.heading('size', text='大小')
    tree.column('name', width=200, stretch=tk.YES); tree.column('publisher', width=120, stretch=tk.YES)
    tree.column('version', width=80, stretch=tk.NO); tree.column('size', width=100, stretch=tk.NO, anchor='e')
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    tree.config(yscrollcommand=scrollbar.set)
    tree_frame.pack(fill=tk.BOTH, expand=True, pady=5)
    button_frame = tk.Frame(frame)
    button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(5, 0))
    app_data_queue = queue.Queue()
    def populate_tree(app_data):
        for item in tree.get_children(): tree.delete(item)
        if not app_data:
            tree.insert('', tk.END, values=(" (没有找到已注册的应用)", "", "", ""), tags=('empty',))
            tree.tag_configure('empty', foreground='grey')
        else:
            for app in app_data: tree.insert('', tk.END, values=(app["display_name"], app["publisher"], app["version"], app["size"]), iid=app["key_name"])
    def check_queue():
        try:
            app_data = app_data_queue.get_nowait()
            populate_tree(app_data)
        except queue.Empty: unregister_win.after(100, check_queue)
    def worker_thread_task():
        app_data = core.find_registered_apps_optimized() # <-- 调用 core 函数
        app_data_queue.put(app_data)
    def refresh_list():
        for item in tree.get_children(): tree.delete(item)
        tree.insert('', tk.END, values=("正在扫描注册表，请稍候...", "", "", ""), tags=('loading',))
        tree.tag_configure('loading', foreground='blue')
        threading.Thread(target=worker_thread_task, daemon=True).start()
        unregister_win.after(100, check_queue)
    def do_unregister():
        selected_iids = tree.selection()
        if not selected_iids or tree.item(selected_iids[0])['values'][0].strip().startswith("("):
            messagebox.showwarning("无效选择", "请选择一个有效的应用进行反注册。", parent=unregister_win)
            return
        confirm = messagebox.askyesno("确认操作", f"您确定要反注册选中的 {len(selected_iids)} 个应用吗？\n\n此操作将移除它们的快捷方式和注册信息，但不会删除软件文件。", icon='warning', parent=unregister_win)
        if not confirm: return
        success_count, fail_count, error_messages = 0, 0, []
        for key_name in selected_iids:
            success, message = core.unregister_application(key_name) # <-- 调用 core 函数
            if success: success_count += 1
            else: fail_count += 1; error_messages.append(message)
        result_message = f"操作完成。\n\n成功: {success_count}\n失败: {fail_count}"
        if error_messages: result_message += "\n\n错误详情:\n" + "\n".join(error_messages)
        messagebox.showinfo("反注册结果", result_message, parent=unregister_win)
        refresh_list()
    tk.Button(button_frame, text="反注册选中项", command=do_unregister).pack(side=tk.LEFT, expand=True, padx=5)
    tk.Button(button_frame, text="刷新列表", command=refresh_list).pack(side=tk.LEFT, expand=True, padx=5)
    tk.Button(button_frame, text="关闭", command=unregister_win.destroy).pack(side=tk.RIGHT, expand=True, padx=5)
    refresh_list()
    unregister_win.transient(root); unregister_win.grab_set(); root.wait_window(unregister_win)

# gui.py (在 open_unregister_window 和 create_main_window 之间添加)

def open_settings_window(root):
    """打开设置窗口，用于管理右键菜单等功能"""
    settings_win = tk.Toplevel(root)
    settings_win.title("设置")
    settings_win.geometry("400x250")
    settings_win.resizable(False, False)
    
    frame = tk.Frame(settings_win, padx=15, pady=15)
    frame.pack(fill=tk.BOTH, expand=True)

    # --- 描述信息 ---
    desc_label = tk.Label(frame, text="将 '使用 GreenAppRegistrar 注册' 添加到 .exe 文件右键菜单，方便快速注册。", wraplength=360, justify=tk.LEFT)
    desc_label.pack(pady=(0, 15), anchor='w')

    # --- 状态显示 ---
    status_frame = tk.Frame(frame)
    status_frame.pack(fill=tk.X, pady=5)
    tk.Label(status_frame, text="当前状态:").pack(side=tk.LEFT)
    status_var = tk.StringVar(value="正在检查...")
    # 移除有问题的字体，使用系统默认字体
    status_label = tk.Label(status_frame, textvariable=status_var, font=("Segoe UI", 9, "bold"))
    status_label.pack(side=tk.LEFT, padx=5)

    # --- 按钮 ---
    button_frame = tk.Frame(frame)
    button_frame.pack(pady=20, fill=tk.X)
    
    add_button = tk.Button(button_frame, text="添加到右键菜单", width=18)
    add_button.pack(side=tk.LEFT, expand=True, padx=5)
    
    remove_button = tk.Button(button_frame, text="从右键菜单移除", width=18)
    remove_button.pack(side=tk.RIGHT, expand=True, padx=5)

    # --- 逻辑与状态更新 ---
    def update_status():
        """检查注册表并更新UI状态"""
        try:
            is_registered = core.is_context_menu_registered()
        except PermissionError: # <-- 将 Exception 改为更精确的 PermissionError
            # 只有在明确知道是权限问题时，才显示这个状态
            is_registered = False
            status_var.set("未知 (权限不足)")
            status_label.config(fg="orange")
            # 权限不足时，两个按钮都应该是可用的，以便用户可以点击触发提权
            add_button.config(state=tk.NORMAL)
            remove_button.config(state=tk.NORMAL)
            return
        # 对于其他任何异常，我们不再捕获，让程序直接报错，以便调试

        if is_registered:
            status_var.set("已注册")
            status_label.config(fg="green")
            add_button.config(state=tk.DISABLED)
            remove_button.config(state=tk.NORMAL)
        else:
            status_var.set("未注册")
            status_label.config(fg="red")
            add_button.config(state=tk.NORMAL)
            remove_button.config(state=tk.DISABLED)


    def handle_add():
        """处理“添加”按钮点击事件，并响应提权信号"""
        status, message = core.add_to_context_menu()

        # 检查是否收到了提权信号
        if status == "NEEDS_ADMIN":
            messagebox.showinfo(
                "需要管理员权限",
                "此操作需要管理员权限。\n\n程序将尝试重新启动并请求授权（UAC）。请在弹出的窗口中点击“是”。",
                parent=settings_win
            )
            # 由GUI层发起提权
            core.run_as_admin()
            # 提权会关闭当前窗口，所以后续代码不会执行
            return

        # 如果不是提权信号，就按正常流程处理
        if status is True:
            messagebox.showinfo("成功", message, parent=settings_win)
        else: # status is False
            messagebox.showerror("错误", message, parent=settings_win)
        
        update_status()

    def handle_remove():
        """处理“移除”按钮点击事件，并响应提权信号"""
        status, message = core.remove_from_context_menu()

        # 检查是否收到了提权信号
        if status == "NEEDS_ADMIN":
            messagebox.showinfo(
                "需要管理员权限",
                "此操作需要管理员权限。\n\n程序将尝试重新启动并请求授权（UAC）。请在弹出的窗口中点击“是”。",
                parent=settings_win
            )
            # 由GUI层发起提权
            core.run_as_admin()
            # 提权会关闭当前窗口，所以后续代码不会执行
            return

        # 如果不是提权信号，就按正常流程处理
        if status is True:
            messagebox.showinfo("成功", message, parent=settings_win)
        else: # status is False
            messagebox.showerror("错误", message, parent=settings_win)
        
        update_status()

    add_button.config(command=handle_add)
    remove_button.config(command=handle_remove)

    # --- 窗口初始化 ---
    update_status() # 窗口打开时立即检查一次状态
    settings_win.transient(root)
    settings_win.grab_set()
    root.wait_window(settings_win)



def create_main_window():
    """创建并运行主GUI窗口"""
    root = tk.Tk()
    root.title("绿色软件注册工具")
    frame = tk.Frame(root, padx=10, pady=10)
    frame.pack()

    # --- 定义控件 ---
    tk.Label(frame, text="主程序路径:").grid(row=0, column=0, sticky="w", pady=2)
    exe_path_entry = tk.Entry(frame, width=50)
    exe_path_entry.grid(row=0, column=1, pady=2)
    tk.Label(frame, text="应用名称:").grid(row=1, column=0, sticky="w", pady=2)
    app_name_entry = tk.Entry(frame, width=50)
    app_name_entry.grid(row=1, column=1, pady=2)
    tk.Label(frame, text="版本号:").grid(row=2, column=0, sticky="w", pady=2)
    version_entry = tk.Entry(frame, width=50)
    version_entry.grid(row=2, column=1, pady=2)
    tk.Label(frame, text="发布者:").grid(row=3, column=0, sticky="w", pady=2)
    publisher_entry = tk.Entry(frame, width=50)
    publisher_entry.grid(row=3, column=1, pady=2)

    def browse_file():
        # ... (此函数代码不变，只需将 extract_version_info 调用改为 core.extract_version_info) ...
        filetypes = [("可执行文件", "*.exe"), ("所有支持的文件", "*.exe;*.bat;*.cmd;*.vbs"), ("所有文件", "*.*")]
        path = filedialog.askopenfilename(filetypes=filetypes)
        if not path: return
        path = os.path.normpath(path)
        exe_path_entry.delete(0, tk.END); 
        exe_path_entry.insert(0, path)
        app_info = core.extract_version_info(path) # <-- 调用 core 函数
        app_name_entry.delete(0, tk.END); 
        app_name_entry.insert(0, app_info.get('name', ''))
        version_entry.delete(0, tk.END)
        version_entry.insert(0, app_info.get('version') or '1.0.0')
        publisher_entry.delete(0, tk.END)
        publisher_entry.insert(0, app_info.get('publisher') or 'PortableApp')

    def handle_register():
        """【已重构】处理注册按钮点击事件，与核心逻辑交互"""
        status, message = core.register_application(
            exe_path_entry.get(), app_name_entry.get(),
            version_entry.get(), publisher_entry.get()
        )
        if status is True:
            messagebox.showinfo("成功", message)
        elif status == "ASK_OVERWRITE":
            response = messagebox.askyesno("警告：应用已注册", f"{message}\n\n您想覆盖/更新现有的注册信息吗？", icon='warning')
            if response:
                # 用户同意覆盖，再次调用，并设置 force_register=True
                success, msg = core.register_application(
                    exe_path_entry.get(), app_name_entry.get(),
                    version_entry.get(), publisher_entry.get(),
                    force_register=True
                )
                if success: messagebox.showinfo("成功", msg)
                else: messagebox.showerror("错误", msg)
        else: # status is False
            messagebox.showerror("错误", message)

    tk.Button(frame, text="浏览...", command=browse_file).grid(row=0, column=2, padx=5)
    # ...
    bottom_frame = tk.Frame(root)
    bottom_frame.pack(pady=10, fill=tk.X, padx=10)

    # 主操作按钮
    register_button = tk.Button(bottom_frame, text="注册到系统", command=handle_register, padx=10, pady=5)
    register_button.pack(fill=tk.X, ipady=5)

    # 辅助功能按钮
    sub_button_frame = tk.Frame(root)
    sub_button_frame.pack(pady=(0, 10), fill=tk.X, padx=10)

    unregister_button = tk.Button(sub_button_frame, text="管理 / 反注册...", command=lambda: open_unregister_window(root))
    unregister_button.pack(side=tk.LEFT, expand=True, padx=5)

    # 新增的设置按钮
    settings_button = tk.Button(sub_button_frame, text="设置...", command=lambda: open_settings_window(root))
    settings_button.pack(side=tk.RIGHT, expand=True, padx=5)

    root.mainloop()

