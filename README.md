# GreenAppRegistrar
A lightweight Windows utility to seamlessly integrate portable ('green') applications into the system, making them appear in 'Apps &amp; features' and the Start Menu, complete with a robust uninstaller.

# GreenAppRegistrar - 绿色软件注册工具

[![Release](https://img.shields.io/github/v/release/3aKHP/GreenAppRegistrar?style=flat-square)](https://github.com/3aKHP/GreenAppRegistrar/releases)
[![License](https://img.shields.io/github/license/3aKHP/GreenAppRegistrar?style=flat-square)](LICENSE)

一个轻量、强大、易用的 Windows 小工具，旨在将“绿色软件”（便携软件）无缝集成到系统中，让你像管理普通安装程序一样管理它们。

---

## 🤔 解决了什么痛点？

很多优秀的绿色软件解压即用，非常方便。但它们的缺点也很明显：
*   不会出现在系统的“应用和功能”列表中，难以统一管理。
*   不会自动创建开始菜单或桌面快捷方式。
*   卸载时需要手动删除文件夹，容易遗忘或删不干净。

**GreenAppRegistrar** 就是为了解决这些问题而生。

## ✨ 功能特性

*   **系统级集成**: 将任何绿色软件（`.exe`, `.bat`, `.cmd`, `.vbs`）注册到 Windows 的“应用和功能”列表中。
*   **自动创建快捷方式**: 在开始菜单中为你的应用创建整洁的快捷方式，并自动关联正确的图标。
*   **一键拖拽注册**: 支持将目标程序的可执行文件直接拖拽到 `GreenAppRegistrar.exe` 上，实现闪电注册。
*   **强大的卸载器**: 为每个注册的应用生成一个 `GreenUninstall.bat` 脚本。这个脚本不仅会清理注册表和快捷方式，还会弹窗确认后**彻底删除整个应用文件夹**，实现完美卸载。
*   **高效管理界面**: 提供一个独立的管理窗口，可以快速浏览所有已注册的绿色软件，并支持批量反注册。
*   **智能扫描**: 使用多线程技术异步扫描注册表，即使注册项繁多，界面也不会卡顿。
*   **纯净无依赖**: 基于 Python 和 Tkinter 构建，无需安装任何额外依赖，下载即用。
*   **高分屏适配**: 自动启用 DPI 感知，在高清屏幕上也能获得清晰的 UI 显示。

## 💻 如何使用

### 方法一：使用图形界面 (推荐)

1.  从 [Releases 页面](https://github.com/你的用户名/GreenAppRegistrar/releases) 下载最新的 `GreenAppRegistrar.exe`。
2.  运行 `GreenAppRegistrar.exe`。
3.  点击“浏览...”选择你绿色软件的主程序文件（如 `PotPlayer.exe`）。
4.  工具会自动填充应用名称，你也可以自定义“应用名称”、“版本号”和“发布者”。
5.  点击“注册到系统”按钮即可。

### 方法二：拖拽注册

1.  从 [Releases 页面](https://github.com/你的用户名/GreenAppRegistrar/releases) 下载最新的 `GreenAppRegistrar.exe`。
2.  将你绿色软件的主程序文件（如 `Notepad++.exe`）直接**拖拽到 `GreenAppRegistrar.exe` 的图标上**。
3.  松开鼠标，程序会自动以后台模式完成注册，无需任何额外操作！


## 🛠️ 从源码构建

如果你想自行修改或编译，请按以下步骤操作：

1.  确保你已安装 Python 和 Git。
2.  安装项目依赖 (主要是 `winshell`):
    ```bash
    pip install winshell
    ```
3.  克隆本仓库:
    ```bash
    git clone https://github.com/3aKHP/GreenAppRegistrar.git
    cd GreenAppRegistrar
    ```
4.  使用 PyInstaller 或 Nuitka 进行编译。例如，使用 PyInstaller 打包成单个 exe 文件：
    ```bash
    pyinstaller --onefile --windowed --name GreenAppRegistrar --icon=your_icon.ico GreenAppRegistrar.py
    ```

## 📄 许可证

本项目基于 [MIT License](LICENSE) 开源。
