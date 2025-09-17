# GreenAppRegistrar
A lightweight Windows utility to seamlessly integrate portable ('green') applications into the system, making them appear in 'Apps & features' and the Start Menu, complete with a robust uninstaller.

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
*   **强大的命令行接口 (CLI)**: 支持 `register`, `unregister`, `list` 等命令，方便高级用户进行脚本化和自动化操作。
*   **自动创建快捷方式**: 在开始菜单中为你的应用创建整洁的快捷方式，并自动关联正确的图标。
*   **一键拖拽注册**: 支持将目标程序的可执行文件直接拖拽到 `GreenAppRegistrar.exe` 上，实现闪电注册。
*   **强大的卸载器**: 为每个注册的应用生成一个 `GreenUninstall.bat` 脚本。这个脚本不仅会清理注册表和快捷方式，还会弹窗确认后**彻底删除整个应用文件夹**，实现完美卸载。
*   **高效管理界面**: 提供一个独立的管理窗口，可以快速浏览所有已注册的绿色软件，并支持批量反注册。
*   **智能扫描**: 使用多线程技术异步扫描注册表，即使注册项繁多，界面也不会卡顿。
*   **纯净无依赖**: 基于 Python 和 Tkinter 构建，无需安装任何额外依赖，下载即用。
*   **高分屏适配**: 自动启用 DPI 感知，在高清屏幕上也能获得清晰的 UI 显示。

## 💻 如何使用

### 方法一：使用图形界面 (推荐)

1.  从 [Releases 页面](https://github.com/3aKHP/GreenAppRegistrar/releases) 下载最新的 `GreenAppRegistrar.exe`。
2.  运行 `GreenAppRegistrar.exe`。
3.  点击“浏览...”选择你绿色软件的主程序文件（如 `PotPlayer.exe`）。
4.  工具会自动填充应用名称，你也可以自定义“应用名称”、“版本号”和“发布者”。
5.  点击“注册到系统”按钮即可。

### 方法二：拖拽注册

1.  将你绿色软件的主程序文件（如 `Notepad++.exe`）直接**拖拽到 `GreenAppRegistrar.exe` 的图标上**。
2.  松开鼠标，程序会自动以后台模式完成注册，无需任何额外操作！

### 方法三：使用命令行接口 (CLI)

对于高级用户和自动化场景，我们提供了功能完整的命令行接口。

**1. 列出所有已注册的应用**
```bash
GreenAppRegistrar.exe list
```

**2. 注册一个新应用**
```bash
# 提供所有详细信息
GreenAppRegistrar.exe register --path "C:\Apps\VSCode\Code.exe" --name "VS Code Portable" --ver "1.90.1" --pub "Microsoft"

# 只提供必要路径，让程序自动探测信息
GreenAppRegistrar.exe register --path "C:\Apps\7-Zip\7zFM.exe"

# 如果应用已存在，强制覆盖注册
GreenAppRegistrar.exe register --path "C:\Apps\7-Zip\7zFM.exe" --force
```

**3. 反注册一个应用**
```bash
# 使用在“应用和功能”中显示的完整名称
GreenAppRegistrar.exe unregister --name "VS Code Portable"
```

**4. 获取帮助**
```bash
GreenAppRegistrar.exe --help
GreenAppRegistrar.exe register --help
```

## 🛠️ 从源码构建

如果你想自行修改或编译，请按以下步骤操作：

1.  确保你已安装 Python (3.6+) 和 Git。
2.  安装项目依赖 (主要是 `winshell` 和 `pefile`):
    ```bash
    pip install winshell pefile
    ```
3.  克隆本仓库:
    ```bash
    git clone https://github.com/3aKHP/GreenAppRegistrar.git
    cd GreenAppRegistrar
    ```
4.  使用 PyInstaller 进行编译。由于程序同时需要 GUI 和命令行输出，推荐使用以下命令：
    ```bash
    # 假设你的主文件是 src/main.py，图标在 assets/icon.ico
    pyinstaller --onefile --name GreenAppRegistrar --icon=assets/icon.ico src/main.py
    ```
    *   注意：我们**没有**使用 `--windowed` 或 `--noconsole` 标志。这会生成一个“控制台”应用，它在需要时可以显示命令行输出，在没有命令行参数时则正常启动 GUI，是此类混合应用的理想选择。

## 📄 许可证

本项目基于 [MIT License](LICENSE) 开源。
```

