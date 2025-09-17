# 关于此Nuitka编译版本的特别说明 (Developer Notes for the Nuitka Build)

**注意：** 本文件是针对 GreenAppRegistrar 特定Nuitka编译版本的补充开发说明。有关项目的基础功能和通用使用方法，请参阅主 [README.md](README.md)。

这个版本的构建流程与主项目推荐的 PyInstaller 方案有所不同，它采用了 **Nuitka** 进行编译。此举旨在探索并解决原生Python脚本在某些场景下的性能瓶颈。

---

## 1. 为什么要切换到 Nuitka？

我们决定为这个版本引入Nuitka，主要基于以下性能考量：

*   **更快的启动和运行速度：** Nuitka将Python代码直接翻译为C语言并编译成本地机器码，极大地提升了程序的启动速度和运行时性能。这在频繁调用CLI进行**批量注册/反注册**的场景下，效果尤其明显。
*   **更小的文件体积：** 相比PyInstaller打包完整的Python解释器，Nuitka通过更精细的依赖分析和C语言编译，通常能生成体积更小的可执行文件。

## 2. 核心：`manifest.xml` 与 UAC 问题

在迁移到Nuitka的过程中，我们遇到了一个关键的“幽灵BUG”：**程序在请求UAC提权时会静默失败，无法弹出UAC窗口。**

*   **根本原因：** 经深入排查，问题根源在于Nuitka默认编译的 `.exe` 文件缺少一个“**应用程序清单 (Application Manifest)**”。这个清单就像程序的“身份证”，它向Windows系统声明该程序能够感知并正确处理UAC。没有这张“身份证”的程序在尝试提权时，会被系统出于兼容性或安全考虑而直接拒绝。

*   **解决方案：** 我们通过手动创建一个 `manifest.xml` 文件，并在其中定义 `<requestedExecutionLevel level="asInvoker" />`，明确告知Windows我们的程序是UAC感知的。然后，在编译时将这个清单文件嵌入到最终的 `.exe` 中，从而彻底解决了UAC静默失败的问题。

## 3. 如何使用 Nuitka 构建此版本

如果你需要重新编译此版本，请严格遵循以下步骤。

### 准备工作

1.  确保你的Python环境中已安装最新版的 **Nuitka** 及其依赖。
    ```bash
    # 推荐使用conda-forge渠道（如果你使用Anaconda）
    conda install -c conda-forge nuitka
    # 或者使用pip
    python -m pip install --upgrade nuitka
    ```
2.  确保项目根目录下存在 `manifest.xml` 文件。

### 编译命令

使用以下命令进行编译。它已经包含了处理UAC问题、Tkinter依赖和隐藏控制台的所有必要参数。

```bash
python -m nuitka ^
    --onefile ^
    --windows-console-mode=disable ^
    --follow-imports ^
    --plugin-enable=tk-inter ^
    --include-data-files=manifest.xml=manifest.xml ^
    main.py
