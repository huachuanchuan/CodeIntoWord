# CodeIntoWord

![CodeIntoWord 图标](assets/logo.png)

CodeIntoWord 是一个 Windows 桌面版 Word 插件。在 Word 的“代码美化”选项卡打开侧边栏，输入或读取选中的代码，选择语言、字号、配色与行号，再把可编辑的代码表格插入光标位置。

## 功能

- 两组可选外观：CodeInWord 风格和 Easy Code Formatter 风格。
- 20 多种语言选项与轻量语法高亮；支持自动识别、实时预览、行号和空白行开关。
- 插入结果是 Word 表格和文字，可以继续编辑，不是图片。
- 代码在本机处理，不上传到网站或服务器。

这两组外观是独立实现的参考样式；本项目不调用相应网站，也不属于这两个产品。

## 安装

1. 确认使用 **Windows、64 位桌面版 Microsoft Word**。当前安装包还依赖本机的 .NET Framework 4.x 和 Office 互操作组件；它们没有打包进 EXE。
2. 保存文档并完全关闭 Word。
3. 下载并运行 [CodeStudio-Word-Setup.exe](dist/CodeStudio-Word-Setup.exe)。安装作用于当前 Windows 用户，不需要管理员权限。
4. 重新打开 Word，在“代码美化”选项卡点击“插入美化代码”。

如需卸载，关闭 Word 后在 PowerShell 中运行：

```powershell
.\CodeStudio-Word-Setup.exe /uninstall
```

当前安装包没有自动检测所有 Office 配置。32 位 Word 或缺少所需 Office/.NET 组件的电脑可能无法加载插件。EXE 没有商业代码签名，Windows 可能显示安全提示；请从本仓库下载并自行核对来源。

## 从源码构建

需要 Windows、PowerShell、Visual Studio Build Tools（含 C# 编译器）以及已安装的 Word/Office 互操作组件。构建脚本会在 `build/` 下生成本机专用的强名称密钥，密钥不会提交到仓库。

```powershell
.\build.ps1
```

生成文件：

```text
build/CodeStudio.Addin.dll
build/CodeStudio-Word-Setup.exe
```

测试会启动独立的 Word 实例并创建临时文档；运行前请先关闭已有 Word 窗口：

```powershell
.\test.ps1
```

## 代码结构

- `src/Connect.cs`：Word 功能区按钮、自定义图标与侧边栏加载。
- `src/WordCodePane.cs`：侧边栏界面、Word 选区读取与插入。
- `src/CodeRenderer.cs`：语法着色和 Word 可导入的 HTML 表格。
- `src/Setup.cs`：当前用户安装、COM 注册与卸载。
- `tests/TestInsert.cs`：临时 Word 文档中的插入与换行验证。

欢迎提交问题和改进建议。
