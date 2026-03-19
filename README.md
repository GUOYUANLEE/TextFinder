# TextFinder - 全文本搜索工具

一个类似 Anytxt Searcher 的 Windows 桌面全文本搜索应用，基于 C# + WPF 开发。

## 功能特性

- 📁 **文件夹/文件搜索** - 支持选择文件夹或特定文件进行搜索
- 🔍 **全文索引** - 使用 SQLite FTS5 建立高速全文索引
- ⚡ **快速搜索** - 支持关键词高亮显示
- 📄 **多格式支持**
  - 文本文件: txt, md, py, js, ts, html, css, json, xml, cs, java, cpp, c, h, go, rs, sql 等
  - Office 文档: docx, xlsx, pptx
  - PDF 文档
- 🎯 **双击打开** - 双击结果可直接打开对应文件
- 💾 **路径记忆** - 自动记住上次使用的搜索路径

## 使用方法

### 1. 编译项目

```bash
cd TextFinder
dotnet build
```

### 2. 运行程序

```bash
dotnet run
```

### 3. 使用流程

1. 点击"浏览..."按钮选择要索引的文件夹
2. 点击"重建索引"按钮建立全文索引（首次使用或文件夹内容变更时）
3. 在搜索框输入关键词
4. 点击"搜索"或按 Enter 开始搜索
5. 双击结果打开对应文件

## 项目结构

```
TextFinder/
├── TextFinder.csproj    # 项目文件
├── App.xaml            # 应用程序入口
├── App.xaml.cs
├── MainWindow.xaml     # 主窗口界面
├── MainWindow.xaml.cs  # 主窗口逻辑
└── README.md
```

## 技术栈

- .NET 8.0
- WPF (Windows Presentation Foundation)
- SQLite + FTS5 (全文搜索)
- PdfPig (PDF 解析)
- DocumentFormat.OpenXml (Office 文档解析)

## 注意事项

- 首次使用建议先选择文件夹并建立索引
- 索引文件存储在 `%LOCALAPPDATA%\TextFinder\index.db`
- 搜索结果默认显示前 500 条
