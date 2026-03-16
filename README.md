论文小标题自动提取工具 (Word VBA Macro)

<div align="center">

一键提取Word文档中的加粗小标题，快速生成文档大纲

#-功能特性 • #-快速开始 • #💡-基础示例代码 • #💻-完整版核心代码 • #⚠️-注意事项

</div>

📋 项目简介

这款VBA宏工具专门为学术研究者、论文作者和编辑设计，能够自动从Word文档中提取所有以加粗格式呈现的小标题和重点内容。通过智能过滤和多种输出格式，帮助用户快速了解文档结构，提高文献管理和写作效率。

✨ 功能特性

功能 基础版 完整版

🔍 提取加粗文本 ✅ ✅

🎯 智能过滤 ❌ ✅

📄 多格式输出 ❌ ✅

📊 层级识别 ❌ ✅

📁 批量处理 ❌ ✅

🎨 用户界面 ❌ ✅

⚡ 易用性 简单 中等

🚀 快速开始

系统要求

• Microsoft Word 2016+ 或 Office 365

• 启用宏功能（文件 → 选项 → 信任中心 → 宏设置 → 启用所有宏）

安装方法

1. 基础版 - 简单复制粘贴即可使用
2. 完整版 - 下载项目文件并导入VBA模块

基本使用步骤

1. 打开需要提取的Word文档
2. 按 Alt + F11 打开VBA编辑器
3. 插入新模块，复制代码
4. 按 F5 运行

💡 基础示例代码

如果你只需要最简单的功能，这个20行代码的版本足够使用

'==================================================
' 基础版：提取所有加粗文本
' 文件名：ExtractBoldText.bas
' 功能：提取当前Word文档中所有加粗文本到新文档
'==================================================

Sub ExtractBoldText()
    Dim doc As Document, newDoc As Document
    Dim rng As Range
    Set doc = ActiveDocument
    Set newDoc = Documents.Add
   
    For Each rng In doc.StoryRanges
        Set rng = rng.Duplicate
        rng.Find.ClearFormatting
        rng.Find.Font.Bold = True
        rng.Find.Format = True
       
        Do While rng.Find.Execute
            newDoc.Range.InsertAfter rng.Text & vbCrLf
            rng.Collapse wdCollapseEnd
        Loop
    Next rng
   
    newDoc.Activate
    MsgBox "已提取所有加粗文本到新文档！"
End Sub


基础版使用说明

1. 代码复制：将上方代码复制到VBA编辑器的模块中
2. 运行：按 F5 或点击运行按钮
3. 输出：自动创建新文档，每行一个加粗文本
4. 注意：此版本会提取文档中所有加粗内容，包括页眉、页脚、文本框等
