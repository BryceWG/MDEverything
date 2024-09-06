
# 多格式转换Markdown工具 (MDEverything)

🚀 这是一个强大的多格式文件转换工具，可以将各种格式的文件转换为Markdown格式。

## 🌟 特性

- 支持多种文件格式转换：PDF、Excel、Word、PowerPoint、HTML等
- 用户友好的图形界面
- 可自定义转换选项
- 支持批量转换
- 支持网页链接转换

## 🛠️ 安装

1. 克隆此仓库：
   ```
   git clone https://github.com/BryceWG/MDEverything.git
   ```
2. 进入项目目录：
   ```
   cd MDEverything
   ```
3. 安装依赖：
   ```
   pip install -r requirements.txt
   ```

## 🚀 使用方法

1. 运行主程序：
   ```
   python main.py
   ```
2. 在图形界面中选择要转换的文件或输入URL
3. 选择输出目录，默认为源文件目录
4. 根据需要调整转换选项
5. 点击"转换"按钮开始转换

## 📚 详细使用方法

### 📄 PDF 转换

1. 选择 PDF 文件
2. 设置选项：
   - DPI：设置图像分辨率（72/144/216）
   - Parse Mode：选择解析模式（auto/scan）
   - 生成标题：是否自动生成文档结构
   - 表格格式：选择 md 或 html 格式
   - 获取图片：选择是否提取图片（none/page/objects/both）
   - 起始页和页数：设置转换的页面范围
3. 输入 API 凭证（app_id 和 secret_code）
4. API 凭证获取和相关调用方式请访问[TextIn通用文档解析](https://www.textin.com/document/pdf_to_markdown)

### 📊 Excel 转换

1. 选择 Excel 文件
2. 在树状视图中选择要转换的工作表
3. 设置选项：
   - 第一行为表头：勾选此项将第一行视为表头

### 🌐 HTML 转换

1. 输入网页 URL 或选择本地 HTML 文件
2. 设置选项：
   - 可选使用 [Jina AI](https://jina.ai/)：启用 AI 辅助转换（需要 API 密钥）
   - 忽略链接：不包含超链接
   - 忽略图片：不包含图片
   - 正文宽度：设置转换后的文本宽度
3. 可选：加载链接网页内其他链接信息，可选择指定网页链接进行转换，实现一次爬取多个网页

### 📝 Word 转换

1. 选择 Word 文档（.docx）
2. 无需额外设置，直接转换

### 📊 PowerPoint 转换

1. 选择 PowerPoint 文件（.pptx）
2. 设置选项：
   - 图片最大宽度：限制转换后的图片宽度
   - 禁用图片提取：不包含图片
   - 不转义特殊字符：保留原始字符
   - 不添加演讲者注释：忽略幻灯片备注
   - 禁用颜色标签：不包含颜色信息
   - 启用幻灯片分隔符：在幻灯片之间添加分隔符
   - 最小文本块大小：设置最小文本块的大小
   - 输出格式：选择输出格式（markdown/wiki/mdk/qmd）

### 📑 LaTeX 转换

1. 选择 LaTeX 文件（.tex）
2. 需要安装 [Pandoc](https://pandoc.org/installing.html) 才能使用此功能

### 🔄 Markdown 合并

1. 选择多个 Markdown 文件
2. 点击"合并选中的 Markdown 文件"按钮
3. 选择输出目录

## ⚙️ 通用设置

- 输出目录：选择转换后文件的保存位置
- 转换按钮：开始转换过程

## 💡 提示

- 对于批量转换，可以选择多个文件同时进行转换
- 转换过程中请保持耐心，特别是对于大文件或复杂文档
- 转换完成后，检查输出文件以确保内容正确

## To-do List 

- [ ] 并行处理加速
- [ ] 简易AI处理功能
- [ ] 优化输出排版
- [ ] 增加图像文件支持
- [ ] 完善.docx .html格式图片储存

## 📚 依赖库

- PyQt6：用于创建图形用户界面
- requests：用于处理HTTP请求
- BeautifulSoup4：用于解析HTML
- python-docx：用于处理Word文档
- openpyxl：用于处理Excel文件
- python-pptx：用于处理PowerPoint文件
- html2text：用于HTML到Markdown的转换
- pypandoc：用于文档格式转换

## ⚠️ 注意事项

- 确保您有足够的权限读取源文件和写入目标目录
- 对于PDF转换，需要提供有效的API密钥（app_id和secret_code）
- 转换大文件可能需要较长时间，请耐心等待
- 某些复杂格式的文件可能无法完美转换，可能需要手动调整
- 使用HTML转换功能时，请确保您有合法权限访问和转换目标网页内容

## 🤝 贡献

欢迎提交问题和拉取请求。对于重大更改，请先开issue讨论您想要改变的内容。

## 📄 许可证

[MIT](https://choosealicense.com/licenses/mit/)
