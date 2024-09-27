# AI PPT 生成器

AI PPT 生成器是一个使用人工智能技术自动生成 PowerPoint 演示文稿的工具。它利用 GPT-3.5 模型生成内容，并使用 PyQt5 构建图形用户界面。

## 功能特点

- 基于用户输入的主题自动生成 PPT 内容
- 可自定义幻灯片数量（5-50 张）
- 实时预览生成的内容
- 一键导出为 PowerPoint 文件
- 简洁直观的用户界面

## 安装

1. 克隆此仓库：
   ```bash
   git clone https://github.com/longshanji/ai-ppt-generator.git
   cd ai-ppt-generator
   ```

2. 安装依赖：
   ```bash
   pip install -r requirements.txt
   ```

3. 创建 `config.ini` 文件并添加以下内容：
   ```ini
   [API]
   OPENROUTER_API_KEY = 您的OpenRouter API密钥
   YOUR_SITE_URL = 您的网站URL
   YOUR_APP_NAME = 您的应用程序名称
   ```

## 使用说明

### 从源代码运行

1. 运行程序：
   ```bash
   python ai_ppt_generator.py
   ```

2. 在主界面中输入 PPT 主题。
3. 使用滑块选择幻灯片数量。
4. 点击"生成 PPT"按钮。
5. 等待内容生成完成，在预览窗口查看内容。
6. 点击"导出 PPT"按钮保存为 PowerPoint 文件。

### 使用打包后的可执行文件

1. 双击运行 `AI PPT生成器.exe`.
2. 按照上述步骤 2-6 操作。

## 打包说明

如果您想创建独立的可执行文件，请按照以下步骤操作：

1. 确保已安装 PyInstaller：
   ```bash
   pip install pyinstaller
   ```

2. 使用提供的 .spec 文件进行打包：
   ```bash
   pyinstaller ai_ppt_generator.spec
   ```

3. 打包完成后，可执行文件将位于 `dist` 目录中。

## 界面说明

- **PPT 主题**：输入 PPT 主题
- **幻灯片数量**：选择幻灯片数量（5-50 张）
- **生成 PPT**：开始生成过程
- **状态栏**：显示当前操作状态
- **进度条**：显示生成进度
- **内容预览**：显示生成的内容预览
- **导出 PPT**：保存为 PowerPoint 文件

## 注意事项

- 确保有稳定的网络连接
- 生成过程可能需要一些时间，取决于幻灯片数量和网络状况
- 生成的内容可能需要进一步编辑和优化
- 请确保 `config.ini` 文件配置正确，并与可执行文件位于同一目录

## 技术细节

- 使用 PyQt5 构建 GUI
- 使用 QThread 实现异步 PPT 生成
- 通过 OpenRouter API 使用 GPT-3.5 模型生成内容
- 使用 python-pptx 创建 PowerPoint 文件
- 使用 PyInstaller 打包为独立可执行文件

## 贡献

欢迎提交 issues 和 pull requests。如果您有任何改进意见或发现了 bug，请随时与我们联系。

## 许可证

[MIT License](LICENSE)

## 联系方式

如有任何问题或建议，请联系 [您的邮箱地址]。
