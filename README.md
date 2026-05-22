# PPT Skill Generator

PPT Skill Generator 是一个可复用的 Codex Skill 项目，用于根据用户输入的主题、页数、风格、受众、用途和内容要求，自动生成 PPT 大纲并输出真实可编辑的 `.pptx` 文件。

当前版本使用规则生成大纲，使用 `python-pptx` 生成 PowerPoint 文件。代码结构保留了后续扩展图片、图表、模板和大模型 API 的空间。

## 功能

- 根据主题自动生成 PPT 大纲
- 自动生成每页标题、正文和要点
- 使用 `python-pptx` 生成真实 `.pptx`
- 默认包含封面页、目录页、内容页、总结页
- 支持商务风、科技风、教育风、营销风、极简风
- 输出文件保存到 `output/`

## 安装方法

```bash
pip install -r requirements.txt
```

建议在虚拟环境中安装：

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Windows PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## 使用方法

从仓库根目录运行：

```bash
python -m src.cli --topic "AI产品发布会" --pages 10 --style "科技风" --audience "投资人"
```

加入用途和内容要求：

```bash
python -m src.cli \
  --topic "AI产品发布会" \
  --pages 10 \
  --style "科技风" \
  --audience "投资人" \
  --purpose "路演介绍" \
  --requirements "展示产品定位、核心能力、市场机会、商业模式、发布节奏和下一步行动"
```

指定输出文件：

```bash
python -m src.cli --topic "年度经营复盘" --pages 8 --style "商务风" --audience "管理层" --output output/business-review.pptx
```

## 示例命令

```bash
python -m src.cli --topic "AI产品发布会" --pages 10 --style "科技风" --audience "投资人" --purpose "路演介绍"
```

成功后会看到类似输出：

```text
Generated PPT: output/ai-product-launch.pptx
Slides: 10
Style: 科技风
Audience: 投资人
```

## 如何让 Codex 调用这个 Skill

把本仓库放入 Codex Skills 目录，或让 Codex 在当前仓库中读取 `skill.md`。

推荐调用方式：

```text
使用 PPT Skill Generator，根据我的需求生成 PPT：
主题【AI产品发布会】，页数【10页】，风格【科技风】，受众【投资人】，用途【路演介绍】，内容要求【展示产品定位、核心能力、市场机会、商业模式和下一步行动】，输出到【output/ai-product-launch.pptx】。
```

Codex 应执行：

```bash
python -m src.cli --topic "AI产品发布会" --pages 10 --style "科技风" --audience "投资人" --purpose "路演介绍" --requirements "展示产品定位、核心能力、市场机会、商业模式和下一步行动" --output output/ai-product-launch.pptx
```

## 一键调用口令

```text
请使用 PPT Skill Generator。根据我的主题、页数、风格、受众、用途和内容要求生成一个可编辑 PPTX。请先读取 skill.md，然后运行 python -m src.cli，并把输出文件保存到 output/。我的需求是：主题【在这里填写】，页数【在这里填写】，风格【商务风/科技风/教育风/营销风/极简风】，受众【在这里填写】，用途【在这里填写】，内容要求【在这里填写】。
```

## 项目结构

```text
ppt-skill-generator/
├── README.md
├── requirements.txt
├── skill.md
├── src/
│   ├── __init__.py
│   ├── ppt_generator.py
│   ├── outline_generator.py
│   └── cli.py
├── examples/
│   └── example_request.md
├── output/
│   └── .gitkeep
└── .gitignore
```
