# PPT Skill Generator

## Skill 名称

PPT Skill Generator

## 适用场景

当用户需要根据主题、页数、风格、受众、用途和内容要求快速生成可编辑 PowerPoint 文件时使用本 Skill。适用于：

- 商务汇报
- 科技产品介绍
- 教育课件
- 营销方案
- 极简风演示
- 项目计划和总结汇报

## 输入格式

用户可以用自然语言描述需求，也可以提供结构化参数：

```text
主题：AI产品发布会
页数：10
风格：科技风
受众：投资人
用途：路演介绍
内容要求：展示产品定位、核心能力、市场机会、商业模式、发布节奏和下一步行动
```

命令行参数：

```bash
python -m src.cli --topic "AI产品发布会" --pages 10 --style "科技风" --audience "投资人" --purpose "路演介绍" --requirements "展示产品定位、核心能力、市场机会、商业模式、发布节奏和下一步行动"
```

## 输出格式

输出一个 `.pptx` 文件，默认保存到 `output/` 文件夹。

示例输出：

```text
output/ai-product-launch.pptx
```

## 生成 PPT 的执行步骤

1. 解析用户输入，确认主题、页数、风格、受众、用途和内容要求。
2. 调用 `src/outline_generator.py` 生成规则化 PPT 大纲。
3. 调用 `src/ppt_generator.py` 将大纲渲染成 PowerPoint 文件。
4. 将生成文件保存到 `output/`。
5. 向用户返回生成文件路径和核心参数摘要。

## 约束条件

- 使用 `python-pptx` 生成真实 `.pptx` 文件，不输出伪代码。
- 默认包含封面页、目录页、内容页、总结页。
- 支持商务风、科技风、教育风、营销风、极简风。
- 当前版本使用规则生成大纲，后续可在 `outline_generator.py` 接入大模型 API。
- 图片、图表、模板能力预留在 `ppt_generator.py` 的结构中，后续可扩展。
- 输出文件应保存到 `output/`，不要提交生成的 `.pptx` 文件。

## 示例输入

```text
请生成一个 10 页 PPT，主题是 AI产品发布会，风格科技感，受众是投资人，用途是路演介绍，内容要包含产品定位、市场机会、核心能力、商业模式和下一步计划。
```

## 示例输出

```text
Generated PPT:
output/ai-product-launch.pptx

Slides:
1. 封面
2. 目录
3. 背景与目标
4. 核心洞察
5. 方案结构
6. 执行路径
7. 资源与节奏
8. 风险与应对
9. 总结
10. 结束页
```
