# Thesis Agent

`thesis-agent` 是一个面向本科毕业设计论文的自动格式修改、视觉确认和内容质量审阅项目。它以给定论文格式模板为标准，把 Word 结构修复、LibreOffice/PDF 渲染、页面图像检查、内容质量审阅和视觉复核包生成串成可重复流水线。

## 目标

- 以 `samples/templates/论文格式.doc` 或转换后的 `论文格式.docx` 为格式标准。
- 对目标论文进行 Word 结构检查、PDF 渲染、页面图像视觉检查和内容完整性审阅。
- 自动处理 `.doc` 到 `.docx` 转换、前置声明页补齐、目录分页、分节空白页、正文页码、静态目录页码回写、主标题样式、题注与图表锚点等常见问题。
- 从模板 `.docx` 中抽取红字说明，生成 `template_red_checklist.md/json`，作为后续硬规则和人工复核的基线。
- 对测试章节偏薄、缺少致谢等内容问题执行保守补强；对可识别的简单系统架构图执行受控重绘，避免连接线断开。
- 将每次运行的 PDF、PNG、JSON 和 Markdown 报告写入 `runs/`，用于迭代比较。
- 用 `configs/sdju_format.json` 统一维护格式规则和本科论文质量要求，避免把学校/专业要求散落在代码里。

## 样本

公开仓库不提交真实学生论文、学校模板原件和运行产物。完整验收需要在本机私有环境中自行准备模板与论文样本。

这些文件包含学生论文内容，只适合在本机私有环境中测试，因此 `samples/` 已加入 `.gitignore`。公开克隆中没有 `samples/` 时，样本依赖的 pytest 用例会自动跳过；需要完整验收时，把私有样本复制到对应路径后再运行测试。

## 环境

当前机器已配置：

- LibreOffice AppImage：`/home/biostar/.local/bin/soffice-headless`
- Poppler：`pdftoppm`、`pdftotext`、`pdfinfo`
- Windows 中文字体 fontconfig 映射：宋体、黑体、Times New Roman 可被 LibreOffice 使用
- 可选 OfficeCLI：`/home/biostar/work/external/bin/officecli`

检查环境：

```bash
PYTHONPATH=src python3 -m thesis_agent doctor
```

先验证标准模板能否无损复建：

```bash
PYTHONPATH=src python3 -m thesis_agent template-selftest \
  --cover samples/templates/附件15\ 学士学位论文封面.docx \
  --body samples/templates/论文格式.docx \
  --out runs/template-selftest
```

该命令会生成 `standard_template.docx`，并把“封面模板 + 正文格式模板”与重建稿逐页渲染为图片进行视觉级比较。封面、声明页、摘要、目录、正文示例、参考文献和致谢页均先以模板自复现为基准，后续学生内容只能填入该基准模板，不应再把学生原稿的段落行距和分页格式带入标准文档。

生成正式处理模板并去除红字批注：

```bash
PYTHONPATH=src python3 -m thesis_agent template-rebuild \
  --cover samples/templates/附件15\ 学士学位论文封面.docx \
  --body samples/templates/论文格式.doc \
  --output runs/template/standard_template.docx \
  --strip-red
```

正式模板会修正正文前罗马页码连续性：封面不显示页码，学术诚信声明从 `I` 开始，AI 使用情况声明、版权使用授权书、摘要、英文摘要和目录继续顺排；同时清除红色批注形状和红色说明文字，避免学生论文成稿残留模板红字。

使用正式模板进行槽位填充：

```bash
PYTHONPATH=src python3 -m thesis_agent fill-template \
  --template runs/template/standard_template-formal.docx \
  --target samples/drafts/物联网2212-杨钰婷-毕业论文初稿.docx \
  --output runs/template/yang-slot-filled.docx
```

`fill-template` 以正式标准模板为唯一版式来源，学生初稿只提供封面元数据、摘要、关键词、正文、图表、参考文献和致谢等内容。输出文档不再保留学生原稿的段落行距、标题样式和分页格式。

运行一次论文审计：

```bash
PYTHONPATH=src python3 -m thesis_agent audit \
  --template samples/templates/论文格式.doc \
  --target samples/drafts/物联网2212-杨钰婷-毕业论文初稿.docx \
  --out runs/yang-audit
```

报告会生成到：

- `runs/yang-audit/report.md`
- `runs/yang-audit/report.json`
- `runs/yang-audit/pdf/`
- `runs/yang-audit/png/`

运行完整处理闭环：

```bash
PYTHONPATH=src python3 -m thesis_agent process \
  --template samples/templates/论文格式.doc \
  --target samples/drafts/物联网2212-杨钰婷-毕业论文初稿.docx \
  --out runs/process/yang
```

批量处理当前样本：

```bash
PYTHONPATH=src python3 -m thesis_agent batch-process \
  --template samples/templates/论文格式.doc \
  --inputs samples/drafts \
  --out runs/acceptance
```

每篇输出目录包含按 `学号-姓名-论文题目.docx` 命名的最终稿、兼容副本 `final.docx`、`slot_fill_report.json`、`audit/report.md`、`content_enhance_report.json`、`diagram_repair_report.json`、`toc_sync_report*.json`、`content_improvement.md`、`process_report.md`、`vision_review_prompt.md` 和 `vision/` 下的页面联系表。若学号、姓名或论文题目无法识别，对应位置使用 `XX`，例如 `XX-杨钰婷-智能危化品监管系统.docx`。

生成一个保守格式修复稿，并立刻渲染审计：

```bash
PYTHONPATH=src python3 -m thesis_agent fix-format \
  --target samples/drafts/物联网2212-杨钰婷-毕业论文初稿.docx \
  --output runs/fixes/yang-fixed.docx \
  --template samples/templates/论文格式.doc \
  --audit-out runs/fixes/yang-fixed-audit
```

`fix-format` 当前只支持 `.docx`，且不会覆盖原文件。它会清理尾部空白页来源、压缩长空段、合并空分节段、修正目录标题与目录条目分页、将正文页码从 `1 绪论` 开始重置为 1、给主章节标题套用一级标题样式，并对题注应用基本的同页保留规则。
如果提供 `--template`，还会从模板 `.docx` 中移植学术诚信声明、AI 使用情况声明和版权使用授权书等前置页。

生成视觉大模型审阅包：

```bash
PYTHONPATH=src python3 -m thesis_agent vision-pack \
  --audit-dir runs/fixes/yang-fixed-audit \
  --out runs/fixes/yang-vision-pack
```

该命令会生成模板和目标论文的页面联系表、重点页面副本，以及 `vision_review_prompt.md`。这些图片和提示词用于视觉大模型逐页审阅，检查程序化规则难以覆盖的版面差异。

## 开发路线

1. `audit`：稳定发现模板和论文之间的格式、视觉、结构、内容质量差异。
2. `fix-format`：基于模板样式、页眉页脚、分页符、题注、目录和引用域自动修复。
3. `process`：把转换、修复、审计、内容计划、视觉包和质量门禁组成单篇闭环。
4. `batch-process`：把单篇闭环应用到样本集或同一目录下的多篇论文。
5. `agent-loop`：后续继续增强多轮自动重写能力，尤其是参考文献规范化和人工/视觉模型反馈回灌。

## 当前验收

最近一次完整验收命令：

```bash
PYTHONPATH=src python3 -m pytest -q
PYTHONPATH=src python3 -m thesis_agent batch-process \
  --template samples/templates/论文格式.doc \
  --inputs samples/drafts \
  --out runs/acceptance-redcheck
```

结果：`28 passed`；5 篇样本全部通过质量门禁，阻断项为 0。最新新模板流程结果见 `runs/acceptance-new-template/batch_result.json`。当前硬门禁包含一级标题红字规则、目录页码一致性、正文前罗马页码、页眉页码右对齐、空白页、参考文献和致谢格式；同时检查最终 DOCX 不残留红字批注、隐藏 TOC 域和缺失的 Word 命名空间声明。

## 质量依据

本项目采用“二本/应用型本科工科毕业设计”的保守合格线：结构完整、工程问题明确、方案和实现可验证、测试充分、文字规范、图表和参考文献合规、学术诚信风险可控。依据整理见 `docs/quality_requirements.md`。
