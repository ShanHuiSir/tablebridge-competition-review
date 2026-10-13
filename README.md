# 表桥 · 参赛项目与复盘归档

跨系统表格数据自动化桥接与处理智能体。这个私人仓库保存表桥的可编辑源码、模拟样例、评审视频、参赛材料与验证记录，方便回顾和复现。归档日期：2026-10-13；不代表已经报名或提交比赛平台。

## 先看这些

- [评审演示视频（3 分 03 秒，中文声音和字幕）](artifacts/跨系统表格数据自动化桥接与处理智能体-演示视频.mp4)：下载后播放，GitHub 可能不直接预览。
- [解说词](artifacts/演示解说词.md) · [字幕 SRT](artifacts/表桥_演示字幕.srt) · [视频验收报告](artifacts/视频验收报告.json)
- [参赛材料整包](artifacts/表桥_参赛材料整包.zip) · [提交说明](artifacts/提交说明.md)
- [原有申报底稿](artifacts/跨系统表格数据自动化桥接与处理智能体-申报书.docx)：人力行政场景底稿，尚未改写成新版通用智能体申报书。
- [版本与验证范围](artifacts/版本与验证范围说明.md) · [复盘笔记](docs/复盘笔记.md)
- [工程详细说明](tablebridge/README.md) · [原始命题文档](docs/参考命题/跨系统表格数据转换智能体人力行政.docx)

## 实现与实际结果

上传源表和目标模板，以自然语言描述转换规则；DeepSeek 根据实际表头生成计划和受限 Python 清洗代码，隔离执行后校验结果。人工解决异常后导出 XLSX、真实 XLS 2003 或指定编码 CSV，并保存流程供复用。原有人事请假结算、加班筛选工作台在 `/legacy`。

演示全程使用虚构员工数据：4 条中文 XLSX 记录 + 3 条英文 GBK CSV 记录，剔除 1 条负金额、去重 1 条，输出 **5 条**。人工将工号 `0004` 的异常手机号核实为模拟值 `13800138001`，未解决异常从 **1 → 0** 后才允许导出。视频中的 [实际导出成果](artifacts/视频实测导出成果.zip) 包含 `import.xls`、`audit.json`、`workflow.py`。

2026-10-12 已验证真实 DeepSeek 模型调用、变化表头规划、执行、异常恢复、三种导出格式回读及输入数值扰动。请求模型为 `deepseek-chat`，供应商响应模型标识为 `deepseek-flash`。见 [模型实测](tablebridge/docs/通用Agent模型实测.json)、[GUI 实测](tablebridge/docs/通用Agent模型GUI实测.json)、[验收记录](tablebridge/docs/通用转换验收记录.md)。现有自动化测试共 **62 项**。

## 本地运行

需要 Python 3.11+；macOS / Python 3.12 已验证。

```bash
cd tablebridge
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
cp .env.example .env
# 在本机 .env 中填写自己的模型密钥
.venv/bin/python app.py
```

打开 <http://127.0.0.1:8877>。模型功能需要有效密钥和网络；载入模拟案例可以先体验本地转换。

```bash
cd tablebridge
.venv/bin/python -m unittest discover -s tests -v
```

Windows 启动和构建脚本已提供，操作说明见工程 README；Windows 实机、EXE 构建和正式 HR/ERP 系统导入尚未验收。本版本为本机单人应用，没有企业多人权限和系统 API 集成。

## 归档结构

| 路径 | 内容 |
| --- | --- |
| `tablebridge/` | 可编辑源码、测试、模拟样例和工程说明 |
| `artifacts/` | 视频、字幕、解说词、申报底稿、源码 ZIP、成果 ZIP、参赛整包和验证报告 |
| `docs/` | 复盘笔记、原始命题、归档检查记录 |
| `SHA256SUMS.json` | 仓库文件校验；不含自身和 Git 元数据 |

源码 ZIP 与可编辑源码对应同一已验收版本。2026-10-13 归档时同步改名后的视频、申报书文件名和相关说明，并重建参赛整包；视频和申报书正文未改动。未收录其他比赛项目、重复解压目录、原始录制缓存、虚拟环境、API 密钥或本机运行数据。

Git 首次提交是完整归档快照，标签 `review-2026-10-13` 用于定位这个版本；没有重建或虚构之前的开发提交。
