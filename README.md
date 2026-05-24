# AI外呼用户反应测试

这个项目用于批量测试移动套餐升档外呼话术。脚本会让三个模型协同完成一次自动化压测：

- A模型：扮演AI客服，按 `prompts/A_v*.txt` 进行外呼沟通。
- C模型：扮演用户，按 `prompts/C_v*.txt` 和Excel场景动态回复。
- B模型：扮演质检员，按 `prompts/B_v1.txt` 审计完整对话并输出JSON结果。

输入文件是 `AI外呼用户反应测试记录表.xlsx`。输出文件会写入 `results/`，该目录默认不提交到Git。

如果需要同时观察多个C用户人设如何刁难AI，可以使用 `--persona-c-list` 一次运行多个C人设。输出Excel会为每个C人设生成一个独立工作表，例如 `C_v1测试结果`、`C_v2测试结果`，便于逐个查看完整通话。不合格行会用浅红底高亮，异常行会用浅橙底高亮。

## 目录说明

```text
.
├── ai_test_script.py                 # 主测试脚本
├── AI外呼用户反应测试记录表.xlsx      # 场景表
├── prompts/
│   ├── A/                            # AI客服人设 A_v1.txt ... A_vN.txt
│   ├── B/                            # 质检员人设 B_v1.txt
│   ├── C/                            # 模拟用户人设 C_v1.txt ... C_v5.txt
│   └── archive/                      # 历史版本归档（不参与运行，默认不入Git）
├── changelog.json                    # 回归对比摘要
└── results/                          # 测试输出，默认忽略
```

C 人设文件首行 `# 简介: xxx` 会被脚本读出来填到结果 Excel 的顶部表头，让人一眼看出当前 sheet 测的是哪种性格的用户。新增 C 人设时按这个格式加一行简介即可。

## 环境准备

安装依赖：

```powershell
pip install -r requirements.txt
```

复制配置模板：

```powershell
Copy-Item config.example.json config.local.json
```

然后编辑 `config.local.json`，填入 `api_key` 和常用参数。`config.local.json` 已加入 `.gitignore`，不会提交到Git。

配置文件示例：

```json
{
  "api_key": "你的API Key",
  "base_url": "https://ark.cn-beijing.volces.com/api/v3",
  "model_a": "doubao-1-5-pro-32k-250115",
  "model_b": "deepseek-v3-2-251201",
  "model_c": "doubao-1-5-pro-32k-250115",
  "concurrency": 10,
  "persona_concurrency": 2,
  "max_retries": 2,
  "retry_delay": 3,
  "max_rounds": 8,
  "input_file": "AI外呼用户反应测试记录表.xlsx",
  "persona_a": "",
  "persona_b": "",
  "persona_c": "C_v1",
  "persona_c_list": ["C_v1", "C_v2", "C_v3", "C_v4", "C_v5"]
}
```

配置优先级：命令行参数 > `config.local.json` > 环境变量 > 默认值。

## 常用命令

全量单版本测试：

```powershell
python ai_test_script.py
```

指定A/B/C人设：

```powershell
python ai_test_script.py --persona A_v2 B_v1 --persona-c C_v5
```

一次运行多个C用户人设，并为每个C人设生成独立工作表：

```powershell
python ai_test_script.py --persona-c-list C_v1 C_v2 C_v3 C_v4 C_v5
```

只跑前3个分类，适合调试：

```powershell
python ai_test_script.py --limit 3
```

随机抽样10个分类，适合每日快速回归：

```powershell
python ai_test_script.py --sample-size 10 --sample-seed 20260523
```

只生成A/C对话，不调用B模型质检，适合快速检查话术是否能跑通：

```powershell
python ai_test_script.py --limit 3 --skip-audit
```

回归对比：

```powershell
python ai_test_script.py --regression --old A_v1 B_v1 --new A_v2 B_v1 --persona-c C_v1
```

## 测试时间太久的优化方案

当前耗时主要来自模型调用次数。一个场景通常包含多轮 A/C 对话，结束后还会调用一次 B 质检；回归模式还会同时跑旧版和新版，所以总调用量会明显放大。

如果重点是观察不同C用户怎么刁难AI，建议保留C模型实时生成，并用 `--persona-c-list` 对比多个C人设。调试阶段可以加 `--skip-audit` 先跳过B质检，这样仍能看到各C人设的完整通话，同时少掉每个场景每个C人设的一次B模型调用。

优先使用这些参数控制耗时：

- `--limit N`：按Excel顺序只跑前N个分类，适合开发调试。
- `--sample-size N --sample-seed S`：随机抽样N个分类，适合稳定复现的快速回归。
- `--skip-audit`：跳过B模型质检，只验证A/C对话生成，能省掉每个场景一次B模型调用。
- `--max-rounds N`：降低单通对话最大轮次，例如先用 `--max-rounds 4` 做冒烟测试。
- `--concurrency N`：提高并发可以缩短墙钟时间，但会增加限流风险；如果接口报限流，反而要降低并发。
- `--persona-concurrency N`：多个C人设并发运行。总并发大约是 `concurrency * persona_concurrency`，设置太高容易触发限流。

推荐分层策略：

1. 本地改prompt时先跑：`--limit 3 --skip-audit --max-rounds 4`
2. 提交前快速回归跑：`--sample-size 10 --sample-seed 固定值 --max-rounds 6`
3. 关键版本发布前再跑全量：不加 `--limit` 和 `--sample-size`

每次测试结束时，控制台都会输出本次测试总耗时。多C人设测试还会额外输出每个C人设单独耗时，便于判断瓶颈。

## 推荐测试工作流

1. 改动prompt后先做冒烟测试，确认脚本和话术能跑通：

```powershell
python ai_test_script.py --limit 2 --persona-c-list C_v1 C_v2 C_v3 C_v4 C_v5 --skip-audit --max-rounds 4 --persona-concurrency 5 --concurrency 2
```

目标：快速看到不同C用户的完整通话，不花时间跑B质检。通常适合几十秒内完成，具体取决于模型响应速度。

2. 发现问题并修完后，跑小范围验收：

```powershell
python ai_test_script.py --limit 5 --persona-c-list C_v1 C_v2 C_v3 C_v4 C_v5 --persona-concurrency 3 --concurrency 3
```

目标：保留B质检，检查关键场景是否合格。不合格行会在对应C人设工作表里标红。这个模式比全量快，但能看到真实质检结果。

3. 如果只改了某类场景，可以用 Excel 顺序配合 `--limit` 先覆盖相关前置场景；更严谨的后续优化是增加按场景名过滤。

4. 发布前跑完整回归：

```powershell
python ai_test_script.py --persona-c-list C_v1 C_v2 C_v3 C_v4 C_v5 --persona-concurrency 2 --concurrency 5
```

目标：覆盖所有场景和C人设。这个模式调用量最大，应优先在网络稳定、额度充足时运行。

5. 如果遇到 `Connection error` 或限流，优先降低并发：

```powershell
python ai_test_script.py --limit 5 --persona-c-list C_v1 C_v2 C_v3 C_v4 C_v5 --persona-concurrency 1 --concurrency 3
```

进一步优化方向：

- 对回归测试固定C模型用户轨迹，让旧版和新版面对相同用户回复，减少随机性并节省C模型调用。
- 为B模型质检增加批量审计模式，把多个短对话合并成一次请求，降低质检调用次数。
- 保存成功场景的对话缓存，只在A/B/C prompt或场景内容变化时重跑。
- 增加失败重跑队列，避免一次限流导致整个场景直接失败。

## 注意事项

- 不要把真实API Key写进源码或提交到Git。
- 中文文件请按UTF-8读取；PowerShell默认读取方式可能显示乱码。
- `results/` 目录输出Excel可能很大，默认不纳入版本管理。
