# AI外呼用户反应测试

这个项目用于批量测试移动套餐升档外呼话术。脚本会让三个模型协同完成一次自动化压测：

- A模型：扮演AI客服，按 `prompts/A_v*.txt` 进行外呼沟通。
- C模型：扮演用户，按 `prompts/C_v*.txt` 和Excel场景动态回复。
- B模型：扮演质检员，按 `prompts/B_v1.txt` 审计完整对话并输出JSON结果。

输入文件是 `AI外呼用户反应测试记录表.xlsx`。输出文件会写入 `results/`，该目录默认不提交到Git。

## 目录说明

```text
.
├── ai_test_script.py                 # 主测试脚本
├── AI外呼用户反应测试记录表.xlsx      # 场景表
├── prompts/                          # A/B/C模型人设和质检规则
├── changelog.json                    # 回归对比摘要
└── results/                          # 测试输出，默认忽略
```

## 环境准备

安装依赖：

```powershell
pip install -r requirements.txt
```

设置火山方舟API Key：

```powershell
$env:ARK_API_KEY="你的API Key"
```

可选环境变量：

```powershell
$env:ARK_BASE_URL="https://ark.cn-beijing.volces.com/api/v3"
$env:MODEL_A="doubao-1-5-pro-32k-250115"
$env:MODEL_B="deepseek-v3-2-251201"
$env:MODEL_C="doubao-1-5-pro-32k-250115"
$env:TEST_CONCURRENCY="10"
$env:MAX_ROUNDS="8"
$env:MAX_RETRIES="2"
```

## 常用命令

全量单版本测试：

```powershell
python ai_test_script.py
```

指定A/B/C人设：

```powershell
python ai_test_script.py --persona A_v2 B_v1 --persona-c C_v5
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

优先使用这些参数控制耗时：

- `--limit N`：按Excel顺序只跑前N个分类，适合开发调试。
- `--sample-size N --sample-seed S`：随机抽样N个分类，适合稳定复现的快速回归。
- `--skip-audit`：跳过B模型质检，只验证A/C对话生成，能省掉每个场景一次B模型调用。
- `--max-rounds N`：降低单通对话最大轮次，例如先用 `--max-rounds 4` 做冒烟测试。
- `--concurrency N`：提高并发可以缩短墙钟时间，但会增加限流风险；如果接口报限流，反而要降低并发。

推荐分层策略：

1. 本地改prompt时先跑：`--limit 3 --skip-audit --max-rounds 4`
2. 提交前快速回归跑：`--sample-size 10 --sample-seed 固定值 --max-rounds 6`
3. 关键版本发布前再跑全量：不加 `--limit` 和 `--sample-size`

进一步优化方向：

- 对回归测试固定C模型用户轨迹，让旧版和新版面对相同用户回复，减少随机性并节省C模型调用。
- 为B模型质检增加批量审计模式，把多个短对话合并成一次请求，降低质检调用次数。
- 保存成功场景的对话缓存，只在A/B/C prompt或场景内容变化时重跑。
- 增加失败重跑队列，避免一次限流导致整个场景直接失败。

## 注意事项

- 不要把真实API Key写进源码或提交到Git。
- 中文文件请按UTF-8读取；PowerShell默认读取方式可能显示乱码。
- `results/` 目录输出Excel可能很大，默认不纳入版本管理。
