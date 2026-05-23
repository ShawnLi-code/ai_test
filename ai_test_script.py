#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AI外呼用户反应测试脚本 v3.0
功能：
- A模型(豆包)扮演AI客服
- B模型(DeepSeek)质检审计
- C模型(豆包)动态扮演用户，单人设+场景注入驱动自然对话
- 人设文件外置，支持版本管理和回归对比
"""

import sys
import io
# 解决Windows控制台GBK编码问题
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

import os
import json
import time
import argparse
import re
import random
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ==================== 配置 ====================
API_KEY = os.getenv("ARK_API_KEY", "")
BASE_URL = os.getenv("ARK_BASE_URL", "https://ark.cn-beijing.volces.com/api/v3")
MODEL_A = os.getenv("MODEL_A", "doubao-1-5-pro-32k-250115")   # A模型-豆包(AI客服)
MODEL_B = os.getenv("MODEL_B", "deepseek-v3-2-251201")         # B模型-DeepSeek(质检)
MODEL_C = os.getenv("MODEL_C", "doubao-1-5-pro-32k-250115")   # C模型-豆包(模拟用户)
CONCURRENCY = int(os.getenv("TEST_CONCURRENCY", "10"))
MAX_RETRIES = int(os.getenv("MAX_RETRIES", "2"))
RETRY_DELAY = int(os.getenv("RETRY_DELAY", "3"))  # 重试间隔(秒)
MAX_ROUNDS = int(os.getenv("MAX_ROUNDS", "8"))  # 单通对话最大轮次

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FILE = os.path.join(SCRIPT_DIR, "AI外呼用户反应测试记录表.xlsx")
PROMPTS_DIR = os.path.join(SCRIPT_DIR, "prompts")
RESULTS_DIR = os.path.join(SCRIPT_DIR, "results")
CHANGELOG_FILE = os.path.join(SCRIPT_DIR, "changelog.json")

# 确保目录存在
os.makedirs(PROMPTS_DIR, exist_ok=True)
os.makedirs(RESULTS_DIR, exist_ok=True)

# ==================== C模型场景上下文 ====================
# 从Excel触发语中提取的场景描述，注入C模型提示中


def build_scene_context(category, triggers):
    """从分类名和触发语构建C模型的场景上下文"""
    triggers_text = "；".join(triggers)
    return f"""# 场景话题
场景：{category}
话题：{triggers_text}
态度：根据场景自然表现，可以是配合、犹豫、拒绝、质疑等"""


def is_conversation_ended(text):
    """检测AI的回复是否包含结束语，表示通话已自然结束"""
    end_signals = [
        "再见", "祝您生活愉快", "祝你生活愉快", "先不打扰你了",
        "先不打扰您了", "稍后会有客户经理联系", "先不打扰了",
    ]
    text_lower = text.lower()
    return any(signal in text_lower for signal in end_signals)


def call_user_model(client, persona_c, scene_context, messages, round_num, max_rounds):
    """调用C模型生成用户回复。
    persona_c: C_v1.txt 人设内容
    scene_context: 从Excel触发语构建的场景描述
    """
    # 接近最大轮次时提示C模型收尾
    hint = ""
    if round_num >= max_rounds - 2:
        hint = "\n\n提示：你们已经聊了很久了，请自然地结束这通电话。"

    system_prompt = persona_c + "\n\n" + scene_context + hint + "\n\n【重要】禁止使用任何形式的括号（包括（）和()），只输出说话内容本身。"

    all_messages = [{"role": "system", "content": system_prompt}] + messages

    last_error = None
    for attempt in range(MAX_RETRIES + 1):
        try:
            # 空回复重试时降低temperature，减少模型自由发挥
            temp = 0.5 if attempt > 0 else 0.9
            response = client.chat.completions.create(
                model=MODEL_C,
                messages=all_messages,
                temperature=temp,
                max_tokens=200,
            )
            reply = response.choices[0].message.content
            if reply is None or reply.strip() == '':
                continue  # 空回复，重试
            reply = reply.strip()
            cleaned = re.sub(r'[（(][^）)]*[）)]', '', reply).strip()
            if not cleaned:
                continue  # 空回复，重试
            return cleaned
        except Exception as e:
            last_error = e
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_DELAY * (attempt + 1))
    # 所有重试都返回空回复的兜底
    if last_error:
        raise last_error
    return "嗯，你说吧"

def load_persona(version_str):
    """从文件加载人设
    version_str: 'A_v1', 'B_v2' 等
    """
    filename = f"{version_str}.txt"
    filepath = os.path.join(PROMPTS_DIR, filename)
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"人设文件不存在: {filepath}")
    with open(filepath, 'r', encoding='utf-8') as f:
        return f.read().strip()


def get_latest_version(model_type):
    """获取指定模型的最新版本号
    model_type: 'A' 或 'B'
    返回: 'A_v1', 'B_v2' 等
    """
    pattern = re.compile(rf'^{model_type}_v(\d+)\.txt$')
    versions = []
    for f in os.listdir(PROMPTS_DIR):
        match = pattern.match(f)
        if match:
            versions.append(int(match.group(1)))
    if not versions:
        raise FileNotFoundError(f"未找到 {model_type}_*.txt 人设文件")
    return f"{model_type}_v{max(versions)}"


# ==================== 工具函数 ====================
def format_duration(seconds):
    """格式化时间显示：秒/分秒"""
    if seconds < 60:
        return f"{seconds:.1f}秒"
    minutes = int(seconds // 60)
    secs = seconds % 60
    return f"{minutes}分{secs:.0f}秒"


def get_column_index(headers, expected_name):
    """按表头名查找列索引，返回0-based索引。"""
    for idx, header in enumerate(headers):
        if header and str(header).strip() == expected_name:
            return idx
    raise ValueError(f"Excel缺少必要列: {expected_name}")


def load_scenarios(filepath):
    """读取Excel表格，按'用户反应分类'分组"""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active

    headers = [cell.value for cell in ws[1]]
    category_idx = get_column_index(headers, "用户反应分类")
    trigger_idx = get_column_index(headers, "用户可能说的话")

    scenarios = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        category = row[category_idx] if category_idx < len(row) else None
        trigger = row[trigger_idx] if trigger_idx < len(row) else None
        if not category or not trigger:
            continue
        category = str(category).strip()
        trigger = str(trigger).strip()
        if category not in scenarios:
            scenarios[category] = []
        scenarios[category].append(trigger)

    wb.close()
    return scenarios


def select_scenarios(scenarios, limit=0, sample_size=0, sample_seed=42):
    """按参数选择要执行的场景，保持默认行为不变。"""
    if sample_size > 0:
        items = list(scenarios.items())
        rng = random.Random(sample_seed)
        rng.shuffle(items)
        return dict(items[:sample_size])
    if limit > 0:
        keys = list(scenarios.keys())[:limit]
        return {k: scenarios[k] for k in keys}
    return scenarios


def call_model(client, model, messages, system_prompt, max_retries=None, temperature=0.7):
    """调用模型API，失败自动重试"""
    if max_retries is None:
        max_retries = MAX_RETRIES
    all_messages = [{"role": "system", "content": system_prompt}] + messages
    last_error = None
    for attempt in range(max_retries + 1):
        try:
            kwargs = dict(
                model=model,
                messages=all_messages,
                temperature=temperature,
            )
            # DeepSeek深度思考模式
            if model == MODEL_B:
                kwargs["temperature"] = 0.3
                kwargs["extra_body"] = {"thinking": {"type": "enabled"}}

            response = client.chat.completions.create(**kwargs)
            return response.choices[0].message.content
        except Exception as e:
            last_error = e
            if attempt < max_retries:
                time.sleep(RETRY_DELAY * (attempt + 1))
    raise last_error


def extract_json(text):
    """从B模型输出中提取JSON"""
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```\w*\n?", "", text)
        text = re.sub(r"\n?```$", "", text)
    match = re.search(r"\{[\s\S]*\}", text)
    if match:
        return match.group(0)
    return text


def format_dialogue_for_audit(dialogue_history):
    """将对话历史格式化为B模型可审计的文本"""
    lines = ["【通话记录】"]
    round_num = 1
    for entry in dialogue_history:
        role_label = "AI客服" if entry["role"] == "assistant" else "用户"
        lines.append(f"[第{round_num}轮]")
        lines.append(f"{role_label}：{entry['content']}")
        lines.append("")
        round_num += 1
    return "\n".join(lines)


def format_dialogue_display(dialogue_history):
    """将对话历史格式化为Excel展示用的文本"""
    lines = []
    for entry in dialogue_history:
        role_label = "【AI客服】" if entry["role"] == "assistant" else "【用户】"
        lines.append(f"{role_label} {entry['content']}")
    return "\n\n".join(lines)


def run_dialogue(client, category, triggers, persona_a, persona_c):
    """
    执行一通完整对话（C模型动态扮演用户）：
    1. 用户接听("喂？") → AI开场白
    2. C模型根据人设卡动态生成用户回复 → AI回复
    3. 循环直到：AI给出结束语 / C模型说再见 / 达到MAX_ROUNDS
    4. 兜底：如果AI还没结束，追加一句收尾
    """
    # 构建场景上下文
    scene_context = build_scene_context(category, triggers)

    # A模型的对话消息列表（含system prompt）
    a_messages = []
    # C模型的对话消息列表（不含system prompt，每次动态构建）
    c_messages = []
    dialogue_history = []
    round_num = 0

    # --- 第1轮：用户接听（固定） ---
    a_messages.append({"role": "user", "content": "喂？"})
    c_messages.append({"role": "assistant", "content": "喂？"})  # C模型扮演用户(assistant)
    try:
        ai_reply = call_model(client, MODEL_A, a_messages, persona_a, temperature=0.7)
    except Exception as e:
        return None, f"A模型第1轮(接听)调用失败: {e}"
    a_messages.append({"role": "assistant", "content": ai_reply})
    c_messages.append({"role": "user", "content": ai_reply})  # AI客服的话对用户来说是"user"
    dialogue_history.append({"role": "assistant", "content": ai_reply})
    round_num += 1

    # --- 动态对话循环 ---
    for i in range(MAX_ROUNDS):
        round_num += 1

        # 检测AI是否已经给出结束语
        if dialogue_history and is_conversation_ended(dialogue_history[-1]["content"]):
            break

        # C模型生成用户回复
        try:
            user_turn = call_user_model(client, persona_c, scene_context, c_messages, round_num, MAX_ROUNDS)
        except Exception as e:
            return None, f"C模型第{round_num}轮调用失败: {e}"

        # 检测用户是否说了结束语
        user_end_signals = ["再见", "挂了", "拜拜", "不聊了", "就这样吧", "不用了再见"]
        user_wants_end = any(s in user_turn for s in user_end_signals)

        # 将用户回复加入两个消息列表
        a_messages.append({"role": "user", "content": user_turn})
        c_messages.append({"role": "assistant", "content": user_turn})  # C模型扮演用户(assistant)
        dialogue_history.append({"role": "user", "content": user_turn})

        # 如果用户说了结束语，AI应该回复结束语然后结束
        if user_wants_end:
            try:
                ai_reply = call_model(client, MODEL_A, a_messages, persona_a, temperature=0.7)
            except Exception as e:
                return None, f"A模型第{round_num}轮(用户结束语)调用失败: {e}"
            a_messages.append({"role": "assistant", "content": ai_reply})
            dialogue_history.append({"role": "assistant", "content": ai_reply})
            break

        # AI回复
        try:
            ai_reply = call_model(client, MODEL_A, a_messages, persona_a, temperature=0.7)
        except Exception as e:
            return None, f"A模型第{round_num}轮调用失败: {e}"
        a_messages.append({"role": "assistant", "content": ai_reply})
        c_messages.append({"role": "user", "content": ai_reply})  # AI客服的话对用户来说是"user"
        dialogue_history.append({"role": "assistant", "content": ai_reply})

        # 检测AI是否给出了结束语
        if is_conversation_ended(ai_reply):
            break

    # --- 兜底：如果AI还没结束，再追问一轮 ---
    if dialogue_history and not is_conversation_ended(dialogue_history[-1]["content"]):
        fallback = "好的，那就这样吧，再见"
        a_messages.append({"role": "user", "content": fallback})
        dialogue_history.append({"role": "user", "content": fallback})
        try:
            ai_reply = call_model(client, MODEL_A, a_messages, persona_a, temperature=0.7)
        except Exception as e:
            return None, f"A模型兜底轮调用失败: {e}"
        a_messages.append({"role": "assistant", "content": ai_reply})
        dialogue_history.append({"role": "assistant", "content": ai_reply})

    return dialogue_history, None


def run_audit(client, dialogue_history, persona_b):
    """将对话记录发给B模型质检"""
    dialogue_text = format_dialogue_for_audit(dialogue_history)
    audit_messages = [
        {"role": "user", "content": f"请审计以下通话记录，严格按JSON格式输出结果。\n\n{dialogue_text}"},
    ]
    try:
        raw = call_model(client, MODEL_B, audit_messages, persona_b, temperature=0.3)
    except Exception as e:
        return False, "{}", f"B模型调用失败: {e}"

    try:
        json_str = extract_json(raw)
        result = json.loads(json_str)
        results_list = result.get("results", [])
        passed = len(results_list) == 0
        return passed, json.dumps(result, ensure_ascii=False, indent=2), None
    except json.JSONDecodeError:
        return False, raw, f"B模型返回非JSON格式"


def run_one_scenario(client, category, triggers, persona_a, persona_b, persona_c, skip_audit=False):
    """执行一个场景分类的完整测试"""
    result = {
        "category": category,
        "triggers": "；".join(triggers),
        "dialogue": "",
        "passed": False,
        "audit_json": "",
        "violation_count": 0,
        "violation_summary": "",
        "error": None,
    }

    dialogue_history, err = run_dialogue(client, category, triggers, persona_a, persona_c)
    if err:
        result["error"] = err
        result["dialogue"] = f"[对话生成失败] {err}"
        return result

    result["dialogue"] = format_dialogue_display(dialogue_history)

    if skip_audit:
        result["passed"] = None
        result["audit_json"] = '{"results": [], "skipped": true}'
        result["violation_summary"] = "已跳过B模型质检"
        return result

    passed, audit_json, err = run_audit(client, dialogue_history, persona_b)
    result["passed"] = passed
    result["audit_json"] = audit_json

    if err:
        result["error"] = err
        return result

    try:
        audit_data = json.loads(audit_json)
        violations = audit_data.get("results", [])
        result["violation_count"] = len(violations)
        if violations:
            summaries = []
            for v in violations:
                std = v.get("quality_standard", "未知")
                reason = v.get("reason", "")
                summaries.append(f"[{std}] {reason}")
            result["violation_summary"] = " | ".join(summaries)
    except json.JSONDecodeError:
        result["violation_count"] = -1
        result["violation_summary"] = "质检结果JSON解析失败"

    return result


def run_test(client, scenarios, persona_a, persona_b, persona_c, label="", skip_audit=False):
    """运行完整测试"""
    results = []
    total = len(scenarios)
    completed = 0
    start_time = time.time()

    if label:
        print(f"\n  [{label}] 开始测试...")

    with ThreadPoolExecutor(max_workers=CONCURRENCY) as executor:
        future_map = {}
        for category, triggers in scenarios.items():
            future = executor.submit(run_one_scenario, client, category, triggers, persona_a, persona_b, persona_c, skip_audit)
            future_map[future] = category

        for future in as_completed(future_map):
            category = future_map[future]
            completed += 1
            try:
                result = future.result()
                results.append(result)
                status = "跳过质检" if result["passed"] is None else ("✓ 合格" if result["passed"] else ("✗ 不合格" if not result["error"] else "⚠ 异常"))
                if label:
                    print(f"    [{label}][{completed}/{total}] {category}: {status} (违规{result['violation_count']}项)")
                else:
                    print(f"  [{completed}/{total}] {category}: {status} (违规{result['violation_count']}项)")
                if result["error"]:
                    print(f"         错误: {result['error'][:100]}")
            except Exception as e:
                if label:
                    print(f"    [{label}][{completed}/{total}] {category}: ⚠ 线程异常: {e}")
                else:
                    print(f"  [{completed}/{total}] {category}: ⚠ 线程异常: {e}")
                results.append({
                    "category": category,
                    "triggers": "；".join(scenarios.get(category, [])),
                    "dialogue": "",
                    "passed": False,
                    "audit_json": "",
                    "violation_count": 0,
                    "violation_summary": "",
                    "error": f"线程异常: {e}",
                })

    elapsed = time.time() - start_time
    return results, elapsed


def update_changelog(old_a, new_a, old_b, new_b, results_old, results_new):
    """更新变更日志"""
    if os.path.exists(CHANGELOG_FILE):
        with open(CHANGELOG_FILE, 'r', encoding='utf-8') as f:
            changelog = json.load(f)
    else:
        changelog = {"versions": []}

    # 计算变化
    old_passed = sum(1 for r in results_old if r["passed"] and not r["error"])
    new_passed = sum(1 for r in results_new if r["passed"] and not r["error"])
    change_summary = f"合格率: {old_passed}/{len(results_old)} → {new_passed}/{len(results_new)}"

    entry = {
        "timestamp": datetime.now().isoformat(),
        "old_persona": {"A": old_a, "B": old_b},
        "new_persona": {"A": new_a, "B": new_b},
        "change_summary": change_summary,
    }
    changelog["versions"].append(entry)

    with open(CHANGELOG_FILE, 'w', encoding='utf-8') as f:
        json.dump(changelog, f, ensure_ascii=False, indent=2)
    print(f"\n变更日志已更新: {CHANGELOG_FILE}")


def write_single_excel(results, output_file, persona_a_ver, persona_b_ver):
    """单版本测试：写入Excel"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"测试结果_{persona_a_ver}"

    headers = [
        "序号", "场景分类", "触发语列表", "对话全文", "质检结果",
        "是否合格", "违规项数", "违规摘要", "质检原始JSON", "错误信息"
    ]
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(name="微软雅黑", bold=True, color="FFFFFF", size=10)
    cell_font = Font(name="微软雅黑", size=9)
    wrap_align = Alignment(wrap_text=True, vertical="top")
    center_align = Alignment(horizontal="center", vertical="top")
    pass_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    fail_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border

    for i, r in enumerate(results):
        row = i + 2
        if r.get("passed") is None:
            result_text = "跳过质检"
            passed_text = "跳过质检"
        else:
            result_text = "合格" if r.get("passed") else ("不合格" if not r.get("error") else "异常")
            passed_text = "是" if r.get("passed") else ("否" if not r.get("error") else "异常")
        values = [
            i + 1,
            r.get("category", ""),
            r.get("triggers", ""),
            r.get("dialogue", ""),
            result_text,
            passed_text,
            r.get("violation_count", 0),
            r.get("violation_summary", ""),
            r.get("audit_json", ""),
            r.get("error", ""),
        ]
        for col, val in enumerate(values, 1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.font = cell_font
            cell.alignment = wrap_align
            cell.border = thin_border

        result_cell = ws.cell(row=row, column=5)
        if r.get("passed") is True:
            result_cell.fill = pass_fill
        elif r.get("passed") is False and not r.get("error"):
            result_cell.fill = fail_fill

    col_widths = [6, 14, 30, 60, 10, 10, 10, 40, 40, 30]
    for col, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = width

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    wb.save(output_file)
    print(f"\n结果已保存到: {output_file}")


def write_regression_excel(results_old, results_new, output_file, old_a_ver, new_a_ver, old_b_ver, new_b_ver):
    """回归对比：并排输出Excel"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "回归对比"

    headers = [
        "序号", "场景分类", "触发语列表",
        f"旧版{old_a_ver}对话", f"旧版{old_a_ver}质检结果", f"旧版{old_a_ver}违规摘要",
        f"新版{new_a_ver}对话", f"新版{new_a_ver}质检结果", f"新版{new_a_ver}违规摘要",
        "变化", "详情"
    ]
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(name="微软雅黑", bold=True, color="FFFFFF", size=10)
    cell_font = Font(name="微软雅黑", size=9)
    wrap_align = Alignment(wrap_text=True, vertical="top")
    center_align = Alignment(horizontal="center", vertical="top")
    pass_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    fail_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    improve_fill = PatternFill(start_color="B4C7E7", end_color="B4C7E7", fill_type="solid")
    regress_fill = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border

    # 按场景分类对齐
    old_map = {r["category"]: r for r in results_old}
    new_map = {r["category"]: r for r in results_new}
    all_categories = list(dict.fromkeys([r["category"] for r in results_old + results_new]))

    for i, cat in enumerate(all_categories):
        row = i + 2
        old_r = old_map.get(cat, {})
        new_r = new_map.get(cat, {})

        old_passed = old_r.get("passed", False)
        new_passed = new_r.get("passed", False)

        if old_passed is None or new_passed is None:
            change = "跳过质检"
            change_fill = None
        elif old_passed and new_passed:
            change = "-"
            change_fill = None
        elif not old_passed and not new_passed:
            change = "仍不合格"
            change_fill = fail_fill
        elif not old_passed and new_passed:
            change = "✓ 修复"
            change_fill = improve_fill
        else:
            change = "✗ 新增违规"
            change_fill = regress_fill

        detail = ""
        if change in ["✓ 修复", "✗ 新增违规"]:
            old_summary = old_r.get("violation_summary", "")
            new_summary = new_r.get("violation_summary", "")
            if old_summary or new_summary:
                detail = f"旧版: {old_summary}\n新版: {new_summary}"

        values = [
            i + 1,
            cat,
            old_r.get("triggers", new_r.get("triggers", "")),
            old_r.get("dialogue", ""),
            "跳过质检" if old_passed is None else ("合格" if old_passed else ("不合格" if not old_r.get("error") else "异常")),
            old_r.get("violation_summary", ""),
            new_r.get("dialogue", ""),
            "跳过质检" if new_passed is None else ("合格" if new_passed else ("不合格" if not new_r.get("error") else "异常")),
            new_r.get("violation_summary", ""),
            change,
            detail,
        ]

        for col, val in enumerate(values, 1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.font = cell_font
            cell.alignment = wrap_align
            cell.border = thin_border

        # 质检结果列着色
        old_result_cell = ws.cell(row=row, column=5)
        new_result_cell = ws.cell(row=row, column=8)
        if old_passed is True:
            old_result_cell.fill = pass_fill
        elif old_passed is False and not old_r.get("error"):
            old_result_cell.fill = fail_fill
        if new_passed is True:
            new_result_cell.fill = pass_fill
        elif new_passed is False and not new_r.get("error"):
            new_result_cell.fill = fail_fill

        # 变化列着色
        change_cell = ws.cell(row=row, column=10)
        if change_fill:
            change_cell.fill = change_fill

    col_widths = [6, 14, 30, 50, 10, 40, 50, 10, 40, 12, 60]
    for col, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = width

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    wb.save(output_file)
    print(f"\n回归对比结果已保存到: {output_file}")


def main():
    global INPUT_FILE, CONCURRENCY, MAX_ROUNDS, MAX_RETRIES

    parser = argparse.ArgumentParser(description="AI外呼用户反应测试脚本")
    parser.add_argument("--persona", nargs=2, metavar=("A_VER", "B_VER"),
                        help="指定人设版本，如: --persona A_v2 B_v1")
    parser.add_argument("--regression", action="store_true",
                        help="启用回归对比模式")
    parser.add_argument("--old", nargs=2, metavar=("A_VER", "B_VER"),
                        help="回归对比的旧版本，如: --old A_v1 B_v1")
    parser.add_argument("--new", nargs=2, metavar=("A_VER", "B_VER"),
                        help="回归对比的新版本，如: --new A_v2 B_v1")
    parser.add_argument("--persona-c", default="C_v1",
                        help="C模型人设版本，如: C_v1(普通) C_v2(投诉) C_v3(法律) C_v4(诱导) C_v5(情绪)")
    parser.add_argument("--limit", type=int, default=0,
                        help="限制测试场景数量，0=全部")
    parser.add_argument("--sample-size", type=int, default=0,
                        help="随机抽样测试的场景分类数量，0=不抽样")
    parser.add_argument("--sample-seed", type=int, default=42,
                        help="随机抽样种子，便于复现")
    parser.add_argument("--concurrency", type=int, default=CONCURRENCY,
                        help=f"并发场景数，默认读取TEST_CONCURRENCY或{CONCURRENCY}")
    parser.add_argument("--max-rounds", type=int, default=MAX_ROUNDS,
                        help=f"单通对话最大轮次，默认读取MAX_ROUNDS或{MAX_ROUNDS}")
    parser.add_argument("--max-retries", type=int, default=MAX_RETRIES,
                        help=f"模型调用失败最大重试次数，默认读取MAX_RETRIES或{MAX_RETRIES}")
    parser.add_argument("--skip-audit", action="store_true",
                        help="只生成A/C对话，不调用B模型质检，用于快速冒烟测试")
    parser.add_argument("--input", default=INPUT_FILE,
                        help="输入Excel路径")

    args = parser.parse_args()
    INPUT_FILE = os.path.abspath(args.input)
    CONCURRENCY = args.concurrency
    MAX_ROUNDS = args.max_rounds
    MAX_RETRIES = args.max_retries

    if not API_KEY:
        print("  [错误] 未设置ARK_API_KEY环境变量")
        sys.exit(1)

    print("=" * 60)
    print("  AI外呼用户反应测试 v3.0")
    print(f"  A模型(AI客服): {MODEL_A}")
    print(f"  B模型(质检):   {MODEL_B}")
    print(f"  C模型(用户):   {MODEL_C}")
    print(f"  并发数: {CONCURRENCY}  最大轮次: {MAX_ROUNDS}")
    print("=" * 60)

    # 加载场景
    print("\n[1/3] 加载场景表格...")
    if not os.path.exists(INPUT_FILE):
        print(f"  [错误] 找不到输入文件: {INPUT_FILE}")
        sys.exit(1)

    scenarios = load_scenarios(INPUT_FILE)
    print(f"  共加载 {sum(len(v) for v in scenarios.values())} 个场景，{len(scenarios)} 个分类")

    scenarios = select_scenarios(scenarios, args.limit, args.sample_size, args.sample_seed)
    if args.sample_size > 0:
        print(f"  随机抽样 {len(scenarios)} 个分类，seed={args.sample_seed}")
    elif args.limit > 0:
        print(f"  限制测试前 {len(scenarios)} 个分类")

    try:
        from openai import OpenAI
    except ImportError:
        print("  [错误] 未安装openai依赖，请先执行: pip install -r requirements.txt")
        sys.exit(1)

    client = OpenAI(base_url=BASE_URL, api_key=API_KEY)

    if args.regression:
        # 回归对比模式
        if not args.old or not args.new:
            print("  [错误] 回归模式需要 --old 和 --new 参数")
            sys.exit(1)

        old_a_ver, old_b_ver = args.old
        new_a_ver, new_b_ver = args.new

        print(f"\n[2/3] 回归对比模式: {old_a_ver}/{old_b_ver} vs {new_a_ver}/{new_b_ver}, C={args.persona_c}")

        # 加载人设
        try:
            persona_a_old = load_persona(old_a_ver)
            persona_b_old = load_persona(old_b_ver)
            persona_a_new = load_persona(new_a_ver)
            persona_b_new = load_persona(new_b_ver)
            persona_c = load_persona(args.persona_c)
        except FileNotFoundError as e:
            print(f"  [错误] {e}")
            sys.exit(1)

        # 并发跑两个版本
        print("\n  并发执行两个版本测试...")
        total_start = time.time()
        with ThreadPoolExecutor(max_workers=2) as executor:
            future_old = executor.submit(run_test, client, scenarios, persona_a_old, persona_b_old, persona_c, "旧版", args.skip_audit)
            future_new = executor.submit(run_test, client, scenarios, persona_a_new, persona_b_new, persona_c, "新版", args.skip_audit)
            results_old, elapsed_old = future_old.result()
            results_new, elapsed_new = future_new.result()
        total_elapsed = time.time() - total_start

        # 按场景排序
        cat_order = list(scenarios.keys())
        results_old.sort(key=lambda r: cat_order.index(r["category"]) if r["category"] in cat_order else 999)
        results_new.sort(key=lambda r: cat_order.index(r["category"]) if r["category"] in cat_order else 999)

        # 写Excel
        print(f"\n[3/3] 写入回归对比Excel...")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(RESULTS_DIR, f"回归对比_{old_a_ver}_vs_{new_a_ver}_{timestamp}.xlsx")
        write_regression_excel(results_old, results_new, output_file, old_a_ver, new_a_ver, old_b_ver, new_b_ver)

        # 更新变更日志
        update_changelog(old_a_ver, new_a_ver, old_b_ver, new_b_ver, results_old, results_new)

        # 汇总
        old_passed = sum(1 for r in results_old if r["passed"] and not r["error"])
        new_passed = sum(1 for r in results_new if r["passed"] and not r["error"])
        fixed = sum(1 for r_old, r_new in zip(results_old, results_new) if not r_old["passed"] and r_new["passed"])
        regressed = sum(1 for r_old, r_new in zip(results_old, results_new) if r_old["passed"] and not r_new["passed"])

        print("\n" + "=" * 60)
        print(f"  回归对比完成")
        print(f"  旧版 ({old_a_ver}/{old_b_ver}): {old_passed}/{len(results_old)} 合格  耗时: {format_duration(elapsed_old)}")
        print(f"  新版 ({new_a_ver}/{new_b_ver}): {new_passed}/{len(results_new)} 合格  耗时: {format_duration(elapsed_new)}")
        print(f"  修复: {fixed} 项")
        print(f"  新增违规: {regressed} 项")
        print(f"  总耗时: {format_duration(total_elapsed)}")
        print("=" * 60)

    else:
        # 单版本测试模式
        if args.persona:
            a_ver, b_ver = args.persona
        else:
            a_ver = get_latest_version("A")
            b_ver = get_latest_version("B")

        print(f"\n[2/3] 单版本测试: A={a_ver}, B={b_ver}, C={args.persona_c}")

        try:
            persona_a = load_persona(a_ver)
            persona_b = load_persona(b_ver)
            persona_c = load_persona(args.persona_c)
        except FileNotFoundError as e:
            print(f"  [错误] {e}")
            sys.exit(1)

        results, elapsed = run_test(client, scenarios, persona_a, persona_b, persona_c, skip_audit=args.skip_audit)

        cat_order = list(scenarios.keys())
        results.sort(key=lambda r: cat_order.index(r["category"]) if r["category"] in cat_order else 999)

        print(f"\n[3/3] 写入Excel...")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(RESULTS_DIR, f"AI外呼测试结果_{a_ver}_{timestamp}.xlsx")
        write_single_excel(results, output_file, a_ver, b_ver)

        total_scenarios = len(results)
        passed_count = sum(1 for r in results if r["passed"] and not r["error"])
        skipped_count = sum(1 for r in results if r["passed"] is None and not r["error"])
        failed_count = sum(1 for r in results if r["passed"] is False and not r["error"])
        error_count = sum(1 for r in results if r["error"])

        print("\n" + "=" * 60)
        print(f"  测试完成: 共{total_scenarios}个分类")
        print(f"  合格: {passed_count}  不合格: {failed_count}  跳过质检: {skipped_count}  异常: {error_count}")
        print(f"  总耗时: {format_duration(elapsed)}")
        print(f"  平均每个场景: {format_duration(elapsed / total_scenarios)}")
        print("=" * 60)


if __name__ == "__main__":
    main()
