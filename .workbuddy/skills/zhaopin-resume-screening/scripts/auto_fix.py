"""
自验证 + 自动修复模块。

验证规则：每段工作经历必须是 【岗位名称】公司名称-时间 格式。
不符合则尝试自动修复：从原始 PDF 文本中重新提取岗位名，或移除误判条目。
"""

import json, re, sys
from pathlib import Path

# 编译正则
COMPANY_SUFFIX_RE = re.compile(
    r'(有限公司|有限责任公司|股份有限公司|集团公司|总公司|分公司|支公司|'
    r'研究院|研究所|学校|学院|旅行社|工厂|分局|办事处|事务所|实验室)'
)
TIME_RE = re.compile(r'(\d{4}\.\d{1,2})\s*-\s*(至今|\d{4}\.\d{1,2})')

# ============= 误判条目检测（应删除） =============

FALSE_POSITIVE_PATTERNS = [
    # 证书/资格证
    r'(从业资格证|资格证书|结业证|驾驶证|C1)',
    # 培训机构
    r'培训机构[:：]',
    # 纯项目名
    r'(存证系统|智慧空压站|流通底座|数据底座|采购项目|交易中心)',
    # 行业标签（单独一行无公司后缀）
    r'^【】\s*(新能源|工业工程|金融|医疗|教育|房地产|互联网|文案|营销|翻译)\s*(-|$)',
    # "XX分公司建立" 等描述
    r'(分公司建立|分公司组建|分公司筹备)',
    # 证券/保险/基金相关非工作条目
    r'(证券从业|保险从业|基金从业)',
    # 纯描述行开头
    r'^(分别与|，|负责项目|项目描述[：:])',
    # 简历标记
    r'^(教育经|自我评价|语言能力|所获证书|项目经历)',
    # 商业模式/创新项目等描述
    r'(商业模式创新|全球商业模式|创新项目|解决方案)',
    # 纯括号内短片段
    r'^【】\s*\([^)]+\)\s*-',
    # 描述碎片
    r'^【】\s*制造',
    # 人才学院/培训
    r'(人才学院|培训学校|培训中心)',
    # 项目名（含项目）
    r'(项目$|项目描述|项目经历|采购项目|分析报告|管理流程)',
    # 工作成果/HR模板
    r'(工作成果示例|核心能力|团队建设|业务拓展|战略规划)',
    # 省/市/区级项目/中心
    r'(公共资源交易|区块链系统建设|产业园区建设|校企合作)',
    # 技能标签
    r'^(英语|读写|听说|精通|熟练|驾驶)',
]


def _is_false_positive(entry):
    """判断是否为误判条目（项目名、证书、描述片段等）"""
    for pat in FALSE_POSITIVE_PATTERNS:
        if re.search(pat, entry):
            return True
    
    # 额外检查公司名部分
    m = re.match(r'^【[^】]*】\s*(.+)$', entry)
    if m:
        company_part = m.group(1)
        # 公司名以 "分别与" 或 "," 开头 = 客户列表
        if re.match(r'^(分别与|，|,)', company_part):
            return True
        # 纯纯技能词/行业词做公司名
        if re.match(r'^[\u4e00-\u9fa5]{2,4}$', company_part) and not COMPANY_SUFFIX_RE.search(company_part):
            # 极短的纯中文，可能不是公司名
            pass
    
    return False


# ============= 岗位+公司拆分合并 =============

def _merge_split_entries(entries):
    """合并被拆成两行的 岗位+公司 条目。

    支持两种顺序：
    (1) 【】公司X + 【岗位A】[时间] → 【岗位A】公司X[时间]
    (2) 【岗位A】[时间] + 【】公司X → 【岗位A】公司X[时间]
    """
    if len(entries) < 2:
        return entries

    merged = []
    i = 0
    while i < len(entries):
        w = entries[i].strip()
        if i + 1 >= len(entries):
            merged.append(w)
            break

        m1 = re.match(r'^【([^】]*)】\s*(.+)$', w)
        nw = entries[i + 1].strip()
        m2 = re.match(r'^【([^】]*)】\s*(.+)$', nw)

        if not m1 or not m2:
            merged.append(w)
            i += 1
            continue

        p1, r1 = m1.group(1), m1.group(2)
        p2, r2 = m2.group(1), m2.group(2)

        # 情况1: 空岗位公司 + 有岗位无公司
        if not p1 and COMPANY_SUFFIX_RE.search(r1) and p2 and not COMPANY_SUFFIX_RE.search(r2):
            t2 = TIME_RE.search(r2)
            time_str = r2[t2.start():].strip() if t2 else ''
            merged.append(f"【{p2}】{r1}" + (f"-{time_str}" if time_str else ""))
            i += 2
            continue

        # 情况2: 有岗位无公司 + 空岗位公司
        if p1 and not COMPANY_SUFFIX_RE.search(r1) and not p2 and COMPANY_SUFFIX_RE.search(r2):
            t1 = TIME_RE.search(r1)
            time_str = r1[t1.start():].strip() if t1 else ''
            merged.append(f"【{p1}】{r2}" + (f"-{time_str}" if time_str else ""))
            i += 2
            continue

        merged.append(w)
        i += 1

    return merged


# ============= 去重 =============

def _deduplicate_entries(entries):
    """移除近似重复的条目（同一公司被检测两次）。

    如：【施工安全员】中建八局土木 + 【】中建八局土木公司 → 保留第一个
    """
    if len(entries) < 2:
        return entries

    result = []
    for i, w in enumerate(entries):
        m = re.match(r'^【([^】]*)】\s*(.+)$', w)
        if not m:
            result.append(w)
            continue

        pos, rest = m.group(1), m.group(2)
        tm = TIME_RE.search(rest)
        company = rest[:tm.start()].strip() if tm else rest.strip()

        # 提取核心公司名（去掉后缀用于比较）
        core = COMPANY_SUFFIX_RE.sub('', company).strip()
        # Also strip standalone "公司"/"集团" and trailing punctuation
        core = re.sub(r'(公司|集团)\s*$', '', core).strip()
        core = re.sub(r'[-.]\s*$', '', core).strip()
        if not core:
            core = company

        is_dup = False
        for prev_w in result:
            prev_m = re.match(r'^【([^】]*)】\s*(.+)$', prev_w)
            if not prev_m:
                continue
            prev_rest = prev_m.group(2)
            prev_tm = TIME_RE.search(prev_rest)
            prev_co = prev_rest[:prev_tm.start()].strip() if prev_tm else prev_rest.strip()
            prev_core = COMPANY_SUFFIX_RE.sub('', prev_co).strip()
            prev_core = re.sub(r'(公司|集团)\s*$', '', prev_core).strip()
            prev_core = re.sub(r'[-.]\s*$', '', prev_core).strip()
            if not prev_core:
                prev_core = prev_co

            # 核心公司名模糊匹配
            if core and prev_core and (
                core == prev_core or
                prev_core in core or
                core in prev_core or
                (len(core) >= 3 and len(prev_core) >= 3 and
                 (core[:min(len(core), len(prev_core))] == prev_core[:min(len(core), len(prev_core))]))
            ):
                is_dup = True
                break

        if not is_dup:
            result.append(w)

    return result


# ============= 岗位名重新提取 =============

POSITION_KW_RE = re.compile(
    r'(经理|总监|代表|主管|工程师|顾问|专员|主任|总裁|助理|负责人|VP|'
    r'店长|组长|资料员|安全员|施工员|操作工|培训师|讲师|组训|'
    r'秘书|文员|销售|招商|运营|BD|售后|售前|技术支持|'
    r'客服|会计|出纳|采购|仓储|物流|跟单|报关|外贸|'
    r'文案|策划|编辑|市场|产品|项目|研发|设计|开发|测试|运维|'
    r'服务员|厨师|司机|保安|保洁|主播|推广|新媒体)'
)


def _extract_position_from_raw(raw_text, company_name):
    """尝试从原始文本中重新提取某公司的岗位名。

    搜索策略：模糊匹配公司名 → 在后续行内找岗位行。
    """
    if not raw_text or not company_name:
        return None

    lines = raw_text.split('\n')

    # 提取公司名的核心部分（去掉常见的追加/后缀变化）
    # 如 "中建八局土木公司" → 搜索 "中建八局土木"
    search_names = [company_name]
    for suffix in ['公司', '有限公司', '分公司', '总公司', '集团']:
        if company_name.endswith(suffix):
            search_names.append(company_name[:-len(suffix)])
    # 如果公司名很长，也尝试前8字
    if len(company_name) > 10:
        search_names.append(company_name[:8])

    found_idx = -1
    for name_variant in search_names:
        for line_idx, line in enumerate(lines):
            if name_variant in line.strip():
                found_idx = line_idx
                break
        if found_idx >= 0:
            break

    if found_idx < 0:
        return None

    # 在公司行后搜索岗位
    for j in range(found_idx + 1, min(found_idx + 12, len(lines))):
        nl = lines[j].strip()
        if not nl or len(nl) < 2 or len(nl) > 25:
            continue
        # 跳过纯数字/符号/薪资行
        if re.match(r'^[\d\s,.，。/、万元月千\(\)（）\-/]+$', nl):
            continue
        # 跳过描述标记
        if re.match(r'^(工作描述|内容[：:]|业绩[：:]|教育经)', nl):
            break
        # 跳过城市名
        if re.match(r'^(青岛|济南|北京|上海|广州|深圳|杭州|成都|武汉|南京|天津|重庆|西安|'
                   r'苏州|郑州|长沙|合肥|大连|沈阳|长春|哈尔滨|昆明|贵阳|滨州|淄博)$', nl):
            continue
        # 岗位关键词匹配
        if POSITION_KW_RE.search(nl) and re.search(r'[\u4e00-\u9fa5]', nl):
            return nl.replace(' ', '')
        # 兜底：短中文行（2-6字），可能是岗名简写
        if 2 <= len(nl) <= 6 and re.match(r'^[\u4e00-\u9fa5]+$', nl):
            return nl

    return None


def validate_and_fix(extracted_data_path, raw_texts_dir, output_path=None, max_rounds=3):
    """验证并自动修复提取结果。重复循环直到无误或达到最大轮数。

    Args:
        extracted_data_path: JSON 提取结果路径
        raw_texts_dir: PDF 原始文本目录（用于重新提取岗位名）
        output_path: 修复后输出的 JSON 路径（默认覆盖原文件）
        max_rounds: 最大修复轮数
    """
    with open(extracted_data_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # 读取原始文本（用于岗位名重新提取）
    raw_texts = {}
    if Path(raw_texts_dir).exists():
        for r in data:
            txt_path = Path(raw_texts_dir) / f"{Path(r['filename']).stem}.txt"
            if txt_path.exists():
                with open(txt_path, 'r', encoding='utf-8') as f:
                    raw_texts[r['filename']] = f.read()

    for round_num in range(1, max_rounds + 1):
        total_fixes = 0
        total_removed = 0

        for r in data:
            ws = r['work_experiences']
            if not ws:
                continue

            wl = [w for w in ws.split('\n') if w.strip()]

            # 先合并拆分条目，再去重
            wl = _merge_split_entries(wl)
            wl = _deduplicate_entries(wl)

            fixed_entries = []
            raw = raw_texts.get(r['filename'], '')

            for w in wl:
                m = re.match(r'^【([^】]*)】\s*(.+)$', w)
                if not m:
                    fixed_entries.append(w)
                    continue

                position, rest = m.group(1), m.group(2)

                # 分离公司名和时间
                tm = TIME_RE.search(rest)
                company = rest[:tm.start()].strip() if tm else rest.strip()
                time_str = rest[tm.start():].strip() if tm else ''

                # ---- 检查0：空岗位 + 有时间的条目，如果"rest"文本已在其他条目中作为岗位名 → 删除 ----
                if not position and time_str:
                    # 收集已修复条目的所有岗位名
                    existing_positions = set()
                    for prev_w in fixed_entries:
                        prev_m = re.match(r'^【([^】]*)】\s*(.+)$', prev_w)
                        if prev_m and prev_m.group(1):
                            existing_positions.add(prev_m.group(1))
                    co_clean = re.sub(r'[-.]\s*$', '', company).strip()
                    if co_clean in existing_positions:
                        total_removed += 1
                        continue

                # ---- 检查1：误判条目 → 删除 ----
                if _is_false_positive(w):
                    total_removed += 1
                    continue

                # ---- 检查2：空岗位 + 有效公司 → 重新提取 ----
                if not position and company and COMPANY_SUFFIX_RE.search(company):
                    new_pos = _extract_position_from_raw(raw, company)
                    if new_pos:
                        position = new_pos
                        w = f"【{position}】{company}"
                        if time_str:
                            w += f"-{time_str}"
                        total_fixes += 1

                # ---- 检查3：岗位名过长 → 尝试缩短 ----
                if len(position) > 40:
                    # 可能是描述拼接，取前15字
                    short_pos = position[:15]
                    w = f"【{short_pos}】{company}"
                    if time_str:
                        w += f"-{time_str}"
                    total_fixes += 1

                fixed_entries.append(w)

            r['work_experiences'] = '\n'.join(fixed_entries) if fixed_entries else ""

        print(f"  Round {round_num}: fixed {total_fixes} positions, removed {total_removed} false positives")

        # 清理纯空行
        for r in data:
            if r['work_experiences']:
                r['work_experiences'] = '\n'.join(
                    [w for w in r['work_experiences'].split('\n') if w.strip()]
                )

        if total_fixes == 0 and total_removed == 0:
            print(f"  All clean after round {round_num}!")
            break

    # 输出
    out_path = output_path or extracted_data_path
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    # 最终统计
    total_issues = 0
    for r in data:
        ws = r['work_experiences']
        if not ws:
            continue
        for w in ws.split('\n'):
            w = w.strip()
            if not w:
                continue
            m = re.match(r'^【([^】]*)】\s*(.+)$', w)
            if not m:
                total_issues += 1
                continue
            pos = m.group(1)
            if not pos:
                total_issues += 1

    print(f"  Final issues remaining: {total_issues}")
    return data, total_issues


if __name__ == "__main__":
    import sys
    json_path = sys.argv[1] if len(sys.argv) > 1 else "extracted_data.json"
    raw_dir = sys.argv[2] if len(sys.argv) > 2 else ""
    output = sys.argv[3] if len(sys.argv) > 3 else json_path

    data, issues = validate_and_fix(json_path, raw_dir, output)
    print(f"Done. {issues} issues remain.")
