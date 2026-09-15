"""
智联招聘简历 PDF 解析模块 v3。

核心解析策略：每段工作经历 = 公司 → 岗位 → 时间 → 工作描述（顺序四部分）

Usage:
    from extract_resumes import batch_extract
    results = batch_extract("/path/to/pdf/folder")
"""

import re
import json
from datetime import datetime
from pathlib import Path
from pypdf import PdfReader


# ============= 基础工具函数 =============

def extract_text_from_pdf(pdf_path):
    reader = PdfReader(str(pdf_path))
    text = ""
    for page in reader.pages:
        t = page.extract_text()
        if t:
            text += t + "\n"
    return text


def fix_fragments(text):
    """Fix PDF text fragments split across lines."""
    t = text

    # Multi-line salary ranges
    for _ in range(3):
        t = re.sub(
            r'([\d,.]+)\s*\n\s*(万|千|百)\s*\n\s*-\s*([\d,.]+)\s*\n\s*(万|千|百)\s*\n\s*/\s*\n\s*(月|年)',
            r'\1\2-\3\4/\5', t
        )

    # Number + unit merges
    for _ in range(2):
        t = re.sub(r'([\d,.]+)\s*\n\s*(万|千|百|元)\s*\n\s*/\s*\n\s*(月|年)', r'\1\2/\3', t)
        t = re.sub(r'([\d,.]+)\s*\n\s*(万|千|百|元)', r'\1\2', t)

    # /月 /年 merges
    t = re.sub(r'/\s*\n\s*(月|年)', r'/\1', t)

    # Time ranges
    t = re.sub(r'-\s*\n\s*(\d{4}\.\d{1,2})', r'-\1', t)
    t = re.sub(r'-\s*\n\s*(至今)', r'-\1', t)

    # Duration with parens
    t = re.sub(r'\(\s*(\d+)\s*\n\s*年\s*\n\s*(\d+)\s*\n\s*个?\s*月?\s*\n\s*\)', r'(\1年\2个月)', t)
    t = re.sub(r'\(\s*(\d+)\s*\n\s*年\s*\n\s*\)', r'(\1年)', t)
    t = re.sub(r'\(\s*(\d+)\s*\n\s*个?\s*月?\s*\n\s*\)', r'(\1个月)', t)

    # Simple duration
    t = re.sub(r'(\d+)\s*\n\s*年\s*\n\s*(\d+)\s*\n\s*个?\s*月?', r'\1年\2个月', t)
    t = re.sub(r'(\d+)\s*\n\s*年', r'\1年', t)
    t = re.sub(r'(\d+)\s*\n\s*个?\s*月?', r'\1个月', t)

    # Company name with parens — unified pattern handling ALL intermediate words
    # 海创链\n(\n青岛\n)\n信息科技有限公司 → 海创链(青岛)信息科技有限公司
    # 百威\n(\n中国\n)\n销售有限公司 → 百威(中国)销售有限公司
    for _ in range(3):
        t = re.sub(
            r'([\u4e00-\u9fa5a-zA-Z\d]+)\s*\n\s*\(\s*\n\s*([\u4e00-\u9fa5a-zA-Z]+)\s*\n\s*\)'
            r'\s*\n\s*(.+?(?:有限公司|有限责任公司|股份有限公司|集团|公司|分局|办事处|研究院))',
            r'\1(\2)\3', t
        )

    # 不限行业
    t = re.sub(r'不\s*\n?\s*限\s*\n?\s*行\s*\n?\s*业', '不限行业', t)

    # Position join across lines
    t = re.sub(r'/\s*\n\s*', '/', t)
    t = re.sub(r'\n\s*/\s*\n', '/', t)

    # Standalone parens: (\n销售团队经理\n) → (销售团队经理)
    t = re.sub(r'\(\s*\n\s*([\u4e00-\u9fa5a-zA-Z/、]+)\s*\n\s*\)', r'(\1)', t)

    # Fix space-separated role fragments
    for prefix, suffix in [
        ('副总经', '理'), ('总经理', '助理'), ('项目经', '理'),
        ('城市经', '理'), ('运营经', '理'), ('销售总', '监'),
        ('市场总', '监'), ('咨', '询服务'), ('总', '助'),
        ('大', '客户'), ('区域总', '监'), ('产品经', '理'),
        ('工程经', '理'), ('技术总', '监'), ('销售团', '队经理'),
        ('企业管理咨', '询'),  # 企业管理咨询 合并
    ]:
        t = re.sub(rf'({prefix})\s+({suffix})', r'\1\2', t)

    return t


# ============= 公司识别 =============

COMPANY_SUFFIX = re.compile(
    r'(有限(?:责任)?公司|股份有限(?:责任)?公司|集团(?:公司)?|总公司|分公司|支公司|'
    r'研究院|研究所|学校|学院|旅行社|工厂|分局|办事处|事务所|实验室)\s*$'
)

# 教育机构后缀行（学校、大学、学院），后续有学历关键词则为教育经历
EDUCATION_SUFFIX = re.compile(r'(大学|学院|学校|职业技术学院|研究院)\s*$')
EDUCATION_KEYWORDS = re.compile(r'(统招|非统招|本科|大专|硕士|博士|学历|学士)')

# 公司识别（包含后缀的行）—— 排除教育机构
def _is_company_line(ls):
    """Check if a line looks like a company name (has standard suffix)."""
    if not COMPANY_SUFFIX.search(ls):
        return False
    return True


def _is_education_company(ls, lines, idx):
    """公司后缀行是否其实是教育经历（如XX大学）"""
    if not EDUCATION_SUFFIX.search(ls):
        return False
    # 检查后续行是否有学历关键词
    for j in range(idx + 1, min(idx + 5, len(lines))):
        al = lines[j].strip()
        if EDUCATION_KEYWORDS.search(al):
            return True
        if re.match(r'^\d{4}\.\d{1,2}\s*-\s*\d{4}\.\d{1,2}\s*$', al.replace(' ', '')):
            return True
    return False


# ============= 岗位识别 =============

POSITION_KW = re.compile(
    r'(经理|总监|代表|主管|工程师|顾问|专员|主任|总裁|助理|负责人|VP|店长|组长|'
    r'资料员|安全员|施工员|培训师|讲师|组训|秘书|文员|'
    r'销售|招商|BD|售后|售前|技术支持|客服|会计|出纳|'
    r'采购|仓储|物流|跟单|报关|外贸|文案|策划|编辑|咨询|'
    r'市场|产品|项目|研发|设计|开发|测试|运维|'
    r'服务员|厨师|司机|保安|保洁|主播|推广|新媒体)'
)


# 强岗位关键词（极少出现在公司名中），用于区分公司名和岗位名
STRONG_POSITION_KW = re.compile(
    r'(经理|总监|主管|代表|专员|主任|总裁|助理|负责人|VP|'
    r'店长|组长|工程师|顾问|培训师|讲师|组训|'
    r'资料员|安全员|施工员|秘书|文员|'
    r'服务员|厨师|司机|保安|保洁|主播)'
)

# 行业特征词常出现在公司名末尾（酒业/科技/商贸等），不含强岗位词时判定为公司名
INDUSTRY_SUFFIX = re.compile(
    r'(酒业|商贸|科技|实业|工贸|制造|电子|信息|智能|'
    r'网络|软件|硬件|能源|建设|地产|物业|物流|医药|环保|'
    r'通信|光电|食品|饮品|服装|纺织|保险|银行|证券|基金|'
    r'传媒|文化|旅游|装饰|家具|家居|厨卫|电力|水务|燃气|'
    r'热力|材料|汽车|设备|仪器|电器|电气|机械|五金|建材|'
    r'咨询|数据|生物|工程|化工|控股|股份)$'
)

def _is_likely_company(ls):
    """判断是否为高概率公司名（含行业后缀且无强岗位关键词）。"""
    if not ls or len(ls) > 20:
        return False
    if COMPANY_SUFFIX.search(ls):
        return True
    if INDUSTRY_SUFFIX.search(ls) and not STRONG_POSITION_KW.search(ls):
        return True
    return False


def _is_position_line(ls):
    """判断是否是岗位行（含职位关键词，长度 <= 25）"""
    if not ls or len(ls) > 25:
        return False
    if not re.search(r'[\u4e00-\u9fa5]', ls):
        return False
    # 不能是纯描述文本（含中文标点的长句）
    if len(ls) > 12 and re.search(r'[，。！？：；、]', ls):
        return False
    # 含公司后缀/行业后缀的行优先判定为公司名，非岗位行
    # 反例：百威(中国)销售有限公司 — 含'销售'在POSITION_KW但实为公司名
    # 反例：青岛拉图拉甘酒业 — 含'销售'但'酒业'是行业后缀
    if _is_likely_company(ls):
        return False
    return bool(POSITION_KW.search(ls))


# ============= 时间识别 =============

TIME_RANGE_RE = re.compile(r'(\d{4}\.\d{1,2})\s*-\s*(至今|\d{4}\.\d{1,2})')
DURATION_RE = re.compile(r'(\d+年\d+个月|\d+个月|\d+年)')


# ============= 描述标记 =============

DESCRIPTION_MARKER = re.compile(r'^(工作描述|内容[：:]|业绩[：:])')

# 章节标记（跳过但不停止解析）
SECTION_MARKER = re.compile(
    r'^(教育经|自我评价|语言能力|所获证书|项目经历|专业技能|培训经历)'
)

# 行末位置关键词（某些PDF会把地点放在岗位行后）
LOCATION_PATTERN = re.compile(r'^(青岛|济南|北京|上海|广州|深圳|杭州|成都|武汉|南京|天津|重庆|'
                              r'西安|苏州|郑州|长沙|合肥|大连|沈阳|长春|哈尔滨|昆明|'
                              r'贵阳|南宁|海口|兰州|西宁|银川|乌鲁木齐|拉萨|'
                              r'滨州|淄博|潍坊|烟台|威海|日照|临沂|德州|聊城|菏泽|'
                              r'泰安|济宁|枣庄|东营|莱芜|徐州|常州|无锡|南通|'
                              r'扬州|镇江|盐城|淮安|连云港|宿迁|泰州|宁波|温州|'
                              r'绍兴|嘉兴|湖州|金华|衢州|舟山|台州|丽水)\s*$')


# ============= 启发式公司识别（无标准后缀） =============

ENTITY_INDICATORS = [
    '商贸', '科技', '实业', '工贸', '制造', '电子', '信息',
    '智能', '网络', '软件', '硬件', '能源', '建设', '地产',
    '物业', '物流', '医药', '医疗', '环保', '通信', '光电',
    '微电子', '半导体', '自动化', '新材料', '集团', '控股',
    '股份', '咨询', '数据', '生物', '工程', '化工',
    '材料', '汽车', '设备', '仪器', '电器', '电气', '机械',
    '五金', '建材', '食品', '饮品', '服装', '纺织', '保险',
    '银行', '证券', '基金', '信托', '传媒', '文化', '旅游',
    '装饰', '家具', '家居', '厨卫', '电力', '水务', '燃气',
    '热力', '研究院', '研究所', '实验室', '中心', '平台',
    '酒业', '洋酒', '日化', '茶叶', '粮油', '乳业', '畜牧',
    '种业', '水产', '餐饮', '酒店', '广告', '印刷', '包装',
    # 注：'投资'、'设计'、'教育'、'服务' 已移除（太泛化，误伤描述行）
    # 短公司名 / 品牌名
    '八局', '建勘', '首创', '博睿', '岚之峰', '栈航', '安步',
    '赛美', '追觅', '恒元', '方策', '融创', '链湾', '鼎信',
    '博多', '火石', '涵养', '海氏', '兆冠', '璞禾', '美联',
    '君德', '全科', '士兰微', '慧尔视', '保融', '金实',
    '鲲鹏', '伊顿', '特来电',
    # V5 新增：极短品牌/公司名
    '美团', '哈啰', '滴滴', '快手', '字节', '腾讯', '百度',
    '阿里', '华为', '小米', '京东', '大玩家', '支付',
    '赤伏', '明日', '韩江', '征和', '驰马', '海创',
    '西部', '超人', '比邻', '矩阵', '哈雷',
]

# 非公司行关键词（描述、标签、UI文本）—— 严格排除
NON_COMPANY_KW = [
    r'^(查看|下载|扫码|点击|微信|智联|App|在电脑|ID[：:]|简历下载|YpXm)',
    r'^(运营管理|战略管理|市场营销|团队管理|项目管理|销售管理|客户管理|成本管理|现场管理|内部规章制度)\s*$',
    r'^(政府文案|执行力强|个人创业|空白市场|优化业务|供应链)\s*$',
    r'^(负责|协助|配合|参与|跟进|维护|拓展|整理|推广|收集|带领|学习|结合|完成|担任|全面|独立|兼职|兼任)',
    r'^(资产运营|产品运营|销售运营|市场运营|渠道运营|社区运营|B端运营|业务运营|内容运营|线下运营)\s*$',
    r'^(职责|业绩|技能|项目描述|工作职责|工作业绩|主要工作|核心工作|岗位职责)',
    r'^(第一阶段|第二阶段|第三阶段|第四阶段|第五阶段)',
    r'(项目$|项目描述|项目经历|采购项目|分析报告|管理流程)',
    # 节标记 — 仅完整匹配（加 ^ 锚定防子串误伤） 
    r'^(教育经\s*历|自我评价|语言能力|所获证书|项目经\s*历|专业技能|培训经历)\s*$',
    r'^(工作描述|内容[：:]|业绩[：:])',
    r'^(大客户|政府客户|企业客户|个人客户|院校客户|代理商|区域销售|渠道销售|电话销售|网络销售|门店销售|面销|陌拜|地推)\s*$',
    r'(管理客户|开发客户|维护客户|拓展市场|收集市场|推广产品|跟进客户|对接客户)',
    r'(原有客户|新增客户|潜在客户|意向客户|存量客户)',
]
NON_COMPANY_RE = re.compile('|'.join(f'({p})' for p in NON_COMPANY_KW))


def _is_heuristic_company(ls, lines, idx):
    """判断无后缀行是否可能是公司名。严格条件：含实体词 + 不匹配排除 + 附近有时间 + 时间后无教育关键词。"""
    if not re.search(r'[\u4e00-\u9fa5]', ls):
        return False
    # STRICT: 无后缀公司名通常 3-15 字符
    if len(ls) < 2 or len(ls) > 15:
        return False
    # 含中文标点的长行 = 描述，非公司名
    if len(ls) > 8 and re.search(r'[，。！？：；、]', ls):
        return False
    if NON_COMPANY_RE.search(ls):
        return False
    if _is_position_line(ls):
        return False
    if not any(ind in ls for ind in ENTITY_INDICATORS):
        return False

    # 必须在后续 8 行内找到时间
    found_j = -1
    for j in range(idx + 1, min(idx + 9, len(lines))):
        nl = lines[j].strip()
        if TIME_RANGE_RE.search(nl):
            found_j = j
            break
        if DURATION_RE.search(nl) and not re.search(r'[\u4e00-\u9fa5]{3,}', nl):
            found_j = j
            break
    if found_j < 0:
        return False

    # 反教育守卫
    for ej in range(found_j + 1, min(found_j + 5, len(lines))):
        al = lines[ej].strip()
        if EDUCATION_KEYWORDS.search(al):
            return False
        if re.match(r'^\d{4}\.\d{1,2}\s*-\s*\d{4}\.\d{1,2}\s*$', al.replace(' ', '')):
            return False

    return True

def parse_resume(raw_text):
    raw = raw_text
    result = {}

    # ---- NAME ----
    skip_kw = ['点击此处', '期望从事', 'ID', '简历下载', '扫码', '支持', '智联',
               '在线', '电脑端', '微信', 'App', 'YpXm', '在电脑']
    name = ""
    for line in raw.strip().split('\n'):
        ls = line.strip()
        if ls and not any(kw in ls for kw in skip_kw):
            clean = ls.replace('\u200b', '').replace(' ', '')
            if len(clean) <= 10:
                name = clean
                break
    result['name'] = name if name else "未知"

    # ---- GENDER ----
    m = re.search(r'(?<![a-zA-Z])(男|女)(?![a-zA-Z])', raw)
    result['gender'] = m.group(1) if m else ""

    # ---- AGE ----
    m = re.search(r'(\d+)\s*岁', raw)
    age = int(m.group(1)) if m else 0
    if age == 0:
        m = re.search(r'\((\d{4})\s*年', raw)
        if m:
            age = datetime.now().year - int(m.group(1))
    result['age'] = age

    # ---- RESIDENCE ----
    m = re.search(r'现居住地[：:]\s*(.+?)(?:\n)', raw)
    result['current_residence'] = m.group(1).strip() if m else ""

    # ---- JOB INTENTIONS ----
    intention_raw = ""
    m = re.search(r'求职意\s*向\s*\n(.*?)\n\s*工作经\s*历', raw, re.DOTALL)
    if m:
        intention_raw = m.group(1)

    intentions = []
    if intention_raw:
        fixed = fix_fragments(intention_raw)
        fixed = re.sub(r'\n\s*[｜|]\s*', '\n', fixed)

        parts = re.split(r'(全职|兼职)', fixed)
        intent_pairs, current = [], ""
        for part in parts:
            if part in ('全职', '兼职'):
                intent_pairs.append((current.strip(), part))
                current = ""
            else:
                current = part
        if current.strip():
            intent_pairs.append((current.strip(), ""))

        for content, _ in intent_pairs:
            content = content.strip()
            if not content:
                continue
            sub_lines = [l.strip() for l in content.split('\n') if l.strip()]
            if not sub_lines:
                continue

            position = re.sub(r'\s+', '', sub_lines[0])
            if re.match(r'^[、，,./]+$', position) or len(position) < 2:
                continue
            location = salary = industry = ""

            for sl in sub_lines[1:]:
                sl_clean = sl.strip().replace(' ', '')
                if not sl_clean:
                    continue
                sal_m = re.search(r'([\d,.]+[万千百元](?:-[\d,.]*[万千百元])?\s*/\s*[年月])', sl_clean)
                if not sal_m:
                    sal_m = re.search(r'([\d,.]+[万千百元])', sl_clean)
                if sal_m:
                    salary = re.sub(r'\s+', '', sal_m.group(1))
                    before = sl_clean[:sal_m.start()].strip()
                    if before and re.match(r'^[\u4e00-\u9fa5、，,]+$', before):
                        location = before
                    continue
                if re.match(r'^[\u4e00-\u9fa5]+[、，,][\u4e00-\u9fa5]+$', sl_clean) or \
                   (re.match(r'^[\u4e00-\u9fa5]{2,10}$', sl_clean) and
                    not re.search(r'[行业不限]', sl_clean)):
                    if not location:
                        location = sl_clean
                    continue
                if '行业' in sl_clean:
                    industry = sl_clean
                elif '不限' in sl_clean:
                    industry = '不限行业'
            if not industry:
                industry = "不限行业"
            intentions.append(f"【{position}】{salary}-{location}-{industry}")

    result['job_intentions'] = '\n'.join(intentions) if intentions else ""

    # ============ WORK EXPERIENCE (v7: 三元组切分 公司→岗位→时间→描述) ============

    idx_w = re.search(r'工作经\s*历', raw)
    if not idx_w:
        result['work_experiences'] = ""
        return result

    work_raw = raw[idx_w.end():]
    # 不再使用固定截断点 — PDF文本顺序不可靠（如孙女士的简历中，项目经历/教育经历
    # 标记出现在真实工作经历之前），硬截断会导致后续工作经历丢失。
    # 改为完全依赖逐行解析 + SECTION_MARKER + EDUCATION_KEYWORDS 过滤。
    # 教育内容误判由 auto_fix 在后续清理
    work_fixed = fix_fragments(work_raw)
    lines = work_fixed.strip().split('\n')
    n = len(lines)

    # ====== 新逻辑：基于三元组切分段落 ======

    # 辅助函数：是否可跳过（空行、纯数字、UI标签等）
    def _is_skip_line(ls):
        if not ls or len(ls) < 2:
            return True
        if re.match(r'^[\d\s,.，。/、万元月千\(\)（）\-/①-⑩]+$', ls):
            return True
        if SECTION_MARKER.match(ls):
            return True
        if EDUCATION_KEYWORDS.search(ls):
            return True
        if LOCATION_PATTERN.match(ls):
            return True
        if DESCRIPTION_MARKER.match(ls):
            return True
        if NON_COMPANY_RE.search(ls):
            return True
        return False

    # 辅助函数：是否为薪资行
    def _is_salary_line(ls):
        return bool(re.search(r'[\d,.]+[万千百元]\s*/\s*[月年]', ls) or
                   re.match(r'^[\d,.]+[万千百元]$', ls.replace(' ', '')))

    # 辅助函数：行是否可能是公司名（必须有实体词特征）
    def _is_possible_company(ls):
        if not re.search(r'[\u4e00-\u9fa5]', ls):
            return False
        if len(ls) < 2 or len(ls) > 30:
            return False
        if _is_skip_line(ls):
            return False
        if _is_position_line(ls):
            return False
        if _is_salary_line(ls):
            return False
        if TIME_RANGE_RE.search(ls):
            return False
        # 必须有公司特征：后缀 或 实体词
        has_suffix = bool(COMPANY_SUFFIX.search(ls))
        has_entity = any(ind in ls for ind in ENTITY_INDICATORS)
        # 4字及以上纯中文，不含实体词 → 大概率是行业标签/技能词，拒绝
        if not has_suffix and not has_entity:
            return False
        return True

    # Step 1: 找所有时间行索引
    time_indices = []
    for i, line in enumerate(lines):
        ls = line.strip()
        if TIME_RANGE_RE.search(ls):
            time_indices.append(i)

    if not time_indices:
        result['work_experiences'] = ""
        return result

    # Step 2: 对每个时间行，反向找岗位和公司
    experiences = []
    used_indices = set()

    for time_idx in time_indices:
        if time_idx in used_indices:
            continue

        # 反向找岗位行 — 找最远的（最接近公司的），不找括号里的子岗位
        # 遇到公司行提前终止，防止过度回溯（如越过公司搜到前面的标签行）
        pos_start = None
        for j in range(time_idx - 1, max(time_idx - 12, -1), -1):
            ls = lines[j].strip()
            if _is_skip_line(ls):
                continue
            if _is_likely_company(ls):
                break  # 遇到公司行，岗位搜索终止
            if _is_position_line(ls):
                pos_start = j

        if pos_start is None:
            continue

        # 检查：岗位和时间之间是否有"工作描述"（如有则说明这不是有效区块）
        has_desc_between = False
        for j in range(pos_start + 1, time_idx):
            if DESCRIPTION_MARKER.match(lines[j].strip()):
                has_desc_between = True
                break
        if has_desc_between:
            continue

        # 反向找公司行（岗位之前，非岗位非薪资）
        co_start = None
        for j in range(pos_start - 1, max(pos_start - 10, -1), -1):
            ls = lines[j].strip()
            if _is_skip_line(ls):
                continue
            if _is_position_line(ls):
                continue
            if _is_salary_line(ls):
                continue
            if _is_possible_company(ls):
                co_start = j
                break

        if co_start is None:
            continue

        # 提取公司名：从 co_start 到 pos_start（规则3：遇岗位名截断）
        company_parts = []
        co_end = pos_start  # 默认公司名结束于岗位开始
        for j in range(co_start, pos_start):
            ls = lines[j].strip()
            # "/" 是公司名和岗位的分隔符，之前的内容是公司名
            if ls == '/':
                co_end = j  # 公司名在此结束，岗位从此行后开始
                break
            if not ls or len(ls) < 2:
                continue  # 空行跳过
            if _is_skip_line(ls):
                continue  # 跳过标签行
            if _is_position_line(ls):
                co_end = j
                break
            if _is_salary_line(ls) or TIME_RANGE_RE.search(ls):
                break
            if _is_possible_company(ls):
                company_parts.append(ls)
        company = ''.join(company_parts).replace(' ', '')

        # 提取岗位名：从 co_end 到 薪资/时间（规则4：遇薪资或时间截断）
        position_parts = []
        pos_end = time_idx  # 默认岗位结束于时间行
        for j in range(co_end, time_idx):
            ls = lines[j].strip()
            if _is_skip_line(ls):
                continue
            if _is_salary_line(ls):
                pos_end = j
                break
            tm = TIME_RANGE_RE.search(ls)
            if tm:
                # 同行在时间前的部分也可能是岗位片段
                before = ls[:tm.start()].strip()
                before = re.sub(r'[\(\)（）]', '', before).strip()
                if before and not re.match(r'^[\d\s,./]+$', before):
                    position_parts.append(before.replace(' ', ''))
                pos_end = j
                break
            if _is_position_line(ls):
                position_parts.append(ls.replace(' ', ''))
        position = ''.join(position_parts)

        # 提取时间
        time_line = lines[time_idx].strip()
        time_m = TIME_RANGE_RE.search(time_line)
        start_t = time_m.group(1)
        end_t = time_m.group(2)
        duration = ""
        dur_text = ""
        for dj in range(time_idx + 1, min(time_idx + 5, n)):
            dur_text += lines[dj].strip()
            d_m = DURATION_RE.search(dur_text)
            if d_m:
                duration = re.sub(r'\s+', '', d_m.group(1))
                break
        time_range = f"{start_t}-{end_t}"
        if duration:
            time_range += f"({duration})"

        # 清理岗位名
        position = re.sub(r'\s+', '', position)
        position = re.sub(r'\d+年\d*个?月?|个月|\d{4}\.\d{1,2}', '', position)
        position = re.sub(r'[\d,.]+\s*[万千百元]\s*/\s*[月年]', '', position)
        position = re.sub(r'[\d,.]+\s*[万千百元]', '', position)

        # 标记已使用的时间行
        for j in range(co_start, time_idx + 1):
            used_indices.add(j)

        # 格式化
        formatted = f"【{position}】{company}-{time_range}"
        experiences.append(formatted)

    # 去重（同一公司+时间）
    seen = set()
    unique = []
    for exp in experiences:
        # 提取公司+时间作为key
        m = re.match(r'^【[^】]*】(.+)-(\d{4}\.\d{1,2}-\S+)$', exp)
        if m:
            key = m.group(1)[:10] + m.group(2)[:10]
            if key not in seen:
                seen.add(key)
                unique.append(exp)
        else:
            unique.append(exp)

    result['work_experiences'] = '\n'.join(unique) if unique else ""
    result['salary_info'] = []  # PDF 暂不支持结构化薪资提取，AI 从 job_intentions 文本中解析
    return result


def batch_extract(pdf_dir):
    pdf_files = sorted(Path(pdf_dir).glob("*.pdf"))
    if not pdf_files:
        raise FileNotFoundError(f"No PDF files found in {pdf_dir}")

    results = []
    for pdf_file in pdf_files:
        text = extract_text_from_pdf(pdf_file)
        parsed = parse_resume(text)
        parsed['filename'] = pdf_file.name
        results.append(parsed)

    return results


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python extract_resumes.py <pdf_directory> [output_json]")
        sys.exit(1)

    pdf_dir = sys.argv[1]
    output = sys.argv[2] if len(sys.argv) > 2 else "extracted_data.json"

    results = batch_extract(pdf_dir)
    with open(output, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"Extracted {len(results)} resumes -> {output}")
