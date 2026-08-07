"""
数据脱敏工具 - 本地运行，数据不外传
启动方式: streamlit run mask_tool.py   （或直接双击 启动.bat 自动配置环境）
浏览器打开 http://localhost:8501

功能：
- 拖拽上传 Excel / CSV，自动识别敏感列
- 为整份文档打「分类标签」（人事/财务/客户/供应商/其他，可自定义修改）
- 一键脱敏，分类联动推荐敏感字段
- 前端一键下载脱敏后文件 & 映射表（明文 / 加密）
- 拖入脱敏后文件 + 映射表，一键还原
"""

import streamlit as st
import pandas as pd
from faker import Faker
import json
import os
import io
import re
import hashlib
import base64
from collections import defaultdict
from datetime import datetime

# cryptography(Fernet) 为可选依赖：仅用于"加密映射表"功能。
# 改为懒加载——其原生扩展(Rust 构建)需要较新的 Visual C++ 运行库，
# 若该环境缺失，脱敏核心功能仍必须可用（详见 _get_fernet）。
def _get_fernet():
    try:
        from cryptography.fernet import Fernet
        return Fernet
    except Exception:
        raise RuntimeError(
            "加密功能不可用：当前环境缺少 cryptography 所需的运行库。"
            "脱敏功能不受影响；如需加密映射表，请更新 Visual C++ 运行库后重试。"
        )


# 初始化Faker（延迟加载以提高性能）
_fake_instance = None

def get_fake():
    global _fake_instance
    if _fake_instance is None:
        _fake_instance = Faker('zh_CN')
    return _fake_instance

def generate_key_from_password(password):
    """从密码生成加密密钥"""
    key = hashlib.sha256(password.encode()).digest()
    return base64.urlsafe_b64encode(key)

def encrypt_data(data, password):
    """使用密码加密数据"""
    key = generate_key_from_password(password)
    f = _get_fernet()(key)
    encrypted_data = f.encrypt(json.dumps(data, ensure_ascii=False).encode('utf-8'))
    return encrypted_data

def decrypt_data(encrypted_data, password):
    """使用密码解密数据"""
    key = generate_key_from_password(password)
    f = _get_fernet()(key)
    decrypted_data = f.decrypt(encrypted_data)
    return json.loads(decrypted_data.decode('utf-8'))


# ========== 文档分类配置 ==========
DOC_CATEGORIES = ["人事", "财务", "客户", "供应商", "其他"]

# 各分类常用的敏感字段关键词（用于联动推荐需要脱敏的列）
CATEGORY_HINTS = {
    "人事": ["姓名", "名字", "身份证", "手机", "电话", "地址", "邮箱", "邮件", "薪资", "工资"],
    "财务": ["银行卡", "卡号", "账号", "手机", "电话", "金额", "工资", "薪资"],
    "客户": ["姓名", "名字", "手机", "电话", "邮箱", "邮件", "地址"],
    "供应商": ["公司", "账号", "手机", "电话", "邮箱", "邮件", "地址"],
    "其他": [],
}


# 脱敏规则配置
MASK_RULES = {
    '姓名': {'func': lambda: get_fake().name(), 'label': '姓名'},
    '名字': {'func': lambda: get_fake().name(), 'label': '姓名'},
    '身份证': {'func': lambda: get_fake().ssn(), 'label': '身份证号'},
    '手机': {'func': lambda: get_fake().phone_number(), 'label': '手机号'},
    '电话': {'func': lambda: get_fake().phone_number(), 'label': '电话'},
    '银行卡': {'func': lambda: get_fake().credit_card_number(), 'label': '银行卡号'},
    '卡号': {'func': lambda: get_fake().credit_card_number(), 'label': '银行卡号'},
    '地址': {'func': lambda: get_fake().address(), 'label': '地址'},
    '邮箱': {'func': lambda: get_fake().email(), 'label': '邮箱'},
    '邮件': {'func': lambda: get_fake().email(), 'label': '邮箱'},
    '公司': {'func': lambda: get_fake().company(), 'label': '公司名'},
    '账号': {'func': lambda: get_fake().credit_card_number(), 'label': '银行卡号'},
    '金额': {'func': lambda: None, 'label': '金额(保留格式)'},
    '薪资': {'func': lambda: None, 'label': '金额(保留格式)'},
    '工资': {'func': lambda: None, 'label': '金额(保留格式)'},
    '发票': {'func': lambda: get_fake().random_number(digits=20), 'label': '随机数字'},
    '税号': {'func': lambda: get_fake().random_number(digits=18), 'label': '随机数字'},
    '编号': {'func': lambda: get_fake().random_number(digits=10), 'label': '随机数字'},
}

# 脱敏方法配置（消除硬编码）
MASK_METHODS = {
    '姓名': {'label': '姓名', 'description': '生成中文姓名'},
    '身份证号': {'label': '身份证号', 'description': '生成18位身份证号'},
    '手机号': {'label': '手机号', 'description': '生成11位手机号'},
    '银行卡号': {'label': '银行卡号', 'description': '生成银行卡号'},
    '地址': {'label': '地址', 'description': '生成中文地址'},
    '邮箱': {'label': '邮箱', 'description': '生成邮箱地址'},
    '公司名': {'label': '公司名', 'description': '生成公司名称'},
    '自定义文本': {'label': '自定义文本', 'description': '生成自定义文本'},
    '随机数字': {'label': '随机数字', 'description': '生成随机数字，保持原格式'},
    '金额(保留格式)': {'label': '金额(保留格式)', 'description': '随机金额，保留小数位/千分位/负号格式'},
    '部分遮蔽(保留前3位)': {'label': '部分遮蔽(保留前3位)', 'description': '保留前3位，其余用*替换'},
    '完全遮蔽': {'label': '完全遮蔽', 'description': '全部用*替换'},
    '保留格式遮蔽': {'label': '保留格式遮蔽', 'description': '保留数字格式，替换为随机数字'},
}

# 获取脱敏方法列表
MASK_METHOD_LIST = list(MASK_METHODS.keys())


def safe_name(s, maxlen=40):
    """去掉文件名中的非法字符，并截断长度，用于导出文件名"""
    s = re.sub(r'[\\/:*?"<>|]', '_', s)
    return s[:maxlen]


def keep_long_numbers_as_text(df):
    """把「数值型且含超过 15 位整数」的列强制转为文本。

    根因：Excel 以双精度浮点读写数字，有效数字仅约 15 位；19 位发票号 /
    长账号在导出↔读回时会被截断为 0。脱敏列此前已转文本，这里兜底
    那些「未脱敏的长数字列」，避免导出文件精度损坏。
    """
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            s = df[col].dropna()
            if len(s) and (s.abs() > 10 ** 15).any():
                try:
                    df[col] = df[col].astype('int64').astype(str)
                except (ValueError, OverflowError):
                    df[col] = df[col].astype(str)
    return df


def clean_column_names(cols):
    """清洗列名：去掉空白/BOM；空名与 Unnamed 改为「列N」；重名追加 _2/_3。

    解决 A1 空置或全空列时 pandas 生成 Unnamed: N 的问题，
    让表头对用户可读、可勾选。
    """
    seen = {}
    out = []
    for i, c in enumerate(cols, 1):
        name = "" if pd.isna(c) else str(c).replace('\ufeff', '').replace('\ufffe', '').strip()
        if not name or name.startswith('Unnamed'):
            name = f"列{i}"
        n = seen.get(name, 0) + 1
        seen[name] = n
        out.append(name if n == 1 else f"{name}_{n}")
    return out


def pick_header_row(raw, max_scan=20):
    """自动判断表头行（应对「首行为合并标题」「A1 空置」等情况）：
    - 首行非空单元格 >= 2：首行即表头
    - 首行全空：取第一个非空行
    - 首行仅 1 格（如合并单元格标题）且下一行非空更多：取下一行
    """
    if len(raw) == 0:
        return 0
    first = raw.iloc[0].notna().sum()
    if first >= 2:
        return 0
    if first == 0:
        for i in range(min(max_scan, len(raw))):
            if raw.iloc[i].notna().sum() > 0:
                return i
        return 0
    if len(raw) > 1 and raw.iloc[1].notna().sum() > first:
        return 1
    return 0


def build_df(raw, header_idx):
    """按表头行构建正式 DataFrame：清洗列名，删除全空行与全空列。"""
    if len(raw) == 0:
        return pd.DataFrame()
    cols = clean_column_names(raw.iloc[header_idx].tolist())
    df = raw.iloc[header_idx + 1:].copy()
    df.columns = cols
    df = df.dropna(how='all').dropna(axis=1, how='all')
    return df.reset_index(drop=True)


def read_file_raw(uploaded_file, sheet_name=0):
    """把上传文件读成原始网格(header=None, dtype=str)，不丢空表头信息。

    CSV：自动探测编码(UTF-8/GBK/GB2312/Latin-1)与分隔符(, ; Tab)，
    异常行跳过（财务导出的脏 CSV 常见）。
    """
    name = uploaded_file.name
    if name.endswith('.csv'):
        encodings = ['utf-8-sig', 'utf-8', 'gbk', 'gb2312', 'latin-1']
        for encoding in encodings:
            try:
                uploaded_file.seek(0)
                return pd.read_csv(uploaded_file, encoding=encoding, dtype=str,
                                   header=None, sep=None, engine='python',
                                   on_bad_lines='skip')
            except UnicodeDecodeError:
                continue
        return None
    uploaded_file.seek(0)
    return pd.read_excel(uploaded_file, dtype=str, header=None, sheet_name=sheet_name)


def gen_masked_value(method, val_str):
    """根据脱敏方法生成一个假值（纯函数，便于碰撞时重试）"""
    if method == "姓名":
        return get_fake().name()
    elif method == "身份证号":
        return get_fake().ssn()
    elif method == "手机号":
        return get_fake().phone_number()
    elif method == "银行卡号":
        return get_fake().credit_card_number()
    elif method == "地址":
        return get_fake().address()
    elif method == "邮箱":
        return get_fake().email()
    elif method == "公司名":
        return get_fake().company()
    elif method == "自定义文本":
        return f"脱敏_{get_fake().hex_color().replace('#', '')}"
    elif method == "随机数字":
        digits = len(re.sub(r'[^0-9]', '', val_str)) or 6
        random_num = get_fake().random_number(digits=digits)
        if val_str.startswith('0') and len(val_str) > len(str(random_num)):
            random_num = str(random_num).zfill(len(val_str))
        return str(random_num)
    elif method == "金额(保留格式)":
        # 每个数字段替换为同长度随机数（段首 1-9 防掉位），
        # 保留负号/千分位/小数点等格式，适合薪资、金额等列
        def _rand_run(n):
            if n == 1:
                return str(get_fake().random_int(min=0, max=9))
            return str(get_fake().random_int(min=10 ** (n - 1), max=10 ** n - 1))
        res = []
        run = 0
        for ch in val_str:
            if ch.isdigit():
                run += 1
            else:
                if run:
                    res.append(_rand_run(run))
                    run = 0
                res.append(ch)
        if run:
            res.append(_rand_run(run))
        return "".join(res)
    elif method == "部分遮蔽(保留前3位)":
        if len(val_str) > 3:
            return val_str[:3] + "*" * (len(val_str) - 3)
        return "*" * len(val_str)
    elif method == "完全遮蔽":
        return "*" * len(val_str)
    elif method == "保留格式遮蔽":
        res = ""
        for ch in val_str:
            res += str(get_fake().random_digit()) if ch.isdigit() else ch
        return res
    else:
        return val_str


def _base_masked_value(method, val_str):
    """确定性生成脱敏值：同一原值跨文件、跨运行得到同一假值。

    以原值的 SHA-256 摘要作为 Faker 种子，使随机类方法的输出成为原值的
    纯函数（财务可跨文件对账/汇总）；遮蔽类方法本身由原值结构决定。
    """
    seed = int(hashlib.sha256(val_str.encode('utf-8')).hexdigest()[:12], 16)
    get_fake().seed_instance(seed)
    return gen_masked_value(method, val_str)


def apply_masking(df, col_rules):
    """纯函数：对 df 按 col_rules(列名->方法) 脱敏，返回 (masked_df, mappings)。

    - 所有列以文本处理：源文件用 dtype=str 读入，避免 Excel 把数字型脱敏值
      数值化、科学计数法化，或破坏前导零（如发票号 04403...）与小数格式
      （如金额 156800.00）；空单元格保持为空（不变成字符串 "nan"）。
    - 确定性：同一原值 -> 同一假值（跨文件/跨运行一致），脱敏可复现。
    - 列内假值唯一：不同原值绝不映射到同一假值（遮蔽类方法如部分遮蔽/完全
      遮蔽不同原值可能产生相同结果，统一追加确定性后缀解决），保证还原时
      reverse_map 不会互相覆盖导致串档。
    """
    masked_df = df.copy()
    mappings = {}

    for col, method in col_rules.items():
        col_map = {}
        vals = df[col].dropna().astype(str)
        # 第 1 趟：对每个去重原值生成基准假值（确定性）
        base_of = {v: _base_masked_value(method, v) for v in vals.unique()}
        # 第 2 趟：基准假值冲突（不同原值 -> 同一假值）时，按原值追加确定性
        # 后缀，保证列内唯一且结果与行序无关
        groups = defaultdict(list)
        for v, b in base_of.items():
            groups[b].append(v)
        for b, vs in groups.items():
            if len(vs) == 1:
                col_map[vs[0]] = b
            else:
                for v in vs:
                    tag = hashlib.sha256(f"{col}\x00{v}".encode('utf-8')).hexdigest()[:6]
                    col_map[v] = f"{b}#{tag}"
        masked_df[col] = vals.map(col_map)
        mappings[col] = col_map
    return masked_df, mappings


def restore_with_mapping(masked_df, mappings):
    """纯函数：用 mappings 还原 masked_df，返回 (restored_df, restored_count)。"""
    df = masked_df.copy()
    restored_count = 0
    for col, col_map in mappings.items():
        if col in df.columns and isinstance(col_map, dict):
            def _norm(v):
                if isinstance(v, float) and v.is_integer():
                    return str(int(v))
                return str(v)
            reverse_map = {}
            for original, masked in col_map.items():
                k = _norm(masked)
                if k not in reverse_map:
                    reverse_map[k] = original
            df[col] = df[col].map(lambda v: reverse_map.get(_norm(v), v))
            restored_count += 1
    return df, restored_count


def recommend_cols(df, category):
    """根据列名关键词 + 分类提示，推荐需要脱敏的列"""
    rec = set()
    hints = CATEGORY_HINTS.get(category, [])
    for col in df.columns:
        c = str(col)
        if any(k in c for k in MASK_RULES):
            rec.add(col)
        if any(k in c for k in hints):
            rec.add(col)
    return list(rec)


def default_method_for(col):
    """根据列名推断默认脱敏方法"""
    for keyword, rule in MASK_RULES.items():
        if keyword in str(col):
            return rule['label']
    return '自定义文本'


st.set_page_config(page_title="数据脱敏工具", page_icon="🔒", layout="wide")
st.title("🔒 数据脱敏工具")
st.caption("所有处理均在本地完成，数据不会上传到任何外部服务器")

# 初始化 session state
if 'df' not in st.session_state:
    st.session_state.df = None
if 'mappings' not in st.session_state:
    st.session_state.mappings = {}
if 'masked_df' not in st.session_state:
    st.session_state.masked_df = None
if 'doc_category' not in st.session_state:
    st.session_state.doc_category = "其他"
if 'original_filename' not in st.session_state:
    st.session_state.original_filename = ""
if 'restored_df' not in st.session_state:
    st.session_state.restored_df = None
if 'loaded_name' not in st.session_state:
    st.session_state.loaded_name = ""
if 'raw_df' not in st.session_state:
    st.session_state.raw_df = None
if 'raw_sheet' not in st.session_state:
    st.session_state.raw_sheet = None
if 'sheet_names' not in st.session_state:
    st.session_state.sheet_names = []
if 'sheet_sel' not in st.session_state:
    st.session_state.sheet_sel = None
if 'header_row' not in st.session_state:
    st.session_state.header_row = -1  # -1 = 自动检测
if 'df_header' not in st.session_state:
    st.session_state.df_header = None

# ========== 侧边栏：简短说明 ==========
with st.sidebar:
    st.header("🧭 使用流程")
    st.markdown("""
    1. **上传 & 预览**：拖入 Excel/CSV，选择文档分类
    2. **脱敏处理**：勾选列、配置规则、一键脱敏
    3. **导出 & 映射表**：下载脱敏文件 / 一键下载映射表
    4. **还原数据**：拖入脱敏文件 + 映射表，一键还原
    """)
    st.divider()
    st.info("💡 映射表含原始↔脱敏对照，属敏感文件，请仅内网保管。")

# ========== 主界面：四个标签页 ==========
tab1, tab2, tab3, tab4 = st.tabs([
    "📤 上传 & 预览",
    "🔒 脱敏处理",
    "📋 导出 & 映射表",
    "🔄 还原数据",
])

# ---------- Tab 1：上传 & 预览 ----------
with tab1:
    st.subheader("📥 拖入文档（Excel / CSV）")
    uploaded_file = st.file_uploader(
        "将文件拖拽到此处，或点击选择",
        type=['xlsx', 'xls', 'csv'],
        key="main_uploader",
        help="支持 .xlsx / .xls / .csv，文件不会离开本机"
    )

    if uploaded_file:
        # 换文件时重置状态：清空旧的脱敏/还原结果，避免串到新文件
        if st.session_state.loaded_name != uploaded_file.name:
            st.session_state.loaded_name = uploaded_file.name
            st.session_state.original_filename = uploaded_file.name
            st.session_state.raw_df = None
            st.session_state.raw_sheet = None
            st.session_state.header_row = -1
            st.session_state.masked_df = None
            st.session_state.mappings = {}
            st.session_state.restored_df = None
            # Excel 多工作表：列出供选择
            sheet_names = []
            if not uploaded_file.name.endswith('.csv'):
                try:
                    uploaded_file.seek(0)
                    sheet_names = pd.ExcelFile(uploaded_file).sheet_names
                except Exception:
                    sheet_names = []
            st.session_state.sheet_names = sheet_names
            st.session_state.sheet_sel = sheet_names[0] if sheet_names else None

        # Excel 多 Sheet 选择（切换后重新读取对应工作表）
        if st.session_state.sheet_names:
            st.session_state.sheet_sel = st.selectbox(
                "选择工作表",
                options=st.session_state.sheet_names,
                index=st.session_state.sheet_names.index(st.session_state.sheet_sel)
                if st.session_state.sheet_sel in st.session_state.sheet_names else 0,
                key=f"sheet_sel_{st.session_state.loaded_name}",
            )

        # 读取原始网格（新文件或切换工作表时）
        if (st.session_state.raw_df is None
                or st.session_state.raw_sheet != st.session_state.sheet_sel):
            try:
                raw = read_file_raw(uploaded_file, st.session_state.sheet_sel)
                if raw is None:
                    st.error("无法识别文件编码，请尝试转换为 UTF-8 编码后重新上传")
                    st.info("提示：可以在 Excel 中另存为 CSV 时选择 UTF-8 编码")
                    st.stop()
                st.session_state.raw_df = raw
                st.session_state.raw_sheet = st.session_state.sheet_sel
            except Exception as e:
                error_msg = str(e)
                if "Invalid file" in error_msg or "File is not a zip file" in error_msg:
                    st.error("文件格式不正确，请确保上传的是有效的 Excel 或 CSV 文件")
                elif "Permission denied" in error_msg:
                    st.error("文件被占用，请关闭 Excel 或其他程序后重试")
                else:
                    st.error(f"文件读取失败: {e}")
                    st.info("常见解决方案：\n1. 确保文件格式正确\n2. 检查文件是否损坏\n3. 尝试重新保存文件")
                st.stop()

        raw = st.session_state.raw_df

        # 表头行：默认自动检测，可手动指定（如首行为合并标题时）
        max_header = min(10, len(raw))
        hdr_opts = ["自动检测"] + [f"第{i}行" for i in range(1, max_header + 1)]
        hdr_idx = 0 if st.session_state.header_row < 0 else min(st.session_state.header_row, len(hdr_opts) - 1)
        hdr_sel = st.selectbox(
            "表头所在行",
            options=hdr_opts,
            index=hdr_idx,
            key=f"hdr_sel_{st.session_state.loaded_name}",
            help="系统已自动判断表头行；若首行是合并标题等，可手动指定"
        )
        if hdr_sel == "自动检测":
            st.session_state.header_row = -1
            header_idx = pick_header_row(raw)
        else:
            header_idx = int(hdr_sel.replace("第", "").replace("行", "")) - 1
            st.session_state.header_row = header_idx

        df = build_df(raw, header_idx)
        if len(df) == 0:
            st.error("文件没有有效数据：表头下没有任何非空数据行")
            st.stop()
        # 表头/工作表变化会重建 df：同步清空旧的脱敏/还原结果，避免列错位
        if st.session_state.get('df_header') != header_idx:
            st.session_state.df_header = header_idx
            st.session_state.masked_df = None
            st.session_state.mappings = {}
            st.session_state.restored_df = None
        st.session_state.df = df
        st.success(f"✅ 文件加载成功: {uploaded_file.name} ({len(df)} 行, {len(df.columns)} 列)")

    if st.session_state.df is not None:
        df = st.session_state.df

        # ===== 文档分类（可下拉选择并手动修改）=====
        st.subheader("📂 文档分类")
        cur = st.session_state.doc_category
        current_choice = cur if cur in DOC_CATEGORIES else "自定义"
        idx = (DOC_CATEGORIES + ["自定义"]).index(current_choice)
        cat = st.selectbox(
            "选择文档分类（可修改）",
            options=DOC_CATEGORIES + ["自定义"],
            index=idx,
            key="cat_sel",
            help="分类将写入导出文件名与映射表，便于归档"
        )
        if cat == "自定义":
            custom = st.text_input(
                "输入自定义分类",
                value="" if cur in DOC_CATEGORIES else cur,
                key="cat_custom"
            )
            effective_cat = (custom or "其他").strip() or "其他"
        else:
            effective_cat = cat
        st.session_state.doc_category = effective_cat
        hints = CATEGORY_HINTS.get(effective_cat, [])
        st.caption(f"当前分类：**{effective_cat}**　|　该分类常用敏感字段：{', '.join(hints) if hints else '—'}")

        st.divider()
        st.subheader("数据预览 (前 10 行)")
        st.dataframe(df.head(10), use_container_width=True)

        st.subheader("列信息")
        col_info = []
        for col in df.columns:
            sample = df[col].dropna().head(3).tolist()
            col_info.append({
                "列名": col,
                "数据类型": str(df[col].dtype),
                "非空数量": df[col].count(),
                "示例值": ", ".join(str(v) for v in sample[:2])
            })
        st.dataframe(pd.DataFrame(col_info), use_container_width=True)

# ---------- Tab 2：脱敏处理 ----------
with tab2:
    if st.session_state.df is None:
        st.info("请先在「上传 & 预览」标签页中上传文件")
    else:
        df = st.session_state.df
        eff_cat = st.session_state.doc_category
        st.subheader(f"🔒 选择要脱敏的列　（当前文档分类：{eff_cat}）")

        auto_recommend = recommend_cols(df, eff_cat)

        selected_cols = st.multiselect(
            "勾选需要脱敏的列",
            options=df.columns.tolist(),
            default=auto_recommend,
            help="系统已根据列名与文档分类自动推荐敏感列，你可以手动调整"
        )

        if selected_cols:
            st.subheader("配置脱敏规则")
            col_rules = {}
            cols_per_row = 3
            for i in range(0, len(selected_cols), cols_per_row):
                row_cols = st.columns(cols_per_row)
                for j, col in enumerate(selected_cols[i:i + cols_per_row]):
                    with row_cols[j]:
                        default_method = default_method_for(col)
                        method = st.selectbox(
                            f"**{col}**",
                            options=MASK_METHOD_LIST,
                            index=MASK_METHOD_LIST.index(default_method) if default_method in MASK_METHOD_LIST else 0,
                            key=f"rule_{col}"
                        )
                        col_rules[col] = method

            st.divider()
            if st.button("🔒 开始脱敏", type="primary", use_container_width=True):
                if not col_rules:
                    st.warning("请至少选择一列进行脱敏")
                else:
                    try:
                        progress = st.progress(0)
                        status_text = st.empty()
                        status_text.text("正在脱敏...")
                        masked_df, mappings = apply_masking(df, col_rules)
                        progress.progress(1.0)
                        st.session_state.masked_df = masked_df
                        st.session_state.mappings = mappings
                        progress.empty()
                        status_text.empty()

                        total_values = sum(len(v) for v in mappings.values())
                        st.success(
                            f"✅ 脱敏完成！分类：{eff_cat}　共处理 {len(col_rules)} 列, "
                            f"{total_values} 个唯一值"
                        )
                    except Exception as e:
                        st.error(f"脱敏处理失败: {e}")
                        st.info("请检查数据格式是否正确，或尝试减少同时处理的列数")

        if st.session_state.masked_df is not None:
            st.subheader("脱敏结果预览")
            st.dataframe(st.session_state.masked_df.head(10), use_container_width=True)

# ---------- Tab 3：导出 & 映射表 ----------
with tab3:
    if st.session_state.masked_df is None:
        st.info("请先完成脱敏（在「脱敏处理」标签页点击「开始脱敏」）")
    else:
        eff_cat = st.session_state.doc_category
        base = safe_name(os.path.splitext(st.session_state.original_filename)[0] or "数据")
        masked_name = f"脱敏后_{eff_cat}_{base}.xlsx"

        st.subheader("📥 下载脱敏后文件")
        st.caption(f"将直接下载到你的电脑（默认下载文件夹）：`{masked_name}`")
        buf = io.BytesIO()
        keep_long_numbers_as_text(st.session_state.masked_df.copy()).to_excel(
            buf, index=False, engine='openpyxl')
        st.download_button(
            label="📥 下载脱敏后文件",
            data=buf.getvalue(),
            file_name=masked_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

        # CSV 导出：纯文本，绝不会像 Excel 那样把长数字/前导零/小数格式数值化，
        # 对发票号、银行账号、金额等字段是「零精度损失」的最安全归档格式。
        csv_bytes = st.session_state.masked_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
        st.download_button(
            label="📥 下载脱敏后文件 (CSV · 零精度损失，推荐归档)",
            data=csv_bytes,
            file_name=masked_name.replace('.xlsx', '.csv'),
            mime="text/csv",
            use_container_width=True
        )

        st.divider()
        st.subheader("🗂️ 一键下载映射表（还原用）")
        st.caption("映射表记录「原始值 ↔ 脱敏值」对照，是还原的唯一依据，请妥善保存。")

        mapping_obj = {
            "_meta": {
                "category": eff_cat,
                "source_file": st.session_state.original_filename,
                "created_at": datetime.now().isoformat(timespec='seconds'),
                "columns": list(st.session_state.mappings.keys()),
            },
            "mappings": st.session_state.mappings,
        }
        mapping_json = json.dumps(mapping_obj, ensure_ascii=False, indent=2)
        st.download_button(
            label="📥 一键下载映射表 (JSON)",
            data=mapping_json.encode('utf-8'),
            file_name=f"映射表_{eff_cat}_{base}.json",
            mime="application/json",
            type="primary",
            use_container_width=True
        )

        with st.expander("🔐 需要加密映射表？（更安全，还原时需密码）"):
            encrypt_password = st.text_input(
                "加密密码",
                type="password",
                help="设置密码后，映射表将被加密，还原时需要输入相同密码"
            )
            if encrypt_password:
                try:
                    encrypted_data = encrypt_data(mapping_obj, encrypt_password)
                    st.download_button(
                        label="📥 下载加密映射表 (.bin)",
                        data=encrypted_data,
                        file_name=f"加密映射表_{eff_cat}_{base}.bin",
                        mime="application/octet-stream",
                        use_container_width=True
                    )
                    st.success("加密映射表已就绪，可下载")
                except Exception as e:
                    st.error(f"加密失败: {e}")
                    st.info("脱敏功能不受影响；如需加密映射表，请更新 Visual C++ 运行库后重试。")

        with st.expander("⚠️ 映射表安全提示"):
            st.warning("""
            **映射表包含原始数据和脱敏数据的对照关系，属于敏感文件！**
            - 仅限内网存储
            - 分析还原后立即删除
            - 不得通过邮件/微信等发送
            - 建议设置密码加密后保管
            - 加密映射表还原时需要输入密码
            """)

# ---------- Tab 4：还原数据 ----------
with tab4:
    st.subheader("🔄 拖入文件，一键还原")
    st.markdown("""
    将 **脱敏后的文件** 与 **映射表** 分别拖入下方，点击「一键还原」即可恢复原始数据。
    - 脱敏后文件：`.xlsx` / `.csv`（即「导出 & 映射表」里下载的脱敏后文件）
    - 映射表：`.json`（明文）或 `.bin`（加密，需密码）
    """)

    restore_masked = st.file_uploader(
        "① 拖入脱敏后的文件 (.xlsx / .xls / .csv)",
        type=['xlsx', 'xls', 'csv'],
        key="restore_masked"
    )
    restore_mapping = st.file_uploader(
        "② 拖入映射表 (.json / .bin)",
        type=['json', 'bin'],
        key="restore_mapping"
    )

    decrypt_password = None
    if restore_mapping and restore_mapping.name.endswith('.bin'):
        decrypt_password = st.text_input(
            "映射表密码（加密映射表需要）",
            type="password",
            key="restore_pwd"
        )

    if restore_masked and restore_mapping:
        if st.button("🔄 一键还原", type="primary", use_container_width=True):
            try:
                # 解析映射表（兼容 旧格式 / 新格式含 _meta）
                raw = restore_mapping.read()
                if restore_mapping.name.endswith('.bin'):
                    if not decrypt_password:
                        st.error("请输入解密密码")
                        st.stop()
                    try:
                        obj = decrypt_data(raw, decrypt_password)
                    except Exception:
                        st.error("解密失败：密码错误或文件损坏")
                        st.info("请确认密码是否正确，或尝试使用原始映射表")
                        st.stop()
                else:
                    try:
                        obj = json.loads(raw.decode('utf-8'))
                    except json.JSONDecodeError:
                        st.error("映射表格式错误：不是有效的 JSON 文件")
                        st.stop()

                if isinstance(obj, dict) and "mappings" in obj:
                    mappings = obj["mappings"]
                    meta = obj.get("_meta", {})
                    if meta:
                        st.caption(
                            f"映射表信息：分类={meta.get('category','—')}　"
                            f"来源={meta.get('source_file','—')}　"
                            f"创建={meta.get('created_at','—')}"
                        )
                else:
                    mappings = obj

                masked_df = None
                if restore_masked.name.endswith('.csv'):
                    encodings = ['utf-8-sig', 'utf-8', 'gbk', 'gb2312', 'latin-1']
                    for encoding in encodings:
                        try:
                            restore_masked.seek(0)
                            masked_df = pd.read_csv(restore_masked, encoding=encoding, dtype=str)
                            break
                        except UnicodeDecodeError:
                            continue
                    if masked_df is None:
                        st.error("无法识别 CSV 编码，请转换为 UTF-8 后重试")
                        st.stop()
                else:
                    masked_df = pd.read_excel(restore_masked, dtype=str)
                masked_df, restored_count = restore_with_mapping(masked_df, mappings)

                st.session_state.restored_df = masked_df
                st.success(f"✅ 还原成功！已还原 {restored_count} 列数据")
            except Exception as e:
                st.error(f"还原失败: {e}")
                st.info("请检查映射表和脱敏文件是否匹配（需为同一份数据脱敏所得）")

    if 'restored_df' in st.session_state and st.session_state.restored_df is not None:
        st.subheader("还原结果预览")
        st.dataframe(st.session_state.restored_df.head(10), use_container_width=True)
        rbase = safe_name(os.path.splitext(st.session_state.original_filename)[0] or "数据")
        rbuf = io.BytesIO()
        keep_long_numbers_as_text(st.session_state.restored_df.copy()).to_excel(
            rbuf, index=False, engine='openpyxl')
        st.download_button(
            label="📥 下载还原后文件",
            data=rbuf.getvalue(),
            file_name=f"还原后_{rbase}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

        # CSV 导出（零精度损失）
        rcsv = st.session_state.restored_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
        st.download_button(
            label="📥 下载还原后文件 (CSV · 零精度损失)",
            data=rcsv,
            file_name=f"还原后_{rbase}.csv",
            mime="text/csv",
            use_container_width=True
        )
