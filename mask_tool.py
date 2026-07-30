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
    '薪资': {'func': lambda: get_fake().credit_card_number(), 'label': '银行卡号'},
    '工资': {'func': lambda: get_fake().credit_card_number(), 'label': '银行卡号'},
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


def apply_masking(df, col_rules):
    """纯函数：对 df 按 col_rules(列名->方法) 脱敏，返回 (masked_df, mappings)。

    - 所有列以文本处理：源文件用 dtype=str 读入，脱敏列再统一转文本，
      彻底避免 Excel 把数字型脱敏值（银行卡号/手机号/随机数字等）数值化、
      科学计数法化，或破坏前导零（如发票号 04403...）与小数格式（如金额 156800.00）。
    - 随机生成类方法的脱敏值在本列内保证唯一，否则还原时 reverse_map 会互相
      覆盖导致串档。
    """
    masked_df = df.copy()
    mappings = {}
    RANDOM_METHODS = {
        "姓名", "身份证号", "手机号", "银行卡号",
        "地址", "邮箱", "公司名", "自定义文本", "随机数字",
    }

    def apply_mask(val, method, col_map, used):
        if pd.isna(val):
            return val
        val_str = str(val)
        if val_str in col_map:
            return col_map[val_str]
        new_val = gen_masked_value(method, val_str)
        if method in RANDOM_METHODS and new_val in used:
            for _ in range(10):
                cand = gen_masked_value(method, val_str)
                if cand not in used:
                    new_val = cand
                    break
            else:
                n = 2
                while f"{new_val}#{n}" in used:
                    n += 1
                new_val = f"{new_val}#{n}"
        col_map[val_str] = new_val
        used.add(new_val)
        return new_val

    for col, method in col_rules.items():
        col_map = {}
        used = set()
        masked_df[col] = masked_df[col].apply(
            lambda x: apply_mask(x, method, col_map, used)
        )
        mappings[col] = col_map
    for col in col_rules:
        masked_df[col] = masked_df[col].astype(str)
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
        try:
            if uploaded_file.name.endswith('.csv'):
                encodings = ['utf-8-sig', 'utf-8', 'gbk', 'gb2312', 'latin-1']
                df = None
                for encoding in encodings:
                    try:
                        uploaded_file.seek(0)
                        # dtype=str：源文件一律按文本读入，保留前导零(发票号 04403...)
                        # 与小数格式(金额 156800.00)，避免被数值化后精度/格式损坏。
                        df = pd.read_csv(uploaded_file, encoding=encoding, dtype=str)
                        break
                    except UnicodeDecodeError:
                        continue
                if df is None:
                    st.error("无法识别文件编码，请尝试转换为 UTF-8 编码后重新上传")
                    st.info("提示：可以在 Excel 中另存为 CSV 时选择 UTF-8 编码")
                else:
                        df.columns = [str(c).replace('\ufeff', '').replace('\ufffe', '') for c in df.columns]
                        st.session_state.df = df
                        st.session_state.original_filename = uploaded_file.name
                        st.success(f"✅ 文件加载成功: {uploaded_file.name} ({len(df)} 行, {len(df.columns)} 列)")
            else:
                df = pd.read_excel(uploaded_file, dtype=str)
                df.columns = [str(c).replace('\ufeff', '').replace('\ufffe', '') for c in df.columns]
                st.session_state.df = df
                st.session_state.original_filename = uploaded_file.name
                st.success(f"✅ 文件加载成功: {uploaded_file.name} ({len(df)} 行, {len(df.columns)} 列)")
        except Exception as e:
            error_msg = str(e)
            if "Invalid file" in error_msg or "File is not a zip file" in error_msg:
                st.error("文件格式不正确，请确保上传的是有效的 Excel 或 CSV 文件")
            elif "Permission denied" in error_msg:
                st.error("文件被占用，请关闭 Excel 或其他程序后重试")
            else:
                st.error(f"文件读取失败: {e}")
                st.info("常见解决方案：\n1. 确保文件格式正确\n2. 检查文件是否损坏\n3. 尝试重新保存文件")

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
                            f"✅ 脱敏完成！分类：{eff_cat}　共处理 {total_cols} 列, "
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
    - 脱敏后文件：`.xlsx`（即「导出 & 映射表」里下载的脱敏后文件）
    - 映射表：`.json`（明文）或 `.bin`（加密，需密码）
    """)

    restore_masked = st.file_uploader(
        "① 拖入脱敏后的文件 (.xlsx / .xls)",
        type=['xlsx', 'xls'],
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
