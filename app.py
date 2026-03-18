import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
from datetime import datetime
from typing import Dict, Optional
from io import BytesIO
import json
import os

# ===================== 全局配置与常量定义 =====================
st.set_page_config(page_title="Temu店铺数据分析工具", layout="wide")
st.title("📊 Temu 店铺数据分析工具（双月对比+销量分析版）")

# --- 初始化session状态 ---
# 登录状态
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False

# 文件上传状态持久化
if 'uploaded_file1' not in st.session_state:
    st.session_state.uploaded_file1 = None  # 本月文件
if 'uploaded_file2' not in st.session_state:
    st.session_state.uploaded_file2 = None  # 上月文件
if 'df_current' not in st.session_state:
    st.session_state.df_current = None     # 本月解析后的数据
if 'df_last' not in st.session_state:
    st.session_state.df_last = None        # 上月解析后的数据
if 'metrics_current' not in st.session_state:
    st.session_state.metrics_current = None # 本月计算指标
if 'metrics_last' not in st.session_state:
    st.session_state.metrics_last = None    # 上月计算指标

CONFIG = {
    "SUPPORTED_FILE_TYPES": ["xlsx", "csv"],
    "ENCODING": "utf-8",
    "DECIMAL_PLACES": 2,
    "FIXED_PASSWORD": "123456"  # 固定密码，不可修改
}

# ===================== 配置文件持久化 =====================
CONFIG_FILE = "alert_config.json"

def load_alert_config():
    """从文件加载警戒值配置"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {
                "ORDER_MARGIN_RATE_THRESHOLD": 20.0,
                "OPERATE_MARGIN_RATE_THRESHOLD": 15.0,
                "SALES_QUANTITY_THRESHOLD": 100,
                "UNIT_PRICE_THRESHOLD": 50.0,
                "UNIT_PROFIT_THRESHOLD": 10.0
            }
    else:
        return {
            "ORDER_MARGIN_RATE_THRESHOLD": 20.0,
            "OPERATE_MARGIN_RATE_THRESHOLD": 15.0,
            "SALES_QUANTITY_THRESHOLD": 100,
            "UNIT_PRICE_THRESHOLD": 50.0,
            "UNIT_PROFIT_THRESHOLD": 10.0
        }

def save_alert_config(config):
    """保存警戒值配置到文件"""
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)

# 自定义警戒值session（从文件加载）
if 'alert_config' not in st.session_state:
    st.session_state.alert_config = load_alert_config()

# ===================== 密码验证函数 =====================
def check_password():
    """简单的密码验证，密码固定为123456"""
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔐 登录")
    
    # 如果已经登录，直接返回True
    if st.session_state.authenticated:
        st.sidebar.success("✅ 已登录")
        if st.sidebar.button("退出登录"):
            st.session_state.authenticated = False
            st.rerun()
        return True
    
    # 未登录则显示密码输入框
    pwd = st.sidebar.text_input("请输入访问密码", type="password", key="login_pwd")
    
    if st.sidebar.button("登录"):
        if pwd == CONFIG["FIXED_PASSWORD"]:
            st.session_state.authenticated = True
            st.success("✅ 密码正确，欢迎使用！")
            st.rerun()
        else:
            st.sidebar.error("❌ 密码错误！")
    
    return False

# ===================== 自定义警戒值设置函数（已添加永久保存） =====================
def render_alert_config_panel():
    st.subheader("⚙️ 自定义警戒值设置")
    st.info("修改后点击【保存设置】生效，所有分析页面将使用新的警戒值判断")
    
    col1, col2 = st.columns(2)
    with col1:
        # 毛利率相关
        order_margin = st.number_input(
            "订单毛利率警戒值(%)",
            min_value=0.0, max_value=100.0, step=0.1,
            value=st.session_state.alert_config["ORDER_MARGIN_RATE_THRESHOLD"],
            key="order_margin_threshold"
        )
        operate_margin = st.number_input(
            "运营毛利率警戒值(%)",
            min_value=0.0, max_value=100.0, step=0.1,
            value=st.session_state.alert_config["OPERATE_MARGIN_RATE_THRESHOLD"],
            key="operate_margin_threshold"
        )
        unit_profit = st.number_input(
            "均单利润警戒值(元/单)",
            min_value=0.0, step=0.1,
            value=st.session_state.alert_config["UNIT_PROFIT_THRESHOLD"],
            key="unit_profit_threshold"
        )
    with col2:
        # 销量/客单价相关
        sales_quantity = st.number_input(
            "销售数量警戒值(单)",
            min_value=0, step=1,
            value=st.session_state.alert_config["SALES_QUANTITY_THRESHOLD"],
            key="sales_quantity_threshold"
        )
        unit_price = st.number_input(
            "客单价警戒值(元/单)",
            min_value=0.0, step=0.1,
            value=st.session_state.alert_config["UNIT_PRICE_THRESHOLD"],
            key="unit_price_threshold"
        )
    
    if st.button("💾 保存警戒值设置", key="save_alert_config"):
        new_config = {
            "ORDER_MARGIN_RATE_THRESHOLD": order_margin,
            "OPERATE_MARGIN_RATE_THRESHOLD": operate_margin,
            "SALES_QUANTITY_THRESHOLD": sales_quantity,
            "UNIT_PRICE_THRESHOLD": unit_price,
            "UNIT_PROFIT_THRESHOLD": unit_profit
        }
        st.session_state.alert_config.update(new_config)
        save_alert_config(new_config)  # 保存到文件
        st.success("✅ 警戒值设置保存成功！")
        st.rerun()
    
    # 显示当前生效的警戒值
    st.markdown("### 📌 当前生效的警戒值")
    alert_df = pd.DataFrame({
        "指标名称": [
            "订单毛利率警戒值", "运营毛利率警戒值", 
            "销售数量警戒值", "客单价警戒值", "均单利润警戒值"
        ],
        "当前值": [
            f"{st.session_state.alert_config['ORDER_MARGIN_RATE_THRESHOLD']}%",
            f"{st.session_state.alert_config['OPERATE_MARGIN_RATE_THRESHOLD']}%",
            f"{st.session_state.alert_config['SALES_QUANTITY_THRESHOLD']} 单",
            f"{st.session_state.alert_config['UNIT_PRICE_THRESHOLD']} 元/单",
            f"{st.session_state.alert_config['UNIT_PROFIT_THRESHOLD']} 元/单"
        ]
    })
    st.dataframe(alert_df, use_container_width=True)

# ===================== 工具函数 =====================
@st.cache_data(show_spinner="正在读取数据...")
def read_data(file) -> Optional[pd.DataFrame]:
    if file is None:
        return None
    try:
        if file.name.endswith('.csv'):
            return pd.read_csv(file, encoding=CONFIG["ENCODING"])
        else:
            return pd.read_excel(file)
    except Exception as e:
        st.warning(f"⚠️ 读取{file.name}失败：{str(e)}")
        return None

def calculate_margin_ratio(numerator, denominator) -> float | pd.Series:
    if isinstance(numerator, (int, float)) and isinstance(denominator, (int, float)):
        if denominator <= 0 or numerator == 0:
            return 0.0
        return round((numerator / denominator * 100), CONFIG["DECIMAL_PLACES"])
    elif isinstance(numerator, pd.Series) and isinstance(denominator, pd.Series):
        return np.where((denominator <= 0) | (numerator == 0), 0.0,
                        round((numerator / denominator * 100), CONFIG["DECIMAL_PLACES"]))
    return 0.0

# 根据警戒值标红的样式函数
def highlight_below_threshold(val, threshold):
    """如果数值低于阈值，返回红色背景样式"""
    try:
        if pd.isna(val) or val == 0:
            return ''
        val_num = float(val)
        if val_num < threshold:
            return 'background-color: #ffcccc; color: red; font-weight: bold'
    except:
        pass
    return ''

def highlight_threshold_values(df, alert_config):
    """为DataFrame应用阈值标红样式"""
    styled_df = df.style
    
    # 客单价标红
    if '客单价(元/单)' in df.columns:
        styled_df = styled_df.applymap(
            lambda x: highlight_below_threshold(x, alert_config["UNIT_PRICE_THRESHOLD"]),
            subset=['客单价(元/单)']
        )
    
    # 均单利润标红
    if '均单利润(元/单)' in df.columns:
        styled_df = styled_df.applymap(
            lambda x: highlight_below_threshold(x, alert_config["UNIT_PROFIT_THRESHOLD"]),
            subset=['均单利润(元/单)']
        )
    
    # 销售数量标红
    if '销售数量' in df.columns:
        styled_df = styled_df.applymap(
            lambda x: highlight_below_threshold(x, alert_config["SALES_QUANTITY_THRESHOLD"]),
            subset=['销售数量']
        )
    
    # 订单毛利率标红
    if '订单毛利率(%)' in df.columns:
        styled_df = styled_df.applymap(
            lambda x: highlight_below_threshold(x, alert_config["ORDER_MARGIN_RATE_THRESHOLD"]),
            subset=['订单毛利率(%)']
        )
    
    # 运营毛利率标红
    if '运营毛利率(%)' in df.columns:
        styled_df = styled_df.applymap(
            lambda x: highlight_below_threshold(x, alert_config["OPERATE_MARGIN_RATE_THRESHOLD"]),
            subset=['运营毛利率(%)']
        )
    
    return styled_df

# 计算单均指标
def calculate_sales_per_unit(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if '销售数量' not in df.columns:
        return df
    
    # 处理销售数量为0或空的情况，避免除零错误
    df['销售数量'] = df['销售数量'].replace(0, np.nan)
    sales_quantity = df['销售数量']
    
    # 单均订单毛利和单均商品成本
    df['单均订单毛利(元/单)'] = np.where(sales_quantity.notna(),
                                      round(df['订单毛利'] / sales_quantity, CONFIG["DECIMAL_PLACES"]), 0.0)
    df['单均商品成本(元/单)'] = np.where(sales_quantity.notna(),
                                     round(df['商品成本'] / sales_quantity, CONFIG["DECIMAL_PLACES"]), 0.0)
    
    # 客单价、均单利润
    df['客单价(元/单)'] = np.where(sales_quantity.notna(),
                                round(df['交易收入'] / sales_quantity, CONFIG["DECIMAL_PLACES"]), 0.0)
    df['均单利润(元/单)'] = np.where(sales_quantity.notna() & df['运营毛利'].notna(),
                                   round(df['运营毛利'] / sales_quantity, CONFIG["DECIMAL_PLACES"]), 0.0)
    
    # 毛利率、运营毛利率（补充到明细行）
    df['订单毛利率(%)'] = calculate_margin_ratio(df['订单毛利'], df['交易收入'])
    df['运营毛利率(%)'] = calculate_margin_ratio(df['运营毛利'], df['交易收入'])
    
    # 恢复销售数量为0（避免显示NaN）
    df['销售数量'] = df['销售数量'].fillna(0).astype(int)
    return df

def format_currency(value: float) -> str:
    return f"¥{value:,.{CONFIG['DECIMAL_PLACES']}f}"

def highlight_negative_values(val):
    if isinstance(val, (int, float)):
        if val > 0:
            return 'color: green; font-weight: bold'
        elif val < 0:
            return 'color: red; font-weight: bold'
    return ''

def to_excel(df_dict: dict) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in df_dict.items():
            safe_sheet_name = sheet_name[:30]
            df.to_excel(writer, sheet_name=safe_sheet_name, index=True)
    output.seek(0)
    return output.getvalue()

def generate_upload_template() -> bytes:
    template_cols = [
        "OA店铺名称", "销售员", "交易收入", "商品成本", "耗材成本", "人工成本",
        "头程运费", "退回运费", "消费者售后预留金额", "消费者售后释放金额",
        "店铺总计提金额", "罚款金额", "运营毛利", "销售数量"
    ]
    sample_data = [
        ["示例店铺A", "销售员甲", 10000.00, 6000.00, 300.00, 500.00, 800.00, 100.00, 200.00, 50.00, 0.00, 0.00, 2000.00, 200],
        ["示例店铺B", "销售员乙", 15000.00, 9000.00, 450.00, 750.00, 1200.00, 150.00, 300.00, 75.00, 0.00, 100.00, 3000.00, 300]
    ]
    template_df = pd.DataFrame(sample_data, columns=template_cols)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        template_df.to_excel(writer, sheet_name="数据模板", index=False)
        readme_df = pd.DataFrame({
            "字段名": template_cols,
            "是否必填": ["是"]*13 + ["可选"],
            "说明": [
                "店铺名称（唯一标识）", "负责该店铺的运营人员", "订单交易总收入（元）",
                "商品采购成本（元）", "包装耗材等成本（元）", "运营人工成本（元）",
                "发货头程运费（元）", "退货产生的运费（元）", "平台预留的售后金额（元）",
                "平台释放的售后金额（元）", "平台总计提金额（元）", "违规产生的罚款金额（元）",
                "店铺运营毛利（元）", "销售订单数量（用于计算客单价/均单利润等指标）"
            ]
        })
        readme_df.to_excel(writer, sheet_name="字段说明", index=False)
    output.seek(0)
    return output.getvalue()

# ===================== 核心计算函数 =====================
def calculate_metrics(df: pd.DataFrame, period_name: str) -> Optional[Dict]:
    if df is None or df.empty:
        return None

    metrics = {
        "周期": period_name,
        "店铺数量": len(df),
        "交易收入": 0.0,
        "订单毛利": 0.0,
        "订单毛利率": 0.0,
        "运营毛利": 0.0,
        "运营毛利率": 0.0,
        "商品成本": 0.0, "耗材成本": 0.0, "人工成本": 0.0,
        "头程运费": 0.0, "退回运费": 0.0, "罚款金额": 0.0,
        "店铺总计提金额": 0.0, "消费者售后预留金额": 0.0, "消费者售后释放金额": 0.0, "售后净额": 0.0,
        "商品成本占比": 0.0, "耗材成本占比": 0.0, "人工成本占比": 0.0,
        "头程运费占比": 0.0, "退回运费占比": 0.0, "罚款金额占比": 0.0,
        "店铺总计提占比": 0.0, "售后净额占比": 0.0,
        "销售数量总计": 0,
        "单均订单毛利(元/单)": 0.0,
        "单均商品成本(元/单)": 0.0,
        "客单价(元/单)": 0.0,
        "均单利润(元/单)": 0.0,
        "has_sales_quantity": '销售数量' in df.columns,
        "店铺数据": {},
        "sales_data": {}
    }

    metrics["交易收入"] = round(df.get('交易收入', 0).sum(), CONFIG["DECIMAL_PLACES"])
    metrics["罚款金额"] = round(df.get('罚款金额', 0).sum(), CONFIG["DECIMAL_PLACES"])
    metrics["运营毛利"] = round(df.get('运营毛利', 0).sum(), CONFIG["DECIMAL_PLACES"])

    metrics["商品成本"] = round(df.get('商品成本', 0).sum(), CONFIG["DECIMAL_PLACES"])
    metrics["耗材成本"] = round(df.get('耗材成本', 0).sum(), CONFIG["DECIMAL_PLACES"])
    metrics["人工成本"] = round(df.get('人工成本', 0).sum(), CONFIG["DECIMAL_PLACES"])
    metrics["头程运费"] = round(df.get('头程运费', 0).sum(), CONFIG["DECIMAL_PLACES"])
    metrics["退回运费"] = round(df.get('退回运费', 0).sum(), CONFIG["DECIMAL_PLACES"])
    metrics["店铺总计提金额"] = round(df.get('店铺总计提金额', 0).sum(), CONFIG["DECIMAL_PLACES"])
    metrics["消费者售后预留金额"] = round(df.get('消费者售后预留金额', 0).sum(), CONFIG["DECIMAL_PLACES"])
    metrics["消费者售后释放金额"] = round(df.get('消费者售后释放金额', 0).sum(), CONFIG["DECIMAL_PLACES"])
    metrics["售后净额"] = round(metrics["消费者售后预留金额"] - metrics["消费者售后释放金额"], CONFIG["DECIMAL_PLACES"])

    order_margin_components = [
        df.get('交易收入', 0),
        -df.get('头程运费', 0),
        -df.get('耗材成本', 0),
        -df.get('人工成本', 0),
        -df.get('商品成本', 0),
        -df.get('消费者售后预留金额', 0),
        df.get('消费者售后释放金额', 0)
    ]
    df['订单毛利'] = pd.concat(order_margin_components, axis=1).sum(axis=1)
    metrics["订单毛利"] = round(df['订单毛利'].sum(), CONFIG["DECIMAL_PLACES"])

    metrics["订单毛利率"] = calculate_margin_ratio(metrics["订单毛利"], metrics["交易收入"])
    metrics["运营毛利率"] = calculate_margin_ratio(metrics["运营毛利"], metrics["交易收入"])
    metrics["商品成本占比"] = calculate_margin_ratio(metrics["商品成本"], metrics["交易收入"])
    metrics["耗材成本占比"] = calculate_margin_ratio(metrics["耗材成本"], metrics["交易收入"])
    metrics["人工成本占比"] = calculate_margin_ratio(metrics["人工成本"], metrics["交易收入"])
    metrics["头程运费占比"] = calculate_margin_ratio(metrics["头程运费"], metrics["交易收入"])
    metrics["退回运费占比"] = calculate_margin_ratio(metrics["退回运费"], metrics["交易收入"])
    metrics["罚款金额占比"] = calculate_margin_ratio(metrics["罚款金额"], metrics["交易收入"])
    metrics["店铺总计提占比"] = calculate_margin_ratio(metrics["店铺总计提金额"], metrics["交易收入"])
    metrics["售后净额占比"] = calculate_margin_ratio(metrics["售后净额"], metrics["交易收入"])

    if metrics["has_sales_quantity"]:
        try:
            sales_series = df['销售数量'].fillna(0)
            metrics["销售数量总计"] = int(sales_series.sum())
            if metrics["销售数量总计"] > 0:
                metrics["单均订单毛利(元/单)"] = round(metrics["订单毛利"] / metrics["销售数量总计"], CONFIG["DECIMAL_PLACES"])
                metrics["单均商品成本(元/单)"] = round(metrics["商品成本"] / metrics["销售数量总计"], CONFIG["DECIMAL_PLACES"])
                metrics["客单价(元/单)"] = round(metrics["交易收入"] / metrics["销售数量总计"], CONFIG["DECIMAL_PLACES"])
                metrics["均单利润(元/单)"] = round(metrics["运营毛利"] / metrics["销售数量总计"], CONFIG["DECIMAL_PLACES"])
        except Exception as e:
            metrics["销售数量总计"] = 0
            st.warning(f"⚠️ 销售数量字段格式异常，已跳过销量相关指标计算：{str(e)}")

    if 'OA店铺名称' in df.columns:
        df = calculate_sales_per_unit(df)
        agg_cols = {
            '交易收入': lambda x: round(x.sum(), CONFIG["DECIMAL_PLACES"]),
            '订单毛利': lambda x: round(x.sum(), CONFIG["DECIMAL_PLACES"]),
            '运营毛利': lambda x: round(x.sum(), CONFIG["DECIMAL_PLACES"]) if '运营毛利' in df.columns else 0,
            '商品成本': lambda x: round(x.sum(), CONFIG["DECIMAL_PLACES"]),
            '耗材成本': lambda x: round(x.sum(), CONFIG["DECIMAL_PLACES"]),
            '人工成本': lambda x: round(x.sum(), CONFIG["DECIMAL_PLACES"]),
            '头程运费': lambda x: round(x.sum(), CONFIG["DECIMAL_PLACES"]),
            '退回运费': lambda x: round(x.sum(), CONFIG["DECIMAL_PLACES"]),
            '罚款金额': lambda x: round(x.sum(), CONFIG["DECIMAL_PLACES"]),
            '店铺总计提金额': lambda x: round(x.sum(), CONFIG["DECIMAL_PLACES"]),
            '消费者售后预留金额': lambda x: round(x.sum(), CONFIG["DECIMAL_PLACES"]),
            '消费者售后释放金额': lambda x: round(x.sum(), CONFIG["DECIMAL_PLACES"]),
            'OA店铺名称': 'count'
        }
        if metrics["has_sales_quantity"]:
            agg_cols.update({
                '销售数量': 'sum',
                '单均订单毛利(元/单)': lambda x: round(x.mean(), CONFIG["DECIMAL_PLACES"]),
                '单均商品成本(元/单)': lambda x: round(x.mean(), CONFIG["DECIMAL_PLACES"]),
                '客单价(元/单)': lambda x: round(x.mean(), CONFIG["DECIMAL_PLACES"]),
                '均单利润(元/单)': lambda x: round(x.mean(), CONFIG["DECIMAL_PLACES"]),
                '订单毛利率(%)': lambda x: round(x.mean(), CONFIG["DECIMAL_PLACES"]),
                '运营毛利率(%)': lambda x: round(x.mean(), CONFIG["DECIMAL_PLACES"])
            })
        shop_agg = df.groupby('OA店铺名称').agg(agg_cols).rename(columns={'OA店铺名称': '店铺数量'})
        shop_agg['店铺数量'] = shop_agg['店铺数量'].astype(int)
        if metrics["has_sales_quantity"]:
            shop_agg['销售数量'] = shop_agg['销售数量'].fillna(0).astype(int)
        # 补充毛利率计算
        shop_agg['订单毛利率(%)'] = calculate_margin_ratio(shop_agg['订单毛利'], shop_agg['交易收入'])
        shop_agg['运营毛利率(%)'] = calculate_margin_ratio(shop_agg['运营毛利'], shop_agg['交易收入'])
        shop_agg['商品成本占比(%)'] = calculate_margin_ratio(shop_agg['商品成本'], shop_agg['交易收入'])
        shop_agg['耗材成本占比(%)'] = calculate_margin_ratio(shop_agg['耗材成本'], shop_agg['交易收入'])
        shop_agg['人工成本占比(%)'] = calculate_margin_ratio(shop_agg['人工成本'], shop_agg['交易收入'])
        shop_agg['头程运费占比(%)'] = calculate_margin_ratio(shop_agg['头程运费'], shop_agg['交易收入'])
        shop_agg['退回运费占比(%)'] = calculate_margin_ratio(shop_agg['退回运费'], shop_agg['交易收入'])
        shop_agg['罚款金额占比(%)'] = calculate_margin_ratio(shop_agg['罚款金额'], shop_agg['交易收入'])
        shop_agg['店铺总计提占比(%)'] = calculate_margin_ratio(shop_agg['店铺总计提金额'], shop_agg['交易收入'])
        shop_agg['售后净额'] = shop_agg['消费者售后预留金额'] - shop_agg['消费者售后释放金额']
        shop_agg['售后净额占比(%)'] = calculate_margin_ratio(shop_agg['售后净额'], shop_agg['交易收入'])
        metrics["店铺数据"] = shop_agg.to_dict('index')

    if '销售员' in df.columns:
        df = calculate_sales_per_unit(df)
        agg_cols = {
            '销售员': 'count',
            '交易收入': lambda x: round(x.sum(), CONFIG["DECIMAL_PLACES"]),
            '订单毛利': lambda x: round(x.sum(), CONFIG["DECIMAL_PLACES"]),
            '运营毛利': lambda x: round(x.sum(), CONFIG["DECIMAL_PLACES"]) if '运营毛利' in df.columns else 0,
            '商品成本': lambda x: round(x.sum(), CONFIG["DECIMAL_PLACES"]),
            '耗材成本': lambda x: round(x.sum(), CONFIG["DECIMAL_PLACES"]),
            '人工成本': lambda x: round(x.sum(), CONFIG["DECIMAL_PLACES"]),
            '头程运费': lambda x: round(x.sum(), CONFIG["DECIMAL_PLACES"]),
            '退回运费': lambda x: round(x.sum(), CONFIG["DECIMAL_PLACES"]),
            '罚款金额': lambda x: round(x.sum(), CONFIG["DECIMAL_PLACES"]),
            '店铺总计提金额': lambda x: round(x.sum(), CONFIG["DECIMAL_PLACES"]),
            '消费者售后预留金额': lambda x: round(x.sum(), CONFIG["DECIMAL_PLACES"]),
            '消费者售后释放金额': lambda x: round(x.sum(), CONFIG["DECIMAL_PLACES"])
        }
        if metrics["has_sales_quantity"]:
            agg_cols.update({
                '销售数量': 'sum',
                '单均订单毛利(元/单)': lambda x: round(x.mean(), CONFIG["DECIMAL_PLACES"]),
                '单均商品成本(元/单)': lambda x: round(x.mean(), CONFIG["DECIMAL_PLACES"]),
                '客单价(元/单)': lambda x: round(x.mean(), CONFIG["DECIMAL_PLACES"]),
                '均单利润(元/单)': lambda x: round(x.mean(), CONFIG["DECIMAL_PLACES"]),
                '订单毛利率(%)': lambda x: round(x.mean(), CONFIG["DECIMAL_PLACES"]),
                '运营毛利率(%)': lambda x: round(x.mean(), CONFIG["DECIMAL_PLACES"])
            })
        sales_agg = df.groupby('销售员').agg(agg_cols).rename(columns={'销售员': '店铺数量'})
        sales_agg['店铺数量'] = sales_agg['店铺数量'].astype(int)
        if metrics["has_sales_quantity"]:
            sales_agg['销售数量'] = sales_agg['销售数量'].fillna(0).astype(int)
        # 补充毛利率计算
        sales_agg['订单毛利率(%)'] = calculate_margin_ratio(sales_agg['订单毛利'], sales_agg['交易收入'])
        sales_agg['运营毛利率(%)'] = calculate_margin_ratio(sales_agg['运营毛利'], sales_agg['交易收入'])
        sales_agg['罚款占收入比(%)'] = calculate_margin_ratio(sales_agg['罚款金额'], sales_agg['交易收入'])
        sales_agg['商品成本占比(%)'] = calculate_margin_ratio(sales_agg['商品成本'], sales_agg['交易收入'])
        sales_agg['耗材成本占比(%)'] = calculate_margin_ratio(sales_agg['耗材成本'], sales_agg['交易收入'])
        sales_agg['人工成本占比(%)'] = calculate_margin_ratio(sales_agg['人工成本'], sales_agg['交易收入'])
        sales_agg['头程运费占比(%)'] = calculate_margin_ratio(sales_agg['头程运费'], sales_agg['交易收入'])
        sales_agg['退回运费占比(%)'] = calculate_margin_ratio(sales_agg['退回运费'], sales_agg['交易收入'])
        sales_agg['店铺总计提占比(%)'] = calculate_margin_ratio(sales_agg['店铺总计提金额'], sales_agg['交易收入'])
        sales_agg['售后净额'] = sales_agg['消费者售后预留金额'] - sales_agg['消费者售后释放金额']
        sales_agg['售后净额占比(%)'] = calculate_margin_ratio(sales_agg['售后净额'], sales_agg['交易收入'])
        metrics["sales_data"] = sales_agg.to_dict('index')

    return metrics

# ===================== 可视化函数 =====================
def plot_margin_chart(df: pd.DataFrame, y_col: str, title: str, threshold_key: str, selected_items=None):
    threshold = st.session_state.alert_config[threshold_key]
    if selected_items:
        df = df.loc[selected_items]
    fig = px.bar(df, x=df.index, y=y_col, title=title)
    fig.update_traces(text=df[y_col].apply(lambda x: f"{x:.2f}%"), textposition='outside',
                      marker_color=['#d62728' if x < threshold else '#2ca02c' for x in df[y_col]])
    fig.add_hline(y=threshold, line_dash="dash", line_color="orange", annotation_text=f"达标线({threshold}%)")
    return fig

def plot_cost_ratio_chart(df: pd.DataFrame, title: str, selected_items=None):
    if selected_items:
        df = df.loc[selected_items]
    cost_cols = [c for c in ['商品成本占比(%)', '耗材成本占比(%)', '人工成本占比(%)',
                             '头程运费占比(%)', '退回运费占比(%)', '罚款金额占比(%)',
                             '店铺总计提占比(%)', '售后净额占比(%)'] if c in df.columns]
    fig = go.Figure()
    colors = ['#FF6B6B','#4ECDC4','#45B7D1','#96CEB4','#FFEAA7','#DDA0DD','#98D8C8','#F7DC6F']
    for i, c in enumerate(cost_cols):
        fig.add_trace(go.Bar(x=df.index, y=df[c], name=c.replace('占比(%)',''), marker_color=colors[i%len(colors)],
                            text=df[c].apply(lambda x: f"{x:.2f}%"),
                            textposition='inside'))
    fig.update_layout(title=title, barmode='stack', yaxis_title='占比(%)', height=600)
    return fig

def plot_sales_quantity_chart(df: pd.DataFrame, title: str, selected_items=None):
    threshold = st.session_state.alert_config["SALES_QUANTITY_THRESHOLD"]
    if selected_items:
        df = df.loc[selected_items]
    df = df.sort_values('销售数量', ascending=False)
    fig = px.bar(df, x=df.index, y='销售数量', title=title, text_auto=True)
    fig.update_traces(text=df['销售数量'].apply(lambda x: f"{x:,}"), textposition='outside',
                      marker_color=['#d62728' if x < threshold else '#2ca02c' for x in df['销售数量']])
    fig.add_hline(y=threshold, line_dash="dash", line_color="orange", annotation_text=f"达标线({threshold}单)")
    fig.update_layout(yaxis_title='销售数量（单）', height=500)
    return fig

def plot_unit_metrics_chart(df: pd.DataFrame, title: str, selected_items=None):
    if selected_items:
        df = df.loc[selected_items]
    unit_price_threshold = st.session_state.alert_config["UNIT_PRICE_THRESHOLD"]
    unit_profit_threshold = st.session_state.alert_config["UNIT_PROFIT_THRESHOLD"]
    
    fig = go.Figure()
    # 客单价
    fig.add_trace(go.Bar(x=df.index, y=df['客单价(元/单)'], name=f'客单价 (警戒值:{unit_price_threshold}元)', 
                         marker_color=['#d62728' if x < unit_price_threshold else '#2ca02c' for x in df['客单价(元/单)']],
                         text=df['客单价(元/单)'].apply(lambda x: f"¥{x:.2f}"), textposition='outside'))
    # 均单利润
    fig.add_trace(go.Bar(x=df.index, y=df['均单利润(元/单)'], name=f'均单利润 (警戒值:{unit_profit_threshold}元)',
                         marker_color=['#FF8C00' if x < unit_profit_threshold else '#006400' for x in df['均单利润(元/单)']],
                         text=df['均单利润(元/单)'].apply(lambda x: f"¥{x:.2f}"), textposition='outside'))
    fig.update_layout(title=title, barmode='group', height=600, yaxis_title='金额(元/单)')
    fig.add_hline(y=unit_price_threshold, line_dash="dash", line_color="red", annotation_text=f"客单价警戒值({unit_price_threshold}元)")
    fig.add_hline(y=unit_profit_threshold, line_dash="dash", line_color="orange", annotation_text=f"均单利润警戒值({unit_profit_threshold}元)")
    return fig

def plot_sales_unit_metrics_chart(df: pd.DataFrame, title: str, selected_items=None):
    if selected_items:
        df = df.loc[selected_items]
    fig = go.Figure()
    fig.add_trace(go.Bar(x=df.index, y=df['单均订单毛利(元/单)'], name='单均订单毛利', marker_color='#FFD700',
                         text=df['单均订单毛利(元/单)'].apply(lambda x: f"¥{x:.2f}"), textposition='outside'))
    fig.add_trace(go.Bar(x=df.index, y=df['单均商品成本(元/单)'], name='单均商品成本', marker_color='#DC143C',
                         text=df['单均商品成本(元/单)'].apply(lambda x: f"¥{x:.2f}"), textposition='outside'))
    if '客单价(元/单)' in df.columns:
        fig.add_trace(go.Bar(x=df.index, y=df['客单价(元/单)'], name='客单价', marker_color='#8A2BE2',
                             text=df['客单价(元/单)'].apply(lambda x: f"¥{x:.2f}"), textposition='outside'))
    if '均单利润(元/单)' in df.columns:
        fig.add_trace(go.Bar(x=df.index, y=df['均单利润(元/单)'], name='均单利润', marker_color='#FF6347',
                             text=df['均单利润(元/单)'].apply(lambda x: f"¥{x:.2f}"), textposition='outside'))
    fig.update_layout(title=title, barmode='group', height=600)
    return fig

# ===================== 页面渲染函数 =====================
def render_monthly_analysis(metrics: Dict, df: pd.DataFrame):
    st.subheader(f"📈 {metrics['周期']} 核心数据")
    c1,c2,c3 = st.columns(3)
    c4,c5,c6 = st.columns(3)

    with c1:
        st.metric("店铺数量", int(metrics["店铺数量"]))
        st.metric("交易收入", format_currency(metrics["交易收入"]))
        st.metric("商品成本占比", f"{metrics['商品成本占比']:.2f}%")
        st.metric("人工成本占比", f"{metrics['人工成本占比']:.2f}%")
    with c2:
        st.metric("订单毛利", format_currency(metrics["订单毛利"]))
        st.metric("运营毛利", format_currency(metrics["运营毛利"]))
        st.metric("耗材成本占比", f"{metrics['耗材成本占比']:.2f}%")
        st.metric("头程运费占比", f"{metrics['头程运费占比']:.2f}%")
    with c3:
        o_rate = metrics["订单毛利率"]
        op_rate = metrics["运营毛利率"]
        order_threshold = st.session_state.alert_config["ORDER_MARGIN_RATE_THRESHOLD"]
        operate_threshold = st.session_state.alert_config["OPERATE_MARGIN_RATE_THRESHOLD"]
        
        st.metric("订单毛利率", f"{o_rate:.2f}%", 
                  delta=f"＜{order_threshold}% 警示" if o_rate<order_threshold else f"≥{order_threshold}% 合格", 
                  delta_color="inverse" if o_rate<order_threshold else "normal")
        st.metric("运营毛利率", f"{op_rate:.2f}%", 
                  delta=f"＜{operate_threshold}% 警示" if op_rate<operate_threshold else f"≥{operate_threshold}% 合格", 
                  delta_color="inverse" if op_rate<operate_threshold else "normal")
        st.metric("退回运费占比", f"{metrics['退回运费占比']:.2f}%")
        st.metric("罚款占比", f"{metrics['罚款金额占比']:.2f}%")
    
    with c4:
        if metrics["has_sales_quantity"]:
            sales_threshold = st.session_state.alert_config["SALES_QUANTITY_THRESHOLD"]
            st.metric("销售数量总计", f"{metrics['销售数量总计']:,} 单",
                      delta=f"＜{sales_threshold} 单 偏低" if metrics['销售数量总计']<sales_threshold else f"≥{sales_threshold} 单 正常",
                      delta_color="inverse" if metrics['销售数量总计']<sales_threshold else "normal")
        else:
            st.metric("提示", "无销售数量数据")
            st.metric("-", "-")
    with c5:
        if metrics["has_sales_quantity"]:
            unit_price_threshold = st.session_state.alert_config["UNIT_PRICE_THRESHOLD"]
            unit_profit_threshold = st.session_state.alert_config["UNIT_PROFIT_THRESHOLD"]
            
            st.metric("客单价", f"{metrics['客单价(元/单)']:.2f} 元/单",
                      delta=f"＜{unit_price_threshold} 元 偏低" if metrics['客单价(元/单)']<unit_price_threshold else f"≥{unit_price_threshold} 元 正常",
                      delta_color="inverse" if metrics['客单价(元/单)']<unit_price_threshold else "normal")
            st.metric("均单利润", f"{metrics['均单利润(元/单)']:.2f} 元/单",
                      delta=f"＜{unit_profit_threshold} 元 偏低" if metrics['均单利润(元/单)']<unit_profit_threshold else f"≥{unit_profit_threshold} 元 正常",
                      delta_color="inverse" if metrics['均单利润(元/单)']<unit_profit_threshold else "normal")
        else:
            st.metric("-", "-")
            st.metric("-", "-")
    with c6:
        st.metric("售后净额", format_currency(metrics["售后净额"]))
        st.metric("售后净额占比", f"{metrics['售后净额占比']:.2f}%")
        st.metric("罚款金额", format_currency(metrics["罚款金额"]))

    st.subheader("📋 原始数据预览")
    st.dataframe(df.head(10), use_container_width=True)

    if metrics["店铺数据"]:
        st.subheader("🏪 店铺分析")
        shop_df = pd.DataFrame(metrics["店铺数据"]).T
        shop_df['店铺数量'] = shop_df['店铺数量'].astype(int)
        if metrics["has_sales_quantity"]:
            shop_df['销售数量'] = shop_df['销售数量'].fillna(0).astype(int)
        shops = st.multiselect("选择店铺", list(shop_df.index), default=list(shop_df.index)[:5])
        if shops:
            styled_shop_df = highlight_threshold_values(shop_df.loc[shops], st.session_state.alert_config)
            st.dataframe(styled_shop_df, use_container_width=True)
            
            st.plotly_chart(plot_margin_chart(shop_df, '订单毛利率(%)', '店铺订单毛利率', 'ORDER_MARGIN_RATE_THRESHOLD', shops), use_container_width=True)
            st.plotly_chart(plot_cost_ratio_chart(shop_df, '店铺成本占比', shops), use_container_width=True)
            if metrics["has_sales_quantity"]:
                st.plotly_chart(plot_sales_quantity_chart(shop_df, '店铺销售数量排名', shops), use_container_width=True)
                st.plotly_chart(plot_sales_unit_metrics_chart(shop_df, '店铺单均/客单价/均单利润指标', shops), use_container_width=True)
                st.plotly_chart(plot_unit_metrics_chart(shop_df, '店铺客单价&均单利润对比', shops), use_container_width=True)

    if metrics["sales_data"]:
        st.subheader("👨‍💼 销售员分析")
        sales_df = pd.DataFrame(metrics["sales_data"]).T
        sales_df['店铺数量'] = sales_df['店铺数量'].astype(int)
        if metrics["has_sales_quantity"]:
            sales_df['销售数量'] = sales_df['销售数量'].fillna(0).astype(int)
        sales = st.multiselect("选择销售员", list(sales_df.index), default=list(sales_df.index)[:5])
        if sales:
            styled_sales_df = highlight_threshold_values(sales_df.loc[sales], st.session_state.alert_config)
            st.dataframe(styled_sales_df, use_container_width=True)
            
            st.plotly_chart(plot_cost_ratio_chart(sales_df, '销售员成本占比', sales), use_container_width=True)
            if metrics["has_sales_quantity"]:
                st.plotly_chart(plot_sales_quantity_chart(sales_df, '销售员销量排名', sales), use_container_width=True)
                st.plotly_chart(plot_sales_unit_metrics_chart(sales_df, '销售员单均/客单价/均单利润指标', sales), use_container_width=True)
                st.plotly_chart(plot_unit_metrics_chart(sales_df, '销售员客单价&均单利润对比', sales), use_container_width=True)

        st.subheader("📌 运营建议")
        for name, row in sales_df.iterrows():
            with st.expander(f"🧑‍💼 {name}", expanded=False):
                order_threshold = st.session_state.alert_config["ORDER_MARGIN_RATE_THRESHOLD"]
                operate_threshold = st.session_state.alert_config["OPERATE_MARGIN_RATE_THRESHOLD"]
                sales_threshold = st.session_state.alert_config["SALES_QUANTITY_THRESHOLD"]
                unit_price_threshold = st.session_state.alert_config["UNIT_PRICE_THRESHOLD"]
                unit_profit_threshold = st.session_state.alert_config["UNIT_PROFIT_THRESHOLD"]
                
                o_rate = row.get('订单毛利率(%)',0)
                op_rate = row.get('运营毛利率(%)',0)
                fine = row.get('罚款占收入比(%)',0)
                
                if o_rate < order_threshold: 
                    st.warning(f"订单毛利率{o_rate:.2f}% 不达标（标准{order_threshold}%）")
                else: 
                    st.success(f"订单毛利率{o_rate:.2f}% 达标（标准{order_threshold}%）")
                
                if op_rate < operate_threshold: 
                    st.warning(f"运营毛利率{op_rate:.2f}% 不达标（标准{operate_threshold}%）")
                else: 
                    st.success(f"运营毛利率{op_rate:.2f}% 达标（标准{operate_threshold}%）")
                
                if fine>0: 
                    st.info(f"罚款占收入比{fine:.2f}%，需减少违规订单")
                
                if metrics["has_sales_quantity"]:
                    qty = row.get('销售数量',0)
                    unit_price = row.get('客单价(元/单)',0)
                    unit_profit = row.get('均单利润(元/单)',0)
                    
                    if qty < sales_threshold: 
                        st.warning(f"销量{qty} 偏低（建议≥{sales_threshold}单）")
                    else: 
                        st.success(f"销量{qty} 正常（≥{sales_threshold}单）")
                    
                    if unit_price < unit_price_threshold: 
                        st.warning(f"客单价{unit_price:.2f}元 偏低（建议≥{unit_price_threshold}元）")
                    else: 
                        st.success(f"客单价{unit_price:.2f}元 正常（≥{unit_price_threshold}元）")
                    
                    if unit_profit < unit_profit_threshold: 
                        st.warning(f"均单利润{unit_profit:.2f}元 偏低（建议≥{unit_profit_threshold}元）")
                    else: 
                        st.success(f"均单利润{unit_profit:.2f}元 正常（≥{unit_profit_threshold}元）")

def render_double_month_analysis(curr: Dict, last: Dict):
    st.subheader("🔍 双月详细对比分析")
    base_metrics = [
        "店铺数量", "交易收入", "订单毛利", "订单毛利率", "运营毛利", "运营毛利率",
        "商品成本", "商品成本占比", "人工成本", "人工成本占比", 
        "头程运费", "头程运费占比", "退回运费", "退回运费占比",
        "罚款金额", "罚款金额占比", "售后净额", "售后净额占比"
    ]
    compare_data = {
        "指标名称": [
            "店铺数量", "交易收入(元)", "订单毛利(元)", "订单毛利率(%)", 
            "运营毛利(元)", "运营毛利率(%)", "罚款金额(元)", "罚款占收入比(%)",
            "商品成本(元)", "商品成本占比(%)", "耗材成本(元)", "耗材成本占比(%)",
            "人工成本(元)", "人工成本占比(%)", "头程运费(元)", "头程运费占比(%)",
            "退回运费(元)", "退回运费占比(%)", "店铺总计提(元)", "店铺总计提占比(%)",
            "售后净额(元)", "售后净额占比(%)"
        ],
        "上月数值": [
            int(last["店铺数量"]), last["交易收入"], last["订单毛利"], last["订单毛利率"],
            last["运营毛利"], last["运营毛利率"], last["罚款金额"],
            calculate_margin_ratio(last["罚款金额"], last["交易收入"]),
            last["商品成本"], last["商品成本占比"], last["耗材成本"], last["耗材成本占比"],
            last["人工成本"], last["人工成本占比"], last["头程运费"], last["头程运费占比"],
            last["退回运费"], last["退回运费占比"], last["店铺总计提金额"], last["店铺总计提占比"],
            last["售后净额"], last["售后净额占比"]
        ],
        "本月数值": [
            int(curr["店铺数量"]), curr["交易收入"], curr["订单毛利"], curr["订单毛利率"],
            curr["运营毛利"], curr["运营毛利率"], curr["罚款金额"],
            calculate_margin_ratio(curr["罚款金额"], curr["交易收入"]),
            curr["商品成本"], curr["商品成本占比"], curr["耗材成本"], curr["耗材成本占比"],
            curr["人工成本"], curr["人工成本占比"], curr["头程运费"], curr["头程运费占比"],
            curr["退回运费"], curr["退回运费占比"], curr["店铺总计提金额"], curr["店铺总计提占比"],
            curr["售后净额"], curr["售后净额占比"]
        ]
    }
    # 新增客单价/均单利润对比
    if curr["has_sales_quantity"] and last["has_sales_quantity"]:
        compare_data["指标名称"].extend([
            "销售数量(单)", "客单价(元/单)", "均单利润(元/单)",
            "单均订单毛利(元/单)"
        ])
        compare_data["上月数值"].extend([
            last["销售数量总计"], last["客单价(元/单)"], last["均单利润(元/单)"],
            last["单均订单毛利(元/单)"]
        ])
        compare_data["本月数值"].extend([
            curr["销售数量总计"], curr["客单价(元/单)"], curr["均单利润(元/单)"],
            curr["单均订单毛利(元/单)"]
        ])
    compare_df = pd.DataFrame(compare_data)
    compare_df["绝对差异"] = compare_df["本月数值"] - compare_df["上月数值"]
    compare_df["相对差异(%)"] = round(
        (compare_df["绝对差异"] / compare_df["上月数值"] * 100).replace([np.inf, -np.inf], 0),
        CONFIG["DECIMAL_PLACES"]
    )
    st.dataframe(
        compare_df.style.applymap(highlight_negative_values, subset=["绝对差异", "相对差异(%)"]),
        use_container_width=True
    )
    st.markdown("### 关键金额指标双月对比")
    fig_amount = go.Figure()
    metrics_list = ["交易收入", "订单毛利", "运营毛利", "商品成本", "人工成本", "头程运费"]
    for metric in metrics_list:
        fig_amount.add_trace(go.Bar(
            x=['上月', '本月'],
            y=[last[metric], curr[metric]],
            name=metric,
            text=[format_currency(last[metric]), format_currency(curr[metric])],
            textposition='outside'
        ))
    fig_amount.update_layout(
        title='交易收入/毛利/核心成本 双月对比',
        xaxis_title='周期',
        yaxis_title='金额(元)',
        barmode='group'
    )
    st.plotly_chart(fig_amount, use_container_width=True)
    if curr["has_sales_quantity"] and last["has_sales_quantity"]:
        st.markdown("### 销量&客单价/均单利润指标双月对比")
        fig_sales = go.Figure()
        fig_sales.add_trace(go.Bar(x=['上月','本月'], y=[last["销售数量总计"], curr["销售数量总计"]], 
                                  name='销售数量', 
                                  text=[f"{last['销售数量总计']:,}", f"{curr['销售数量总计']:,}"],
                                  textposition='outside'))
        fig_sales.add_trace(go.Bar(x=['上月','本月'], y=[last["客单价(元/单)"], curr["客单价(元/单)"]], 
                                  name='客单价',
                                  text=[f"¥{last['客单价(元/单)']:.2f}", f"¥{curr['客单价(元/单)']:.2f}"],
                                  textposition='outside'))
        fig_sales.add_trace(go.Bar(x=['上月','本月'], y=[last["均单利润(元/单)"], curr["均单利润(元/单)"]], 
                                  name='均单利润',
                                  text=[f"¥{last['均单利润(元/单)']:.2f}", f"¥{curr['均单利润(元/单)']:.2f}"],
                                  textposition='outside'))
        fig_sales.update_layout(title='销量 & 客单价 & 均单利润对比', barmode='group')
        st.plotly_chart(fig_sales, use_container_width=True)
    st.markdown("### 双月差异总结")
    income_diff = curr["交易收入"] - last["交易收入"]
    income_rate = calculate_margin_ratio(income_diff, last["交易收入"])
    if income_diff > 0:
        st.success(f"✅ 交易收入环比增长 {format_currency(income_diff)}（{income_rate:.2f}%）")
    elif income_diff < 0:
        st.error(f"❌ 交易收入环比下降 {format_currency(abs(income_diff))}（{abs(income_rate):.2f}%）")
    else:
        st.info(f"➡️ 交易收入环比无变化")
    order_diff = curr["订单毛利"] - last["订单毛利"]
    order_rate = calculate_margin_ratio(order_diff, last["订单毛利"])
    if order_diff > 0:
        st.success(f"✅ 订单毛利环比增长 {format_currency(order_diff)}（{order_rate:.2f}%）")
    elif order_diff < 0:
        st.error(f"❌ 订单毛利环比下降 {format_currency(abs(order_diff))}（{abs(order_rate):.2f}%）")
    else:
        st.info(f"➡️ 订单毛利环比无变化")
    if curr["has_sales_quantity"] and last["has_sales_quantity"]:
        qty_diff = curr["销售数量总计"] - last["销售数量总计"]
        qty_rate = calculate_margin_ratio(qty_diff, last["销售数量总计"])
        if qty_diff > 0:
            st.success(f"✅ 销售数量环比增长 {qty_diff:,} 单（{qty_rate:.2f}%）")
        elif qty_diff < 0:
            st.error(f"❌ 销售数量环比下降 {abs(qty_diff):,} 单（{abs(qty_rate):.2f}%）")
        else:
            st.info(f"➡️ 销售数量环比无变化")
        
        unit_price_diff = curr["客单价(元/单)"] - last["客单价(元/单)"]
        unit_price_rate = calculate_margin_ratio(unit_price_diff, last["客单价(元/单)"])
        if unit_price_diff > 0:
            st.success(f"✅ 客单价环比增长 {unit_price_diff:.2f} 元（{unit_price_rate:.2f}%）")
        elif unit_price_diff < 0:
            st.error(f"❌ 客单价环比下降 {abs(unit_price_diff):.2f} 元（{abs(unit_price_rate):.2f}%）")
        else:
            st.info(f"➡️ 客单价环比无变化")
        
        unit_profit_diff = curr["均单利润(元/单)"] - last["均单利润(元/单)"]
        unit_profit_rate = calculate_margin_ratio(unit_profit_diff, last["均单利润(元/单)"])
        if unit_profit_diff > 0:
            st.success(f"✅ 均单利润环比增长 {unit_profit_diff:.2f} 元（{unit_profit_rate:.2f}%）")
        elif unit_profit_diff < 0:
            st.error(f"❌ 均单利润环比下降 {abs(unit_profit_diff):.2f} 元（{abs(unit_profit_rate):.2f}%）")
        else:
            st.info(f"➡️ 均单利润环比无变化")
    return compare_df

def render_shop_margin_ranking(metrics_current: Dict, metrics_last: Dict) -> Optional[pd.DataFrame]:
    st.subheader("🏪 店铺运营毛利增长排名")
    shop_curr = pd.DataFrame(metrics_current["店铺数据"]).T
    shop_last = pd.DataFrame(metrics_last["店铺数据"]).T
    common_shops = list(set(shop_curr.index) & set(shop_last.index))
    if not common_shops:
        st.info("⚠️ 两个月无共同店铺，无法生成排名")
        return None
    rank_df = pd.DataFrame(index=common_shops)
    rank_df.index.name = "店铺名称"
    rank_df["运营毛利_上月(元)"] = shop_last["运营毛利"].round(CONFIG["DECIMAL_PLACES"])
    rank_df["运营毛利_本月(元)"] = shop_curr["运营毛利"].round(CONFIG["DECIMAL_PLACES"])
    rank_df["运营毛利_差异(元)"] = (rank_df["运营毛利_本月(元)"] - rank_df["运营毛利_上月(元)"]).round(CONFIG["DECIMAL_PLACES"])
    rank_df["运营毛利_环比(%)"] = calculate_margin_ratio(rank_df["运营毛利_差异(元)"], rank_df["运营毛利_上月(元)"])
    # 新增毛利率、客单价、均单利润排名字段
    rank_df["订单毛利率_上月(%)"] = shop_last["订单毛利率(%)"].round(CONFIG["DECIMAL_PLACES"])
    rank_df["订单毛利率_本月(%)"] = shop_curr["订单毛利率(%)"].round(CONFIG["DECIMAL_PLACES"])
    rank_df["运营毛利率_上月(%)"] = shop_last["运营毛利率(%)"].round(CONFIG["DECIMAL_PLACES"])
    rank_df["运营毛利率_本月(%)"] = shop_curr["运营毛利率(%)"].round(CONFIG["DECIMAL_PLACES"])
    
    if metrics_current["has_sales_quantity"] and metrics_last["has_sales_quantity"]:
        rank_df["商品成本占比_上月(%)"] = shop_last["商品成本占比(%)"].round(CONFIG["DECIMAL_PLACES"])
        rank_df["商品成本占比_本月(%)"] = shop_curr["商品成本占比(%)"].round(CONFIG["DECIMAL_PLACES"])
        rank_df["人工成本占比_上月(%)"] = shop_last["人工成本占比(%)"].round(CONFIG["DECIMAL_PLACES"])
        rank_df["人工成本占比_本月(%)"] = shop_curr["人工成本占比(%)"].round(CONFIG["DECIMAL_PLACES"])
        rank_df["销售数量_上月"] = shop_last["销售数量"].fillna(0).astype(int)
        rank_df["销售数量_本月"] = shop_curr["销售数量"].fillna(0).astype(int)
        rank_df["销量差异"] = rank_df["销售数量_本月"] - rank_df["销售数量_上月"]
        # 客单价、均单利润
        rank_df["客单价_上月(元/单)"] = shop_last["客单价(元/单)"].round(CONFIG["DECIMAL_PLACES"])
        rank_df["客单价_本月(元/单)"] = shop_curr["客单价(元/单)"].round(CONFIG["DECIMAL_PLACES"])
        rank_df["均单利润_上月(元/单)"] = shop_last["均单利润(元/单)"].round(CONFIG["DECIMAL_PLACES"])
        rank_df["均单利润_本月(元/单)"] = shop_curr["均单利润(元/单)"].round(CONFIG["DECIMAL_PLACES"])
    rank_df = rank_df.sort_values("运营毛利_差异(元)", ascending=False)
    rank_df.insert(0, "排名", range(1, len(rank_df)+1))
    
    # 应用阈值标红样式
    styled_rank_df = highlight_threshold_values(rank_df, st.session_state.alert_config)
    st.dataframe(
        styled_rank_df,
        use_container_width=True
    )
    if len(rank_df) > 0:
        top_shop = rank_df.index[0]
        top_growth = rank_df["运营毛利_差异(元)"].iloc[0]
        st.success(f"🥇 运营毛利增长TOP1店铺：{top_shop}（增长 {format_currency(top_growth)}）")
    return rank_df

def render_sales_margin_ranking(metrics_current: Dict, metrics_last: Dict) -> Optional[pd.DataFrame]:
    st.subheader("👨‍💼 销售员运营毛利增长排名")
    sales_curr = pd.DataFrame(metrics_current["sales_data"]).T
    sales_last = pd.DataFrame(metrics_last["sales_data"]).T
    common_sales = list(set(sales_curr.index) & set(sales_last.index))
    if not common_sales:
        st.info("⚠️ 两个月无共同销售员，无法生成排名")
        return None
    rank_df = pd.DataFrame(index=common_sales)
    rank_df.index.name = "销售员"
    rank_df["运营毛利_上月(元)"] = sales_last["运营毛利"].round(CONFIG["DECIMAL_PLACES"])
    rank_df["运营毛利_本月(元)"] = sales_curr["运营毛利"].round(CONFIG["DECIMAL_PLACES"])
    rank_df["运营毛利_差异(元)"] = (rank_df["运营毛利_本月(元)"] - rank_df["运营毛利_上月(元)"]).round(CONFIG["DECIMAL_PLACES"])
    rank_df["运营毛利_环比(%)"] = calculate_margin_ratio(rank_df["运营毛利_差异(元)"], rank_df["运营毛利_上月(元)"])
    # 新增毛利率、客单价、均单利润排名字段
    rank_df["订单毛利率_上月(%)"] = sales_last["订单毛利率(%)"].round(CONFIG["DECIMAL_PLACES"])
    rank_df["订单毛利率_本月(%)"] = sales_curr["订单毛利率(%)"].round(CONFIG["DECIMAL_PLACES"])
    rank_df["运营毛利率_上月(%)"] = sales_last["运营毛利率(%)"].round(CONFIG["DECIMAL_PLACES"])
    rank_df["运营毛利率_本月(%)"] = sales_curr["运营毛利率(%)"].round(CONFIG["DECIMAL_PLACES"])
    
    if metrics_current["has_sales_quantity"] and metrics_last["has_sales_quantity"]:
        rank_df["商品成本占比_上月(%)"] = sales_last["商品成本占比(%)"].round(CONFIG["DECIMAL_PLACES"])
        rank_df["商品成本占比_本月(%)"] = sales_curr["商品成本占比(%)"].round(CONFIG["DECIMAL_PLACES"])
        rank_df["人工成本占比_上月(%)"] = sales_last["人工成本占比(%)"].round(CONFIG["DECIMAL_PLACES"])
        rank_df["人工成本占比_本月(%)"] = sales_curr["人工成本占比(%)"].round(CONFIG["DECIMAL_PLACES"])
        rank_df["销售数量_上月"] = sales_last["销售数量"].fillna(0).astype(int)
        rank_df["销售数量_本月"] = sales_curr["销售数量"].fillna(0).astype(int)
        rank_df["销量差异"] = rank_df["销售数量_本月"] - rank_df["销售数量_上月"]
        # 客单价、均单利润
        rank_df["客单价_上月(元/单)"] = sales_last["客单价(元/单)"].round(CONFIG["DECIMAL_PLACES"])
        rank_df["客单价_本月(元/单)"] = sales_curr["客单价(元/单)"].round(CONFIG["DECIMAL_PLACES"])
        rank_df["均单利润_上月(元/单)"] = sales_last["均单利润(元/单)"].round(CONFIG["DECIMAL_PLACES"])
        rank_df["均单利润_本月(元/单)"] = sales_curr["均单利润(元/单)"].round(CONFIG["DECIMAL_PLACES"])
    rank_df = rank_df.sort_values("运营毛利_差异(元)", ascending=False)
    rank_df.insert(0, "排名", range(1, len(rank_df)+1))
    
    # 应用阈值标红样式
    styled_rank_df = highlight_threshold_values(rank_df, st.session_state.alert_config)
    st.dataframe(
        styled_rank_df,
        use_container_width=True
    )
    if len(rank_df) > 0:
        top_sales = rank_df.index[0]
        top_growth = rank_df["运营毛利_差异(元)"].iloc[0]
        st.success(f"🥇 运营毛利增长TOP1销售员：{top_sales}（增长 {format_currency(top_growth)}）")
    return rank_df

def render_sales_double_month_analysis(metrics_current: Dict, metrics_last: Dict) -> pd.DataFrame:
    st.subheader("👨‍💼 运营人员对应店铺双月差异分析")
    sales_current = pd.DataFrame(metrics_current["sales_data"]).T
    sales_last = pd.DataFrame(metrics_last["sales_data"]).T
    common_sales = list(set(sales_current.index) & set(sales_last.index))
    if not common_sales:
        return pd.DataFrame()
    return pd.DataFrame()

# ===================== 主入口 =====================
def main():
    # 侧边栏
    with st.sidebar:
        st.image("https://img.icons8.com/fluency/96/000000/bar-chart.png", width=100)
        st.title("📌 功能菜单")
        
        # 先进行密码验证
        if not check_password():
            st.stop()  # 密码错误或未登录时停止执行后续代码
        
        # 密码验证通过后显示菜单（已删除规则页面）
        menu_option = st.radio(
            "选择页面",
            [
                "📊 单月数据分析",
                "📈 双月对比分析",
                "⚙️ 警戒值设置",
                "📥 下载数据模板"
            ]
        )

    if menu_option == "📥 下载数据模板":
        st.subheader("📥 数据模板下载")
        st.download_button(
            label="📥 下载Excel数据模板",
            data=generate_upload_template(),
            file_name="Temu店铺数据上传模板.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.info("💡 按模板填写后上传即可自动分析")
        return

    # 警戒值设置页面
    if menu_option == "⚙️ 警戒值设置":
        render_alert_config_panel()
        return

    # 数据上传区域
    with st.sidebar:
        st.divider()
        uploaded_file1 = st.file_uploader(
            "上传本月数据", type=CONFIG["SUPPORTED_FILE_TYPES"], key="current"
        )
        uploaded_file2 = st.file_uploader(
            "上传上月数据", type=CONFIG["SUPPORTED_FILE_TYPES"], key="last"
        )

    df_current = read_data(uploaded_file1)
    df_last = read_data(uploaded_file2)

    if menu_option == "📊 单月数据分析":
        st.title("📊 单月店铺数据分析")
        if df_current is None:
            st.warning("请先上传本月数据！")
            return
        metrics_current = calculate_metrics(df_current, "本月")
        render_monthly_analysis(metrics_current, df_current)

    elif menu_option == "📈 双月对比分析":
        st.title("📈 双月店铺数据对比")
        if df_current is None or df_last is None:
            st.warning("请同时上传本月与上月数据！")
            return
        metrics_current = calculate_metrics(df_current, "本月")
        metrics_last = calculate_metrics(df_last, "上月")
        render_double_month_analysis(metrics_current, metrics_last)
        render_shop_margin_ranking(metrics_current, metrics_last)
        render_sales_margin_ranking(metrics_current, metrics_last)

if __name__ == "__main__":
    main()
