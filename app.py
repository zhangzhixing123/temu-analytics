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

# ===================== 全局配置 =====================
st.set_page_config(page_title="Temu店铺数据分析工具", layout="wide")
st.title("📊 Temu 店铺数据分析工具（双月对比+销量分析版）")

# --- 初始化session状态 ---
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False

if 'uploaded_file1' not in st.session_state:
    st.session_state.uploaded_file1 = None
if 'uploaded_file2' not in st.session_state:
    st.session_state.uploaded_file2 = None
if 'df_current' not in st.session_state:
    st.session_state.df_current = None
if 'df_last' not in st.session_state:
    st.session_state.df_last = None
if 'metrics_current' not in st.session_state:
    st.session_state.metrics_current = None
if 'metrics_last' not in st.session_state:
    st.session_state.metrics_last = None

CONFIG = {
    "SUPPORTED_FILE_TYPES": ["xlsx", "csv"],
    "ENCODING": "utf-8",
    "DECIMAL_PLACES": 2,
    "FIXED_PASSWORD": "123456"
}

# ===================== 配置文件持久化 =====================
CONFIG_FILE = "alert_config.json"

def load_alert_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {
                "ORDER_MARGIN_RATE_THRESHOLD": 20.0,
                "OPERATE_MARGIN_RATE_THRESHOLD": 15.0,
                "SALES_QUANTITY_THRESHOLD": 3000,
                "UNIT_PRICE_THRESHOLD": 10.0,
                "UNIT_PROFIT_THRESHOLD": 2.0
            }
    else:
        return {
            "ORDER_MARGIN_RATE_THRESHOLD": 20.0,
            "OPERATE_MARGIN_RATE_THRESHOLD": 15.0,
            "SALES_QUANTITY_THRESHOLD": 3000,
            "UNIT_PRICE_THRESHOLD": 10.0,
            "UNIT_PROFIT_THRESHOLD": 2.0
        }

def save_alert_config(config):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)

if 'alert_config' not in st.session_state:
    st.session_state.alert_config = load_alert_config()

# ===================== 密码验证 =====================
def check_password():
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔐 登录")
    
    if st.session_state.authenticated:
        st.sidebar.success("✅ 已登录")
        if st.sidebar.button("退出登录"):
            st.session_state.authenticated = False
            st.rerun()
        return True
    
    pwd = st.sidebar.text_input("请输入访问密码", type="password", key="login_pwd")
    
    if st.sidebar.button("登录"):
        if pwd == CONFIG["FIXED_PASSWORD"]:
            st.session_state.authenticated = True
            st.success("✅ 密码正确，欢迎使用！")
            st.rerun()
        else:
            st.sidebar.error("❌ 密码错误！")
    
    return False

# ===================== 警戒值设置 =====================
def render_alert_config_panel():
    st.subheader("⚙️ 自定义警戒值设置")
    st.info("修改后点击【保存设置】生效")
    
    col1, col2 = st.columns(2)
    with col1:
        order_margin = st.number_input("订单毛利率警戒值(%)", min_value=0.0, max_value=100.0, step=0.1,
            value=st.session_state.alert_config["ORDER_MARGIN_RATE_THRESHOLD"], key="order_margin_threshold")
        operate_margin = st.number_input("运营毛利率警戒值(%)", min_value=0.0, max_value=100.0, step=0.1,
            value=st.session_state.alert_config["OPERATE_MARGIN_RATE_THRESHOLD"], key="operate_margin_threshold")
        unit_profit = st.number_input("均单利润警戒值(元/单)", min_value=0.0, step=0.1,
            value=st.session_state.alert_config["UNIT_PROFIT_THRESHOLD"], key="unit_profit_threshold")
    with col2:
        sales_quantity = st.number_input("销售数量警戒值(单)", min_value=0, step=1,
            value=st.session_state.alert_config["SALES_QUANTITY_THRESHOLD"], key="sales_quantity_threshold")
        unit_price = st.number_input("客单价警戒值(元/单)", min_value=0.0, step=0.1,
            value=st.session_state.alert_config["UNIT_PRICE_THRESHOLD"], key="unit_price_threshold")
    
    if st.button("💾 保存警戒值设置", key="save_alert_config"):
        new_config = {
            "ORDER_MARGIN_RATE_THRESHOLD": order_margin,
            "OPERATE_MARGIN_RATE_THRESHOLD": operate_margin,
            "SALES_QUANTITY_THRESHOLD": sales_quantity,
            "UNIT_PRICE_THRESHOLD": unit_price,
            "UNIT_PROFIT_THRESHOLD": unit_profit
        }
        st.session_state.alert_config.update(new_config)
        save_alert_config(new_config)
        st.success("✅ 保存成功！")
        st.rerun()
    
    st.markdown("### 📌 当前生效的警戒值")
    alert_df = pd.DataFrame({
        "指标名称": ["订单毛利率", "运营毛利率", "销售数量", "客单价", "均单利润"],
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
        st.warning(f"⚠️ 读取失败：{str(e)}")
        return None

def calculate_margin_ratio(numerator, denominator):
    if isinstance(numerator, (int, float)) and isinstance(denominator, (int, float)):
        if denominator <= 0:
            return 0.0
        return round((numerator / denominator * 100), CONFIG["DECIMAL_PLACES"])
    elif isinstance(numerator, pd.Series) and isinstance(denominator, pd.Series):
        return np.where(denominator <= 0, 0.0,
                        round((numerator / denominator * 100), CONFIG["DECIMAL_PLACES"]))
    return 0.0

def highlight_below_threshold(val, threshold):
    try:
        if pd.isna(val) or val == 0:
            return ''
        val_num = float(val) if isinstance(val, (int, float)) else 0
        if val_num < threshold:
            return 'background-color: #ffcccc; color: red; font-weight: bold'
    except:
        pass
    return ''

def highlight_threshold_values(df, alert_config):
    styled_df = df.style
    if '客单价(元/单)' in df.columns:
        styled_df = styled_df.apply(lambda x: [highlight_below_threshold(v, alert_config["UNIT_PRICE_THRESHOLD"]) for v in x], subset=['客单价(元/单)'])
    if '均单利润(元/单)' in df.columns:
        styled_df = styled_df.apply(lambda x: [highlight_below_threshold(v, alert_config["UNIT_PROFIT_THRESHOLD"]) for v in x], subset=['均单利润(元/单)'])
    if '销售数量' in df.columns:
        styled_df = styled_df.apply(lambda x: [highlight_below_threshold(v, alert_config["SALES_QUANTITY_THRESHOLD"]) for v in x], subset=['销售数量'])
    if '订单毛利率(%)' in df.columns:
        styled_df = styled_df.apply(lambda x: [highlight_below_threshold(v, alert_config["ORDER_MARGIN_RATE_THRESHOLD"]) for v in x], subset=['订单毛利率(%)'])
    if '运营毛利率(%)' in df.columns:
        styled_df = styled_df.apply(lambda x: [highlight_below_threshold(v, alert_config["OPERATE_MARGIN_RATE_THRESHOLD"]) for v in x], subset=['运营毛利率(%)'])
    return styled_df

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
    template_cols = ["OA店铺名称", "销售员", "交易收入", "商品成本", "耗材成本", "人工成本",
                     "头程运费", "退回运费", "消费者售后预留金额", "消费者售后释放金额",
                     "店铺总计提金额", "罚款金额", "运营毛利", "销售数量"]
    sample_data = [["示例店铺A", "销售员甲", 10000, 6000, 300, 500, 800, 100, 200, 50, 0, 0, 2000, 200],
                   ["示例店铺B", "销售员乙", 15000, 9000, 450, 750, 1200, 150, 300, 75, 0, 100, 3000, 300]]
    template_df = pd.DataFrame(sample_data, columns=template_cols)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        template_df.to_excel(writer, sheet_name="数据模板", index=False)
        readme_df = pd.DataFrame({"字段名": template_cols, "是否必填": ["是"]*13 + ["可选"],
            "说明": ["店铺名称", "销售员", "交易收入", "商品成本", "耗材成本", "人工成本",
                    "头程运费", "退回运费", "售后预留", "售后释放", "平台提成", "罚款", "运营毛利", "销售数量"]})
        readme_df.to_excel(writer, sheet_name="字段说明", index=False)
    output.seek(0)
    return output.getvalue()

# ===================== 核心计算函数 =====================
def preprocess_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """预处理：确保数值类型正确"""
    df = df.copy()
    numeric_cols = ['交易收入', '商品成本', '运营毛利', '头程运费', '耗材成本', 
                    '人工成本', '退回运费', '罚款金额', '店铺总计提金额',
                    '消费者售后预留金额', '消费者售后释放金额']
    if '销售数量' in df.columns:
        numeric_cols.append('销售数量')
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    return df

def calculate_order_margin(df: pd.DataFrame) -> pd.DataFrame:
    """计算订单毛利"""
    components = []
    if '交易收入' in df.columns:
        components.append(df['交易收入'])
    if '头程运费' in df.columns:
        components.append(-df['头程运费'])
    if '耗材成本' in df.columns:
        components.append(-df['耗材成本'])
    if '人工成本' in df.columns:
        components.append(-df['人工成本'])
    if '商品成本' in df.columns:
        components.append(-df['商品成本'])
    if '消费者售后预留金额' in df.columns:
        components.append(-df['消费者售后预留金额'])
    if '消费者售后释放金额' in df.columns:
        components.append(df['消费者售后释放金额'])
    if components:
        df['订单毛利'] = pd.concat(components, axis=1).sum(axis=1)
    else:
        df['订单毛利'] = 0
    return df

def calculate_sales_per_unit(df: pd.DataFrame) -> pd.DataFrame:
    """计算单均指标"""
    df = df.copy()
    if '销售数量' not in df.columns:
        return df
    
    original_sales = df['销售数量'].copy()
    df['销售数量'] = df['销售数量'].replace(0, np.nan)
    sales_qty = df['销售数量']
    valid_mask = sales_qty.notna()
    
    df['单均订单毛利(元/单)'] = 0.0
    df['单均商品成本(元/单)'] = 0.0
    df['客单价(元/单)'] = 0.0
    df['均单利润(元/单)'] = 0.0
    
    if valid_mask.any():
        if '订单毛利' in df.columns:
            df.loc[valid_mask, '单均订单毛利(元/单)'] = round(df.loc[valid_mask, '订单毛利'] / df.loc[valid_mask, '销售数量'], CONFIG["DECIMAL_PLACES"])
        if '商品成本' in df.columns:
            df.loc[valid_mask, '单均商品成本(元/单)'] = round(df.loc[valid_mask, '商品成本'] / df.loc[valid_mask, '销售数量'], CONFIG["DECIMAL_PLACES"])
        if '交易收入' in df.columns:
            df.loc[valid_mask, '客单价(元/单)'] = round(df.loc[valid_mask, '交易收入'] / df.loc[valid_mask, '销售数量'], CONFIG["DECIMAL_PLACES"])
        if '运营毛利' in df.columns:
            profit_valid = valid_mask & (df['运营毛利'].notna())
            df.loc[profit_valid, '均单利润(元/单)'] = round(df.loc[profit_valid, '运营毛利'] / df.loc[profit_valid, '销售数量'], CONFIG["DECIMAL_PLACES"])
    
    df['销售数量'] = original_sales.fillna(0).astype(int)
    
    if '订单毛利' in df.columns and '交易收入' in df.columns:
        df['订单毛利率(%)'] = calculate_margin_ratio(df['订单毛利'], df['交易收入'])
    if '运营毛利' in df.columns and '交易收入' in df.columns:
        df['运营毛利率(%)'] = calculate_margin_ratio(df['运营毛利'], df['交易收入'])
    return df

def calculate_metrics(df: pd.DataFrame, period_name: str) -> Optional[Dict]:
    if df is None or df.empty:
        return None
    
    # 保存原始数据副本
    original_df = df.copy()
    df = preprocess_dataframe(df)
    df = calculate_order_margin(df)
    df = calculate_sales_per_unit(df)
    
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
        "sales_data": {},
        "original_df": original_df
    }
    
    # 汇总
    metrics["交易收入"] = round(df.get('交易收入', 0).sum(), 2)
    metrics["罚款金额"] = round(df.get('罚款金额', 0).sum(), 2)
    metrics["运营毛利"] = round(df.get('运营毛利', 0).sum(), 2)
    metrics["商品成本"] = round(df.get('商品成本', 0).sum(), 2)
    metrics["耗材成本"] = round(df.get('耗材成本', 0).sum(), 2)
    metrics["人工成本"] = round(df.get('人工成本', 0).sum(), 2)
    metrics["头程运费"] = round(df.get('头程运费', 0).sum(), 2)
    metrics["退回运费"] = round(df.get('退回运费', 0).sum(), 2)
    metrics["店铺总计提金额"] = round(df.get('店铺总计提金额', 0).sum(), 2)
    metrics["消费者售后预留金额"] = round(df.get('消费者售后预留金额', 0).sum(), 2)
    metrics["消费者售后释放金额"] = round(df.get('消费者售后释放金额', 0).sum(), 2)
    metrics["售后净额"] = round(metrics["消费者售后预留金额"] - metrics["消费者售后释放金额"], 2)
    metrics["订单毛利"] = round(df['订单毛利'].sum(), 2)
    
    # 占比
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
    
    # 销量
    if metrics["has_sales_quantity"]:
        try:
            total_qty = int(df['销售数量'].fillna(0).sum())
            metrics["销售数量总计"] = total_qty
            if total_qty > 0:
                metrics["单均订单毛利(元/单)"] = round(metrics["订单毛利"] / total_qty, 2)
                metrics["单均商品成本(元/单)"] = round(metrics["商品成本"] / total_qty, 2)
                metrics["客单价(元/单)"] = round(metrics["交易收入"] / total_qty, 2)
                metrics["均单利润(元/单)"] = round(metrics["运营毛利"] / total_qty, 2)
        except:
            pass
    
    # 店铺分析
    if 'OA店铺名称' in df.columns:
        shop_agg = df.groupby('OA店铺名称').agg({
            '交易收入': 'sum', '订单毛利': 'sum', '运营毛利': 'sum',
            '商品成本': 'sum', '耗材成本': 'sum', '人工成本': 'sum',
            '头程运费': 'sum', '退回运费': 'sum', '罚款金额': 'sum',
            '店铺总计提金额': 'sum', '消费者售后预留金额': 'sum', '消费者售后释放金额': 'sum'
        })
        if metrics["has_sales_quantity"]:
            shop_agg['销售数量'] = df.groupby('OA店铺名称')['销售数量'].sum()
            shop_agg['单均订单毛利(元/单)'] = round(shop_agg['订单毛利'] / shop_agg['销售数量'], 2)
            shop_agg['单均商品成本(元/单)'] = round(shop_agg['商品成本'] / shop_agg['销售数量'], 2)
            shop_agg['客单价(元/单)'] = round(shop_agg['交易收入'] / shop_agg['销售数量'], 2)
            shop_agg['均单利润(元/单)'] = round(shop_agg['运营毛利'] / shop_agg['销售数量'], 2)
        
        shop_agg['店铺数量'] = df.groupby('OA店铺名称').size()
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
    
    # 销售员分析
    if '销售员' in df.columns and metrics["has_sales_quantity"]:
        sales_agg = df.groupby('销售员').agg({
            '交易收入': 'sum', '订单毛利': 'sum', '运营毛利': 'sum',
            '商品成本': 'sum', '耗材成本': 'sum', '人工成本': 'sum',
            '头程运费': 'sum', '退回运费': 'sum', '罚款金额': 'sum',
            '店铺总计提金额': 'sum', '消费者售后预留金额': 'sum', '消费者售后释放金额': 'sum',
            '销售数量': 'sum'
        })
        sales_agg['店铺数量'] = df.groupby('销售员').size()
        sales_agg['订单毛利率(%)'] = calculate_margin_ratio(sales_agg['订单毛利'], sales_agg['交易收入'])
        sales_agg['运营毛利率(%)'] = calculate_margin_ratio(sales_agg['运营毛利'], sales_agg['交易收入'])
        sales_agg['客单价(元/单)'] = round(sales_agg['交易收入'] / sales_agg['销售数量'], 2)
        sales_agg['均单利润(元/单)'] = round(sales_agg['运营毛利'] / sales_agg['销售数量'], 2)
        sales_agg['单均订单毛利(元/单)'] = round(sales_agg['订单毛利'] / sales_agg['销售数量'], 2)
        sales_agg['单均商品成本(元/单)'] = round(sales_agg['商品成本'] / sales_agg['销售数量'], 2)
        
        metrics["sales_data"] = sales_agg.to_dict('index')
    
    return metrics

# ===================== 可视化函数 =====================
def plot_margin_chart(df, y_col, title, threshold_key, selected_items=None):
    threshold = st.session_state.alert_config[threshold_key]
    if selected_items:
        df = df.loc[selected_items]
    fig = px.bar(df, x=df.index, y=y_col, title=title)
    fig.update_traces(text=df[y_col].apply(lambda x: f"{x:.2f}%"), textposition='outside',
                      marker_color=['#d62728' if x < threshold else '#2ca02c' for x in df[y_col]])
    fig.add_hline(y=threshold, line_dash="dash", line_color="orange", annotation_text=f"达标线({threshold}%)")
    return fig

def plot_cost_ratio_chart(df, title, selected_items=None):
    if selected_items:
        df = df.loc[selected_items]
    cost_cols = [c for c in ['商品成本占比(%)', '耗材成本占比(%)', '人工成本占比(%)',
                             '头程运费占比(%)', '退回运费占比(%)', '罚款金额占比(%)',
                             '店铺总计提占比(%)', '售后净额占比(%)'] if c in df.columns]
    if not cost_cols:
        fig = go.Figure()
        fig.add_annotation(text="无成本占比数据", x=0.5, y=0.5, showarrow=False)
        return fig
    fig = go.Figure()
    colors = ['#FF6B6B','#4ECDC4','#45B7D1','#96CEB4','#FFEAA7','#DDA0DD','#98D8C8','#F7DC6F']
    for i, c in enumerate(cost_cols):
        fig.add_trace(go.Bar(x=df.index, y=df[c], name=c.replace('占比(%)',''), marker_color=colors[i%len(colors)],
                            text=df[c].apply(lambda x: f"{x:.2f}%"), textposition='inside'))
    fig.update_layout(title=title, barmode='stack', yaxis_title='占比(%)', height=600)
    return fig

def plot_sales_quantity_chart(df, title, selected_items=None):
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

def plot_unit_metrics_chart(df, title, selected_items=None):
    if selected_items:
        df = df.loc[selected_items]
    up_thresh = st.session_state.alert_config["UNIT_PRICE_THRESHOLD"]
    uprof_thresh = st.session_state.alert_config["UNIT_PROFIT_THRESHOLD"]
    
    if '客单价(元/单)' not in df.columns or '均单利润(元/单)' not in df.columns:
        fig = go.Figure()
        fig.add_annotation(text="无客单价或均单利润数据", x=0.5, y=0.5, showarrow=False)
        return fig
    
    fig = go.Figure()
    fig.add_trace(go.Bar(x=df.index, y=df['客单价(元/单)'], name=f'客单价 (警戒值:{up_thresh}元)',
                         marker_color=['#d62728' if x < up_thresh else '#2ca02c' for x in df['客单价(元/单)']],
                         text=df['客单价(元/单)'].apply(lambda x: f"¥{x:.2f}"), textposition='outside'))
    fig.add_trace(go.Bar(x=df.index, y=df['均单利润(元/单)'], name=f'均单利润 (警戒值:{uprof_thresh}元)',
                         marker_color=['#FF8C00' if x < uprof_thresh else '#006400' for x in df['均单利润(元/单)']],
                         text=df['均单利润(元/单)'].apply(lambda x: f"¥{x:.2f}"), textposition='outside'))
    fig.update_layout(title=title, barmode='group', height=600, yaxis_title='金额(元/单)')
    fig.add_hline(y=up_thresh, line_dash="dash", line_color="red", annotation_text=f"客单价警戒值({up_thresh}元)")
    fig.add_hline(y=uprof_thresh, line_dash="dash", line_color="orange", annotation_text=f"均单利润警戒值({uprof_thresh}元)")
    return fig

def plot_sales_unit_metrics_chart(df, title, selected_items=None):
    if selected_items:
        df = df.loc[selected_items]
    
    fig = go.Figure()
    
    if '单均订单毛利(元/单)' in df.columns:
        fig.add_trace(go.Bar(x=df.index, y=df['单均订单毛利(元/单)'], name='单均订单毛利', marker_color='#FFD700',
                             text=df['单均订单毛利(元/单)'].apply(lambda x: f"¥{x:.2f}"), textposition='outside'))
    
    if '单均商品成本(元/单)' in df.columns:
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

# ===================== 导出功能 =====================
def export_analysis_to_excel(metrics, df, period_name):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        core_metrics = {
            "指标": ["店铺数量", "交易收入", "订单毛利", "订单毛利率(%)", "运营毛利", "运营毛利率(%)",
                    "商品成本", "商品成本占比(%)", "耗材成本", "耗材成本占比(%)", "人工成本", "人工成本占比(%)",
                    "头程运费", "头程运费占比(%)", "退回运费", "退回运费占比(%)", "罚款金额", "罚款金额占比(%)",
                    "售后净额", "售后净额占比(%)"],
            "数值": [metrics["店铺数量"], metrics["交易收入"], metrics["订单毛利"], metrics["订单毛利率"],
                    metrics["运营毛利"], metrics["运营毛利率"], metrics["商品成本"], metrics["商品成本占比"],
                    metrics["耗材成本"], metrics["耗材成本占比"], metrics["人工成本"], metrics["人工成本占比"],
                    metrics["头程运费"], metrics["头程运费占比"], metrics["退回运费"], metrics["退回运费占比"],
                    metrics["罚款金额"], metrics["罚款金额占比"], metrics["售后净额"], metrics["售后净额占比"]]
        }
        if metrics["has_sales_quantity"]:
            core_metrics["指标"].extend(["销售数量总计", "客单价(元/单)", "均单利润(元/单)", "单均订单毛利(元/单)", "单均商品成本(元/单)"])
            core_metrics["数值"].extend([metrics["销售数量总计"], metrics["客单价(元/单)"], metrics["均单利润(元/单)"],
                                         metrics["单均订单毛利(元/单)"], metrics["单均商品成本(元/单)"]])
        pd.DataFrame(core_metrics).to_excel(writer, sheet_name=f"{period_name}核心指标", index=False)
        df.to_excel(writer, sheet_name=f"{period_name}原始数据", index=False)
        if metrics.get("店铺数据"):
            pd.DataFrame(metrics["店铺数据"]).T.to_excel(writer, sheet_name=f"{period_name}店铺数据")
        if metrics.get("sales_data"):
            pd.DataFrame(metrics["sales_data"]).T.to_excel(writer, sheet_name=f"{period_name}销售员数据")
    output.seek(0)
    return output.getvalue()

def render_export_button(metrics, df, period_name):
    col1, _, _ = st.columns([1, 1, 4])
    with col1:
        if st.button("📥 导出Excel数据", key=f"export_{period_name}", use_container_width=True):
            with st.spinner("正在生成..."):
                excel_data = export_analysis_to_excel(metrics, df, period_name)
                st.download_button(label="📥 点击下载", data=excel_data,
                    file_name=f"Temu数据分析_{period_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"download_{period_name}")

def render_double_export_button(curr_metrics, last_metrics, curr_df, last_df):
    col1, _, _ = st.columns([1, 1, 4])
    with col1:
        if st.button("📥 导出双月对比数据", key="export_double", use_container_width=True):
            with st.spinner("正在生成..."):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    pd.DataFrame({"指标": ["店铺数量", "交易收入", "订单毛利", "订单毛利率(%)", "运营毛利", "运营毛利率(%)"],
                                  "数值": [curr_metrics["店铺数量"], curr_metrics["交易收入"], curr_metrics["订单毛利"],
                                          curr_metrics["订单毛利率"], curr_metrics["运营毛利"], curr_metrics["运营毛利率"]]}
                                ).to_excel(writer, sheet_name="本月核心指标", index=False)
                    pd.DataFrame({"指标": ["店铺数量", "交易收入", "订单毛利", "订单毛利率(%)", "运营毛利", "运营毛利率(%)"],
                                  "数值": [last_metrics["店铺数量"], last_metrics["交易收入"], last_metrics["订单毛利"],
                                          last_metrics["订单毛利率"], last_metrics["运营毛利"], last_metrics["运营毛利率"]]}
                                ).to_excel(writer, sheet_name="上月核心指标", index=False)
                    compare_data = {"指标": ["交易收入", "订单毛利", "运营毛利", "销售数量", "客单价", "均单利润"],
                                    "本月": [curr_metrics["交易收入"], curr_metrics["订单毛利"], curr_metrics["运营毛利"],
                                            curr_metrics["销售数量总计"] if curr_metrics["has_sales_quantity"] else 0,
                                            curr_metrics["客单价(元/单)"] if curr_metrics["has_sales_quantity"] else 0,
                                            curr_metrics["均单利润(元/单)"] if curr_metrics["has_sales_quantity"] else 0],
                                    "上月": [last_metrics["交易收入"], last_metrics["订单毛利"], last_metrics["运营毛利"],
                                            last_metrics["销售数量总计"] if last_metrics["has_sales_quantity"] else 0,
                                            last_metrics["客单价(元/单)"] if last_metrics["has_sales_quantity"] else 0,
                                            last_metrics["均单利润(元/单)"] if last_metrics["has_sales_quantity"] else 0]}
                    pd.DataFrame(compare_data).to_excel(writer, sheet_name="双月对比", index=False)
                    curr_df.to_excel(writer, sheet_name="本月原始数据", index=False)
                    last_df.to_excel(writer, sheet_name="上月原始数据", index=False)
                output.seek(0)
                st.download_button(label="📥 点击下载", data=output.getvalue(),
                    file_name=f"Temu双月对比_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_double")

# ===================== 页面渲染 =====================
def render_monthly_analysis(metrics, df):
    st.subheader(f"📈 {metrics['周期']} 核心数据")
    render_export_button(metrics, df, metrics['周期'])
    
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
        o_thresh = st.session_state.alert_config["ORDER_MARGIN_RATE_THRESHOLD"]
        op_thresh = st.session_state.alert_config["OPERATE_MARGIN_RATE_THRESHOLD"]
        st.metric("订单毛利率", f"{o_rate:.2f}%", delta=f"＜{o_thresh}% 警示" if o_rate<o_thresh else f"≥{o_thresh}% 合格", delta_color="inverse" if o_rate<o_thresh else "normal")
        st.metric("运营毛利率", f"{op_rate:.2f}%", delta=f"＜{op_thresh}% 警示" if op_rate<op_thresh else f"≥{op_thresh}% 合格", delta_color="inverse" if op_rate<op_thresh else "normal")
        st.metric("退回运费占比", f"{metrics['退回运费占比']:.2f}%")
        st.metric("罚款占比", f"{metrics['罚款金额占比']:.2f}%")
    with c4:
        if metrics["has_sales_quantity"]:
            s_thresh = st.session_state.alert_config["SALES_QUANTITY_THRESHOLD"]
            st.metric("销售数量总计", f"{metrics['销售数量总计']:,} 单", delta=f"＜{s_thresh} 单 偏低" if metrics['销售数量总计']<s_thresh else f"≥{s_thresh} 单 正常", delta_color="inverse" if metrics['销售数量总计']<s_thresh else "normal")
        else:
            st.metric("提示", "无销售数量数据")
    with c5:
        if metrics["has_sales_quantity"]:
            up_thresh = st.session_state.alert_config["UNIT_PRICE_THRESHOLD"]
            uprof_thresh = st.session_state.alert_config["UNIT_PROFIT_THRESHOLD"]
            st.metric("客单价", f"{metrics['客单价(元/单)']:.2f} 元/单", delta=f"＜{up_thresh} 元 偏低" if metrics['客单价(元/单)']<up_thresh else f"≥{up_thresh} 元 正常", delta_color="inverse" if metrics['客单价(元/单)']<up_thresh else "normal")
            st.metric("均单利润", f"{metrics['均单利润(元/单)']:.2f} 元/单", delta=f"＜{uprof_thresh} 元 偏低" if metrics['均单利润(元/单)']<uprof_thresh else f"≥{uprof_thresh} 元 正常", delta_color="inverse" if metrics['均单利润(元/单)']<uprof_thresh else "normal")
        else:
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
        if metrics["has_sales_quantity"] and '销售数量' in shop_df.columns:
            shop_df['销售数量'] = shop_df['销售数量'].fillna(0).astype(int)
        shops = st.multiselect("选择店铺", list(shop_df.index), default=list(shop_df.index)[:5] if len(shop_df) >= 5 else list(shop_df.index))
        if shops:
            st.dataframe(highlight_threshold_values(shop_df.loc[shops], st.session_state.alert_config), use_container_width=True)
            st.plotly_chart(plot_margin_chart(shop_df, '订单毛利率(%)', '店铺订单毛利率', 'ORDER_MARGIN_RATE_THRESHOLD', shops), use_container_width=True)
            st.plotly_chart(plot_cost_ratio_chart(shop_df, '店铺成本占比', shops), use_container_width=True)
            if metrics["has_sales_quantity"] and '销售数量' in shop_df.columns:
                st.plotly_chart(plot_sales_quantity_chart(shop_df, '店铺销售数量排名', shops), use_container_width=True)
                st.plotly_chart(plot_sales_unit_metrics_chart(shop_df, '店铺单均指标', shops), use_container_width=True)
                st.plotly_chart(plot_unit_metrics_chart(shop_df, '店铺客单价&均单利润对比', shops), use_container_width=True)
    
    if metrics["sales_data"]:
        st.subheader("👨‍💼 销售员分析")
        sales_df = pd.DataFrame(metrics["sales_data"]).T
        if '店铺数量' in sales_df.columns:
            sales_df['店铺数量'] = sales_df['店铺数量'].astype(int)
        if metrics["has_sales_quantity"] and '销售数量' in sales_df.columns:
            sales_df['销售数量'] = sales_df['销售数量'].fillna(0).astype(int)
        sales = st.multiselect("选择销售员", list(sales_df.index), default=list(sales_df.index)[:5] if len(sales_df) >= 5 else list(sales_df.index))
        if sales:
            st.dataframe(highlight_threshold_values(sales_df.loc[sales], st.session_state.alert_config), use_container_width=True)
            st.plotly_chart(plot_cost_ratio_chart(sales_df, '销售员成本占比', sales), use_container_width=True)
            if metrics["has_sales_quantity"] and '销售数量' in sales_df.columns:
                st.plotly_chart(plot_sales_quantity_chart(sales_df, '销售员销量排名', sales), use_container_width=True)
                st.plotly_chart(plot_sales_unit_metrics_chart(sales_df, '销售员单均指标', sales), use_container_width=True)
                st.plotly_chart(plot_unit_metrics_chart(sales_df, '销售员客单价&均单利润对比', sales), use_container_width=True)
        
        st.subheader("📌 运营建议")
        for name, row in sales_df.iterrows():
            with st.expander(f"🧑‍💼 {name}", expanded=False):
                o_thresh = st.session_state.alert_config["ORDER_MARGIN_RATE_THRESHOLD"]
                op_thresh = st.session_state.alert_config["OPERATE_MARGIN_RATE_THRESHOLD"]
                s_thresh = st.session_state.alert_config["SALES_QUANTITY_THRESHOLD"]
                up_thresh = st.session_state.alert_config["UNIT_PRICE_THRESHOLD"]
                uprof_thresh = st.session_state.alert_config["UNIT_PROFIT_THRESHOLD"]
                o_rate = row.get('订单毛利率(%)', 0)
                op_rate = row.get('运营毛利率(%)', 0)
                fine = row.get('罚款占收入比(%)', 0)
                
                if o_rate < o_thresh:
                    st.warning(f"订单毛利率{o_rate:.2f}% 不达标（标准{o_thresh}%）")
                else:
                    st.success(f"订单毛利率{o_rate:.2f}% 达标（标准{o_thresh}%）")
                if op_rate < op_thresh:
                    st.warning(f"运营毛利率{op_rate:.2f}% 不达标（标准{op_thresh}%）")
                else:
                    st.success(f"运营毛利率{op_rate:.2f}% 达标（标准{op_thresh}%）")
                if fine > 0:
                    st.info(f"罚款占收入比{fine:.2f}%，需减少违规订单")
                if metrics["has_sales_quantity"]:
                    qty = row.get('销售数量', 0)
                    unit_price = row.get('客单价(元/单)', 0)
                    unit_profit = row.get('均单利润(元/单)', 0)
                    if qty < s_thresh:
                        st.warning(f"销量{qty} 偏低（建议≥{s_thresh}单）")
                    else:
                        st.success(f"销量{qty} 正常（≥{s_thresh}单）")
                    if unit_price < up_thresh:
                        st.warning(f"客单价{unit_price:.2f}元 偏低（建议≥{up_thresh}元）")
                    else:
                        st.success(f"客单价{unit_price:.2f}元 正常（≥{up_thresh}元）")
                    if unit_profit < uprof_thresh:
                        st.warning(f"均单利润{unit_profit:.2f}元 偏低（建议≥{uprof_thresh}元）")
                    else:
                        st.success(f"均单利润{unit_profit:.2f}元 正常（≥{uprof_thresh}元）")

def render_double_month_analysis(curr, last, curr_df, last_df):
    st.subheader("🔍 双月详细对比分析")
    render_double_export_button(curr, last, curr_df, last_df)
    
    compare_data = {
        "指标名称": ["店铺数量", "交易收入(元)", "订单毛利(元)", "订单毛利率(%)", "运营毛利(元)", "运营毛利率(%)",
                    "罚款金额(元)", "罚款占收入比(%)", "商品成本(元)", "商品成本占比(%)", "耗材成本(元)", "耗材成本占比(%)",
                    "人工成本(元)", "人工成本占比(%)", "头程运费(元)", "头程运费占比(%)", "退回运费(元)", "退回运费占比(%)",
                    "店铺总计提(元)", "店铺总计提占比(%)", "售后净额(元)", "售后净额占比(%)"],
        "上月数值": [int(last["店铺数量"]), last["交易收入"], last["订单毛利"], last["订单毛利率"], last["运营毛利"], last["运营毛利率"],
                    last["罚款金额"], calculate_margin_ratio(last["罚款金额"], last["交易收入"]),
                    last["商品成本"], last["商品成本占比"], last["耗材成本"], last["耗材成本占比"],
                    last["人工成本"], last["人工成本占比"], last["头程运费"], last["头程运费占比"],
                    last["退回运费"], last["退回运费占比"], last["店铺总计提金额"], last["店铺总计提占比"],
                    last["售后净额"], last["售后净额占比"]],
        "本月数值": [int(curr["店铺数量"]), curr["交易收入"], curr["订单毛利"], curr["订单毛利率"], curr["运营毛利"], curr["运营毛利率"],
                    curr["罚款金额"], calculate_margin_ratio(curr["罚款金额"], curr["交易收入"]),
                    curr["商品成本"], curr["商品成本占比"], curr["耗材成本"], curr["耗材成本占比"],
                    curr["人工成本"], curr["人工成本占比"], curr["头程运费"], curr["头程运费占比"],
                    curr["退回运费"], curr["退回运费占比"], curr["店铺总计提金额"], curr["店铺总计提占比"],
                    curr["售后净额"], curr["售后净额占比"]]
    }
    
    if curr["has_sales_quantity"] and last["has_sales_quantity"]:
        compare_data["指标名称"].extend(["销售数量(单)", "客单价(元/单)", "均单利润(元/单)", "单均订单毛利(元/单)"])
        compare_data["上月数值"].extend([last["销售数量总计"], last["客单价(元/单)"], last["均单利润(元/单)"], last["单均订单毛利(元/单)"]])
        compare_data["本月数值"].extend([curr["销售数量总计"], curr["客单价(元/单)"], curr["均单利润(元/单)"], curr["单均订单毛利(元/单)"]])
    
    compare_df = pd.DataFrame(compare_data)
    compare_df["绝对差异"] = compare_df["本月数值"] - compare_df["上月数值"]
    compare_df["相对差异(%)"] = round((compare_df["绝对差异"] / compare_df["上月数值"] * 100).replace([np.inf, -np.inf], 0), 2)
    
    styled_compare = compare_df.style.apply(lambda x: [highlight_negative_values(v) for v in x], subset=["绝对差异", "相对差异(%)"])
    st.dataframe(styled_compare, use_container_width=True)
    
    st.markdown("### 关键金额指标双月对比")
    fig_amount = go.Figure()
    for metric in ["交易收入", "订单毛利", "运营毛利", "商品成本", "人工成本", "头程运费"]:
        fig_amount.add_trace(go.Bar(x=['上月', '本月'], y=[last[metric], curr[metric]], name=metric,
                                    text=[format_currency(last[metric]), format_currency(curr[metric])], textposition='outside'))
    fig_amount.update_layout(title='交易收入/毛利/核心成本 双月对比', barmode='group')
    st.plotly_chart(fig_amount, use_container_width=True)
    
    if curr["has_sales_quantity"] and last["has_sales_quantity"]:
        st.markdown("### 销量&客单价/均单利润指标双月对比")
        fig_sales = go.Figure()
        fig_sales.add_trace(go.Bar(x=['上月','本月'], y=[last["销售数量总计"], curr["销售数量总计"]], name='销售数量',
                                   text=[f"{last['销售数量总计']:,}", f"{curr['销售数量总计']:,}"], textposition='outside'))
        fig_sales.add_trace(go.Bar(x=['上月','本月'], y=[last["客单价(元/单)"], curr["客单价(元/单)"]], name='客单价',
                                   text=[f"¥{last['客单价(元/单)']:.2f}", f"¥{curr['客单价(元/单)']:.2f}"], textposition='outside'))
        fig_sales.add_trace(go.Bar(x=['上月','本月'], y=[last["均单利润(元/单)"], curr["均单利润(元/单)"]], name='均单利润',
                                   text=[f"¥{last['均单利润(元/单)']:.2f}", f"¥{curr['均单利润(元/单)']:.2f}"], textposition='outside'))
        fig_sales.update_layout(title='销量 & 客单价 & 均单利润对比', barmode='group')
        st.plotly_chart(fig_sales, use_container_width=True)
    
    st.markdown("### 双月差异总结")
    income_diff = curr["交易收入"] - last["交易收入"]
    if income_diff > 0:
        st.success(f"✅ 交易收入环比增长 {format_currency(income_diff)}")
    elif income_diff < 0:
        st.error(f"❌ 交易收入环比下降 {format_currency(abs(income_diff))}")
    order_diff = curr["订单毛利"] - last["订单毛利"]
    if order_diff > 0:
        st.success(f"✅ 订单毛利环比增长 {format_currency(order_diff)}")
    elif order_diff < 0:
        st.error(f"❌ 订单毛利环比下降 {format_currency(abs(order_diff))}")
    if curr["has_sales_quantity"] and last["has_sales_quantity"]:
        qty_diff = curr["销售数量总计"] - last["销售数量总计"]
        if qty_diff > 0:
            st.success(f"✅ 销售数量环比增长 {qty_diff:,} 单")
        elif qty_diff < 0:
            st.error(f"❌ 销售数量环比下降 {abs(qty_diff):,} 单")
        up_diff = curr["客单价(元/单)"] - last["客单价(元/单)"]
        if up_diff > 0:
            st.success(f"✅ 客单价环比增长 {up_diff:.2f} 元")
        elif up_diff < 0:
            st.error(f"❌ 客单价环比下降 {abs(up_diff):.2f} 元")
        uprof_diff = curr["均单利润(元/单)"] - last["均单利润(元/单)"]
        if uprof_diff > 0:
            st.success(f"✅ 均单利润环比增长 {uprof_diff:.2f} 元")
        elif uprof_diff < 0:
            st.error(f"❌ 均单利润环比下降 {abs(uprof_diff):.2f} 元")

def render_shop_margin_ranking(curr, last):
    st.subheader("🏪 店铺运营毛利增长排名")
    if not curr.get("店铺数据") or not last.get("店铺数据"):
        st.info("⚠️ 无店铺数据")
        return
    shop_curr = pd.DataFrame(curr["店铺数据"]).T
    shop_last = pd.DataFrame(last["店铺数据"]).T
    common = list(set(shop_curr.index) & set(shop_last.index))
    if not common:
        st.info("⚠️ 两个月无共同店铺")
        return
    rank_df = pd.DataFrame(index=common)
    rank_df["运营毛利_上月"] = shop_last["运营毛利"].round(2)
    rank_df["运营毛利_本月"] = shop_curr["运营毛利"].round(2)
    rank_df["运营毛利_差异"] = rank_df["运营毛利_本月"] - rank_df["运营毛利_上月"]
    rank_df = rank_df.sort_values("运营毛利_差异", ascending=False)
    rank_df.insert(0, "排名", range(1, len(rank_df)+1))
    st.dataframe(highlight_threshold_values(rank_df, st.session_state.alert_config), use_container_width=True)
    if len(rank_df) > 0:
        st.success(f"🥇 增长TOP1：{rank_df.index[0]}（增长 {format_currency(rank_df['运营毛利_差异'].iloc[0])}）")

def render_sales_margin_ranking(curr, last):
    st.subheader("👨‍💼 销售员运营毛利增长排名")
    if not curr.get("sales_data") or not last.get("sales_data"):
        st.info("⚠️ 无销售员数据")
        return
    sales_curr = pd.DataFrame(curr["sales_data"]).T
    sales_last = pd.DataFrame(last["sales_data"]).T
    common = list(set(sales_curr.index) & set(sales_last.index))
    if not common:
        st.info("⚠️ 两个月无共同销售员")
        return
    rank_df = pd.DataFrame(index=common)
    rank_df["运营毛利_上月"] = sales_last["运营毛利"].round(2)
    rank_df["运营毛利_本月"] = sales_curr["运营毛利"].round(2)
    rank_df["运营毛利_差异"] = rank_df["运营毛利_本月"] - rank_df["运营毛利_上月"]
    rank_df = rank_df.sort_values("运营毛利_差异", ascending=False)
    rank_df.insert(0, "排名", range(1, len(rank_df)+1))
    st.dataframe(highlight_threshold_values(rank_df, st.session_state.alert_config), use_container_width=True)
    if len(rank_df) > 0:
        st.success(f"🥇 增长TOP1：{rank_df.index[0]}（增长 {format_currency(rank_df['运营毛利_差异'].iloc[0])}）")

# ===================== 主入口 =====================
def main():
    with st.sidebar:
        st.image("https://img.icons8.com/fluency/96/000000/bar-chart.png", width=100)
        st.title("📌 功能菜单")
        if not check_password():
            st.stop()
        menu_option = st.radio("选择页面", ["📊 单月数据分析", "📈 双月对比分析", "⚙️ 警戒值设置", "📥 下载数据模板"])

    if menu_option == "📥 下载数据模板":
        st.subheader("📥 数据模板下载")
        st.download_button(label="📥 下载Excel数据模板", data=generate_upload_template(),
                          file_name="Temu店铺数据上传模板.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.info("💡 按模板填写后上传即可自动分析")
        return

    if menu_option == "⚙️ 警戒值设置":
        render_alert_config_panel()
        return

    with st.sidebar:
        st.divider()
        uploaded_file1 = st.file_uploader("上传本月数据", type=CONFIG["SUPPORTED_FILE_TYPES"], key="current")
        uploaded_file2 = st.file_uploader("上传上月数据", type=CONFIG["SUPPORTED_FILE_TYPES"], key="last")

    df_current = read_data(uploaded_file1)
    df_last = read_data(uploaded_file2)
    st.session_state.df_current = df_current
    st.session_state.df_last = df_last

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
        render_double_month_analysis(metrics_current, metrics_last, df_current, df_last)
        render_shop_margin_ranking(metrics_current, metrics_last)
        render_sales_margin_ranking(metrics_current, metrics_last)

if __name__ == "__main__":
    main()
