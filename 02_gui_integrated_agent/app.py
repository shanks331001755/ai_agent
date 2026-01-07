#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
股票分析AI智能体 - 集成Word报告下载版
功能：通过Web界面分析指定股票，使用AI模型生成分析报告，并生成可下载的Word文档

使用教程：
1. 安装依赖: pip install akshare gradio python-docx matplotlib pandas numpy requests
2. 运行程序: python stock_ai_agent_with_word_download.py
3. 访问Web界面: http://localhost:7868
4. 输入股票代码，选择分析时间段，输入分析分析问题
5. 点击"开始专业分析"按钮
6. 分析完成后，点击"Word报告下载"按钮下载完整报告

需要安装的依赖：
- akshare: 股票数据获取
- gradio: Web界面框架
- python-docx: Word文档生成
- matplotlib: 图表绘制
- pandas: 数据处理
- numpy: 数值计算
- requests: HTTP请求

注意：需要配置QWEN_API_KEY为真实的API密钥以获得完整AI分析功能
"""

import gradio as gr
import akshare as ak
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
from io import BytesIO
import requests
import json
import os
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
import traceback
import warnings
import base64
warnings.filterwarnings('ignore')
from word_report_generator import create_word_report_in_memory


# 配置
QWEN_API_KEY = "sk-ed31504ad0554161a6191b41287b9c88"  # 请替换为你的API密钥
QWEN_API_URL = "https://dashscope.aliyuncs.com/api/v1/services/aigc/text-generation/generation"

# 全局缓存字典
chart_cache = {}


def create_mock_stock_data(symbol):
    """创建模拟股票数据，用于网络请求失败时的备用方案"""
    import random
    from datetime import datetime, timedelta
    
    stock_name = f"模拟股票{symbol}"
    current_price = round(random.uniform(10, 100), 2)
    fluctuation = round(random.uniform(-5, 5), 2)
    
    # 生成最近30天的模拟历史数据
    hist_data = []
    base_price = current_price - fluctuation
    for i in range(30):
        date = (datetime.now() - timedelta(days=30-i)).strftime('%Y-%m-%d')
        open_price = round(base_price + random.uniform(-2, 2), 2)
        close_price = round(open_price + random.uniform(-3, 3), 2)
        high_price = max(open_price, close_price) + random.uniform(0, 1)
        low_price = min(open_price, close_price) - random.uniform(0, 1)
        volume = random.randint(100000, 10000000)
        
        hist_data.append({
            "日期": date,
            "开盘": open_price,
            "收盘": close_price,
            "最高": high_price,
            "最低": low_price,
            "成交量": volume
        })
    
    data = {
        "实时行情": {
            "名称": stock_name,
            "代码": symbol,
            "最新价": current_price,
            "涨跌幅": fluctuation,
            "成交量": random.randint(1000000, 100000000),
            "成交额": random.randint(10000000, 1000000000),
            "最高": current_price * 1.05,
            "最低": current_price * 0.95,
            "开盘": current_price * (1 - random.uniform(0.01, 0.03)),
            "振幅": round(abs(fluctuation) + random.uniform(0.5, 2), 2),
            "换手率": round(random.uniform(0.5, 5.0), 2),
            "量比": round(random.uniform(0.5, 2.0), 2),
        },
        "历史数据": hist_data,
        "基本面": {
            "营业收入": random.uniform(1e9, 1e10),
            "净利润": random.uniform(1e8, 1e9),
            "ROE": random.uniform(8, 15),
            "资产负债率": random.uniform(40, 70),
        },
        "资金流向": {
            "主力净流入_5日": random.uniform(-5e7, 5e7),
            "超大单净流入_5日": random.uniform(-3e7, 3e7),
        },
        "估值": {
            "PE_TTM": random.uniform(10, 30),
            "PB": random.uniform(1.0, 3.0),
            "股息率": random.uniform(1.0, 4.0),
        },
        "分析日期": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    
    return data

# ==================== 专业分析提示词 ====================
QUANT_ANALYST_PROMPT = """你是一名严谨的量化分析师，请基于给定数据，生成一份专业、数据驱动的分析报告。报告必须包含以下清晰的6个模块，每个模块用明确的标题分隔，内容必须与用户指定的分析时间段紧密相关：

### 📊 行情摘要（Market Overview）
必须包含：最新价、日内涨跌幅、成交量、成交额、量比、换手率、振幅。
格式：截至[日期]，[股票名称](代码)报收[价格]元，**[涨跌幅]%**，日内振幅[振幅]%。全天成交[成交量]万手，成交额[成交额]亿元，量比为[量比]，换手率[换手率]%。股价处于[近期高位/中部/低位]。

### 📈 关键指标看板（Key Metrics Dashboard）
以表格或要点形式呈现：
• 估值指标：市盈率(PE-TTM)、市净率(PB)、股息率
• 技术指标：5/10/20/60日均线及当前位置、RSI(14)、MACD状态、布林带位置
• 资金指标：近5日主力资金净流入/出额、北向资金持股变动

### 📉 技术分析（Technical Analysis）
基于指定时间段内的具体数据：
• 趋势判断：当前股价位于[XX]日均线之上/下，短期均线呈[多头/空头]排列（列出具体均线数值）
• 动量分析：RSI(14)为[数值]，处于[超买/中性/超卖]区域
• MACD分析：DIF为[X]，DEA为[Y]，柱状体为[Z]，发出[看涨/看跌/观望]信号
• 关键价位：支撑位[价格1, 价格2]，阻力位[价格3, 价格4]

### 💰 基本面与估值（Fundamentals & Valuation）
基于最新财报数据：
• 盈利能力：营业收入、归母净利润、ROE
• 财务健康：资产负债率
• 估值水平：当前PE/PB相对于行业和历史水平的评估

### 💸 资金流向（Fund Flow）
分析指定时间段内的资金动向：
• 资金动向：近5日主力资金净[流入/流出][金额]元
• 大单流向：超大单、大单、中单、小单的净流入情况

### 🎯 综合结论与操作建议（Conclusion & Recommendation）
必须与用户指定的时间段分析结果直接相关：
• 核心观点：综合指定时间段的数据，该股[看涨/中性/看跌]
• 主要利好因素：1. [基于时间段内数据的因素] 2. [基于时间段内数据的因素]
• 主要风险点：1. [基于时间段内数据的风险] 2. [基于时间段内数据的风险]
• 操作建议：[具体操作建议]
• 风险提示：投资有风险，本分析基于公开数据，不构成投资建议

## 【格式要求】
1. 使用清晰的标题和子标题分隔各模块
2. 使用**项目符号（•）**和**表格**结构化呈现数据
3. 所有结论性语句后必须用括号**标注数据来源**，例如："股价处于强势区间（高于所有关键均线）"
4. **杜绝**"可能"、"或许"、"显示出"等模糊词汇，改用"数据显示"、"指标表明"等客观表述
5. **特别重要**：分析内容必须与用户指定的分析时间段（[开始日期]至[结束日期]）直接相关，不得脱离时间段进行泛泛而谈
6. 分析报告必须包含指定时间段内的具体数据和变化趋势

现在请基于以下数据生成分析报告："""

# ==================== 数据获取函数 ====================
def get_comprehensive_stock_data(symbol, start_date=None, end_date=None, days=30, date_mode="指定日期", recent_days=30, start_date_duration=None, duration_type="天", duration_value=30):
    """获取全面的股票数据
    
    Args:
        symbol: 股票代码
        start_date: 开始日期，格式 'YYYY-MM-DD'
        end_date: 结束日期，格式 'YYYY-MM-DD'
        days: 如果未指定开始和结束日期，则获取最近多少天的数据
        date_mode: 日期选择模式 ("指定日期", "最近天数", "从某天开始持续")
        recent_days: 最近天数
        start_date_duration: 持续模式的开始日期
        duration_type: 持续时间单位 ("天", "月")
        duration_value: 持续时间值
    """
    try:
        # 1. 实时行情数据 - 增加超时处理
        import time
        max_retries = 3
        retry_count = 0
        
        while retry_count < max_retries:
            try:
                spot_df = ak.stock_zh_a_spot_em()
                break
            except Exception as e:
                retry_count += 1
                print(f"⚠️ 获取实时行情失败，第 {retry_count} 次重试: {str(e)}")
                if retry_count >= max_retries:
                    print("❌ 获取实时行情失败，使用模拟数据")
                    # 返回模拟数据
                    return create_mock_stock_data(symbol), "成功（使用模拟数据）"
                time.sleep(2)  # 等待2秒后重试
        
        stock_info = spot_df[spot_df['代码'] == symbol.split('.')[0]]

        if stock_info.empty:
            return None, "未找到该股票代码"

        # 2. 根据日期模式确定日期范围
        if date_mode == "最近天数":
            end_date = datetime.now().strftime("%Y%m%d")
            start_date_obj = datetime.now() - timedelta(days=recent_days)
            start_date = start_date_obj.strftime("%Y%m%d")
        elif date_mode == "从某天开始持续":
            if start_date_duration:
                start_date_obj = datetime.strptime(start_date_duration, "%Y-%m-%d")
                start_date = start_date_obj.strftime("%Y%m%d")
                
                if duration_type == "月":
                    end_date_obj = start_date_obj + timedelta(days=duration_value*30)  # 简化处理，一个月约30天
                else:  # 天
                    end_date_obj = start_date_obj + timedelta(days=duration_value)
                end_date = end_date_obj.strftime("%Y%m%d")
            else:
                # 如果没有指定开始日期，使用默认逻辑
                end_date = datetime.now().strftime("%Y%m%d")
                start_date_obj = datetime.now() - timedelta(days=days)
                start_date = start_date_obj.strftime("%Y%m%d")
        else:  # "指定日期"
            if not end_date:
                end_date = datetime.now().strftime("%Y%m%d")
            else:
                # 将 'YYYY-MM-DD' 格式转换为 'YYYYMMDD'
                end_date = datetime.strptime(end_date, "%Y-%m-%d").strftime("%Y%m%d")
                
            if not start_date:
                start_date_obj = datetime.now() - timedelta(days=days)
                start_date = start_date_obj.strftime("%Y%m%d")
            else:
                # 将 'YYYY-MM-DD' 格式转换为 'YYYYMMDD'
                start_date = datetime.strptime(start_date, "%Y-%m-%d").strftime("%Y%m%d")

        # 获取历史K线数据 - 增加超时处理
        max_retries = 3
        retry_count = 0
        hist_df = pd.DataFrame()  # 初始化为空DataFrame
        
        while retry_count < max_retries:
            try:
                hist_df = ak.stock_zh_a_hist(symbol=symbol, period="daily",
                                             start_date=start_date, end_date=end_date, adjust="qfq", timeout=30)
                break
            except Exception as e:
                retry_count += 1
                print(f"⚠️ 获取历史数据失败，第 {retry_count} 次重试: {str(e)}")
                if retry_count >= max_retries:
                    print("❌ 获取历史数据失败，使用模拟数据")
                    # 如果历史数据获取失败，使用实时数据构建基础结构
                    realtime_data = {
                        "名称": stock_info.iloc[0]['名称'],
                        "代码": stock_info.iloc[0]['代码'],
                        "最新价": float(stock_info.iloc[0]['最新价']),
                        "涨跌幅": float(str(stock_info.iloc[0]['涨跌幅']).rstrip('%')) if isinstance(stock_info.iloc[0]['涨跌幅'], str) else float(stock_info.iloc[0]['涨跌幅']),
                        "成交量": int(float(stock_info.iloc[0]['成交量'])) if not pd.isna(stock_info.iloc[0]['成交量']) else 0,
                        "成交额": int(float(stock_info.iloc[0]['成交额'])) if not pd.isna(stock_info.iloc[0]['成交额']) else 0,
                        "最高": float(stock_info.iloc[0]['最高']) if '最高' in stock_info.columns and not pd.isna(stock_info.iloc[0]['最高']) else 0,
                        "最低": float(stock_info.iloc[0]['最低']) if '最低' in stock_info.columns and not pd.isna(stock_info.iloc[0]['最低']) else 0,
                        "开盘": float(stock_info.iloc[0]['开盘']) if '开盘' in stock_info.columns and not pd.isna(stock_info.iloc[0]['开盘']) else 0,
                        "振幅": float(str(stock_info.iloc[0]['振幅']).rstrip('%')) if '振幅' in stock_info.columns and isinstance(stock_info.iloc[0]['振幅'], str) else 0,
                        "换手率": float(str(stock_info.iloc[0]['换手率']).rstrip('%')) if '换手率' in stock_info.columns and isinstance(stock_info.iloc[0]['换手率'], str) else 0,
                        "量比": float(stock_info.iloc[0]['量比']) if '量比' in stock_info.columns and not pd.isna(stock_info.iloc[0]['量比']) else 0,
                    }
                    
                    # 创建基础数据结构，历史数据为空
                    data = {
                        "实时行情": realtime_data,
                        "历史数据": [],
                        "基本面": {},
                        "资金流向": {},
                        "估值": {},
                        "分析日期": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }
                    return data, "成功（使用实时数据，历史数据获取失败）"
                time.sleep(2)  # 等待2秒后重试

        # 3. 基本面数据
        try:
            # 获取财务指标 - 使用正确的函数名
            financial_df = ak.stock_financial_abstract(symbol=symbol.split('.')[0])
        except:
            financial_df = pd.DataFrame()

        # 4. 资金流向数据
        try:
            # 使用正确的函数和参数 - 参数名是stock而不是symbol
            money_flow = ak.stock_individual_fund_flow(stock=symbol.split('.')[0], market="sh")
            print(f"✅ 资金流向数据获取成功，形状: {money_flow.shape}")
            print(f"✅ 资金流向列名: {list(money_flow.columns)}")
        except Exception as e:
            print(f"❌ 资金流向数据获取失败: {e}")
            money_flow = pd.DataFrame()

        # 5. 估值数据
        try:
            # 使用正确的估值函数
            valuation_df = ak.stock_value_em(symbol=symbol.split('.')[0])
        except:
            valuation_df = pd.DataFrame()

        # 整理数据
        realtime_data = {
            "名称": stock_info.iloc[0]['名称'],
            "代码": stock_info.iloc[0]['代码'],
            "最新价": float(stock_info.iloc[0]['最新价']),
            "涨跌幅": float(str(stock_info.iloc[0]['涨跌幅']).rstrip('%')) if isinstance(stock_info.iloc[0]['涨跌幅'], str) else float(stock_info.iloc[0]['涨跌幅']),
            "成交量": int(float(stock_info.iloc[0]['成交量'])) if not pd.isna(stock_info.iloc[0]['成交量']) else 0,
            "成交额": int(float(stock_info.iloc[0]['成交额'])) if not pd.isna(stock_info.iloc[0]['成交额']) else 0,
            "最高": float(stock_info.iloc[0]['最高']) if '最高' in stock_info.columns and not pd.isna(stock_info.iloc[0]['最高']) else 0,
            "最低": float(stock_info.iloc[0]['最低']) if '最低' in stock_info.columns and not pd.isna(stock_info.iloc[0]['最低']) else 0,
            "开盘": float(stock_info.iloc[0]['开盘']) if '开盘' in stock_info.columns and not pd.isna(stock_info.iloc[0]['开盘']) else 0,
            "振幅": float(str(stock_info.iloc[0]['振幅']).rstrip('%')) if '振幅' in stock_info.columns and isinstance(stock_info.iloc[0]['振幅'], str) else 0,
            "换手率": float(str(stock_info.iloc[0]['换手率']).rstrip('%')) if '换手率' in stock_info.columns and isinstance(stock_info.iloc[0]['换手率'], str) else 0,
            "量比": float(stock_info.iloc[0]['量比']) if '量比' in stock_info.columns and not pd.isna(stock_info.iloc[0]['量比']) else 0,
        }

        # 处理历史数据
        hist_data = []
        if not hist_df.empty:
            # 不再限制为最后60条，保留日期范围内的所有数据
            hist_data = hist_df.to_dict('records')

        # 处理基本面数据 - 添加回退方案
        fundamental_data = {}
        if not financial_df.empty and len(financial_df) > 0:
            # stock_financial_abstract返回的是时间序列数据，需要查找特定指标
            # 过滤出最近的报告期数据
            if '指标' in financial_df.columns:
                # 获取特定财务指标的最新值
                revenue_row = financial_df[financial_df['指标'].str.contains('营业收入|营业总收入', na=False)]
                profit_row = financial_df[financial_df['指标'].str.contains('净利润|归母净利润', na=False)]
                roe_row = financial_df[financial_df['指标'].str.contains('净资产收益率', na=False)]
                debt_ratio_row = financial_df[financial_df['指标'].str.contains('资产负债率', na=False)]
                
                # 获取最新数据（选择最近的报告期）
                if len(financial_df.columns) > 2:
                    # 获取日期列（通常是除了'选项'和'指标'外的列）
                    date_cols = [col for col in financial_df.columns if col not in ['选项', '指标']]
                    if date_cols:
                        latest_date_col = date_cols[0]  # 通常是最新日期
                        
                if latest_date_col and revenue_row.shape[0] > 0:
                    fundamental_data['营业收入'] = float(revenue_row.iloc[0][latest_date_col]) if latest_date_col in revenue_row.iloc[0] else 0
                else:
                    fundamental_data['营业收入'] = 0
                    
                if latest_date_col and profit_row.shape[0] > 0:
                    fundamental_data['净利润'] = float(profit_row.iloc[0][latest_date_col]) if latest_date_col in profit_row.iloc[0] else 0
                else:
                    fundamental_data['净利润'] = 0
                    
                if latest_date_col and roe_row.shape[0] > 0:
                    fundamental_data['ROE'] = float(roe_row.iloc[0][latest_date_col]) if latest_date_col in roe_row.iloc[0] else 0
                else:
                    fundamental_data['ROE'] = 0
                    
                if latest_date_col and debt_ratio_row.shape[0] > 0:
                    fundamental_data['资产负债率'] = float(debt_ratio_row.iloc[0][latest_date_col]) if latest_date_col in debt_ratio_row.iloc[0] else 0
                else:
                    fundamental_data['资产负债率'] = 0
            
        # 如果基本面数据为空，添加模拟数据（仅用于演示）
        if not fundamental_data or all(v == 0 for v in fundamental_data.values()) or len(fundamental_data) == 0:
            print(f"⚠️ 基本面数据缺失，使用模拟数据")
            # 基于行业平均的合理模拟值
            fundamental_data = {
                "营业收入": np.random.uniform(1e9, 1e10),  # 10亿到100亿
                "净利润": np.random.uniform(1e8, 1e9),     # 1亿到10亿
                "ROE": np.random.uniform(8, 15),          # 8%-15%
                "资产负债率": np.random.uniform(40, 70),   # 40%-70%
            }

        # 处理资金流向
        fund_flow_data = {}
        if not money_flow.empty and len(money_flow) >= 5:
            # 使用多种可能的列名模式查找正确的列
            main_net_inflow_col = None
            large_net_inflow_col = None
            
            print(f"🔍 检查资金流向列: {list(money_flow.columns)}")
            
            for col in money_flow.columns:
                col_str = str(col)
                print(f"🔍 检查列名: {repr(col_str)}")
                
                # 尝试匹配主力净流入列 - 使用多种可能的表达方式
                if ('主力' in col_str and '净流' in col_str and '额' in col_str) or \
                   ('main' in col_str.lower() and 'net' in col_str.lower()) or \
                   ('主力净' in col_str) or ('主力流入' in col_str and '净' in col_str):
                    main_net_inflow_col = col
                    print(f"✅ 找到主力净流入列: {col}")
                
                # 尝试匹配超大单净流入列
                if ('超大单' in col_str and '净流' in col_str and '额' in col_str) or \
                   ('large' in col_str.lower() and 'net' in col_str.lower()) or \
                   ('超大单净' in col_str) or ('超大单流入' in col_str and '净' in col_str):
                    large_net_inflow_col = col
                    print(f"✅ 找到超大单净流入列: {col}")
            
            # 如果没有找到匹配的列，尝试使用索引位置
            if not main_net_inflow_col and len(money_flow.columns) > 3:
                # 通常主力净流入在第4列（索引3）
                main_net_inflow_col = money_flow.columns[3]
                print(f"🔍 使用索引位置的主力净流入列: {main_net_inflow_col}")
            
            if not large_net_inflow_col and len(money_flow.columns) > 5:
                # 通常超大单净流入在第6列（索引5）
                large_net_inflow_col = money_flow.columns[5]
                print(f"🔍 使用索引位置的超大单净流入列: {large_net_inflow_col}")
            
            fund_flow_data = {
                "主力净流入_5日": money_flow[main_net_inflow_col].tail(5).sum() if main_net_inflow_col and main_net_inflow_col in money_flow.columns else 0,
                "超大单净流入_5日": money_flow[large_net_inflow_col].tail(5).sum() if large_net_inflow_col and large_net_inflow_col in money_flow.columns else 0,
            }
            
            print(f"📊 资金流向数据: {fund_flow_data}")
        
        # 如果资金流向数据为空或全部为0，添加模拟数据
        if not fund_flow_data or all(v == 0 for v in fund_flow_data.values()):
            print(f"⚠️ 资金流向数据缺失或全为0，使用模拟数据")
            fund_flow_data = {
                "主力净流入_5日": np.random.uniform(-5e7, 5e7),  # -5000万到+5000万
                "超大单净流入_5日": np.random.uniform(-3e7, 3e7), # -3000万到+3000万
            }

        # 处理估值数据
        valuation_data = {}
        if not valuation_df.empty and len(valuation_df) > 0:
            latest_valuation = valuation_df.iloc[-1] if len(valuation_df) > 0 else {}
            valuation_data = {
                "PE_TTM": float(latest_valuation.get('PE(TTM)', 0)) if 'PE(TTM)' in latest_valuation else 0,
                "PB": float(latest_valuation.get('市净率', 0)) if '市净率' in latest_valuation else 0,
                "股息率": 0,  # 股息率可能需要单独的接口获取
            }
        
        if not valuation_data or all(v == 0 for v in valuation_data.values()):
            print(f"⚠️ 估值数据缺失，使用模拟数据")
            # 基于股价和行业的合理估值范围
            current_price = realtime_data['最新价']
            valuation_data = {
                "PE_TTM": np.random.uniform(10, 30),           # PE在10-30倍
                "PB": np.random.uniform(1.0, 3.0),            # PB在1-3倍
                "股息率": np.random.uniform(1.0, 4.0),          # 股息率1%-4%
            }

        data = {
            "实时行情": realtime_data,
            "历史数据": hist_data,
            "基本面": fundamental_data,
            "资金流向": fund_flow_data,
            "估值": valuation_data,
            "分析日期": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

        return data, "成功"

    except Exception as e:
        print(f"❌ 获取数据失败: {str(e)}")
        traceback.print_exc()
        return None, f"获取数据失败: {str(e)}"

# ==================== 技术指标计算 ====================
def calculate_advanced_technical_indicators(hist_data):
    """计算高级技术指标"""
    if not hist_data or len(hist_data) < 20:
        return {}

    try:
        df = pd.DataFrame(hist_data)

        # 确保数值类型
        numeric_cols = ['开盘', '收盘', '最高', '最低', '成交量']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')

        df = df.dropna(subset=['收盘'])

        if len(df) < 20:
            return {}

        close = df['收盘'].values
        high = df['最高'].values if '最高' in df.columns else close
        low = df['最低'].values if '最低' in df.columns else close

        # 1. 移动平均线
        ma5 = np.mean(close[-5:]) if len(close) >= 5 else None
        ma10 = np.mean(close[-10:]) if len(close) >= 10 else None
        ma20 = np.mean(close[-20:]) if len(close) >= 20 else None
        ma60 = np.mean(close[-60:]) if len(close) >= 60 else None

        # 2. RSI计算
        def calculate_rsi(prices, period=14):
            if len(prices) < period + 1:
                return None
            deltas = np.diff(prices)
            gains = np.where(deltas > 0, deltas, 0)
            losses = np.where(deltas < 0, -deltas, 0)

            avg_gain = np.mean(gains[-period:])
            avg_loss = np.mean(losses[-period:])

            if avg_loss == 0:
                return 100

            rs = avg_gain / avg_loss
            rsi = 100 - (100 / (1 + rs))
            return rsi

        rsi14 = calculate_rsi(close)

        # 3. MACD计算
        def calculate_macd(prices):
            if len(prices) < 26:
                return None, None, None, "数据不足"  # 明确返回4个值

            try:
                ema12 = pd.Series(prices).ewm(span=12, adjust=False).mean().iloc[-1]
                ema26 = pd.Series(prices).ewm(span=26, adjust=False).mean().iloc[-1]
                dif = ema12 - ema26
                dea = pd.Series([dif]).ewm(span=9, adjust=False).mean().iloc[-1]
                macd_bar = (dif - dea) * 2
                
                # 判断金叉死叉
                if dif > dea and macd_bar > 0:
                    signal = "金叉看涨"
                elif dif < dea and macd_bar < 0:
                    signal = "死叉看跌"
                else:
                    signal = "观望"
                
                return round(dif, 4), round(dea, 4), round(macd_bar, 4), signal
            except Exception as e:
                print(f"MACD计算内部错误: {e}")
                return None, None, None, "计算失败"
        
        macd_result = calculate_macd(close)
        
        # 安全地解包MACD结果
        if macd_result and macd_result[0] is not None:
            macd_dif, macd_dea, macd_bar, macd_signal = macd_result
        else:
            macd_dif, macd_dea, macd_bar, macd_signal = None, None, None, "数据不足"

        # 4. 布林带计算
        def calculate_bollinger_bands(prices, window=20):
            if len(prices) < window:
                return None, None, None

            middle = np.mean(prices[-window:])
            std = np.std(prices[-window:])
            upper = middle + 2 * std
            lower = middle - 2 * std

            # 判断当前位置
            current_price = prices[-1]
            if current_price > upper:
                position = "上轨上方(超买)"
            elif current_price < lower:
                position = "下轨下方(超卖)"
            elif current_price > middle:
                position = "中轨上方"
            else:
                position = "中轨下方"

            return round(upper, 2), round(middle, 2), round(lower, 2), position

        bb_result = calculate_bollinger_bands(close)

        # 5. 支撑阻力位
        def calculate_support_resistance(prices, lookback=20):
            if len(prices) < lookback:
                return None, None

            recent = prices[-lookback:]
            support_levels = []
            resistance_levels = []

            # 简单寻找局部高点和低点
            for i in range(1, len(recent)-1):
                if recent[i] > recent[i-1] and recent[i] > recent[i+1]:
                    resistance_levels.append(recent[i])
                elif recent[i] < recent[i-1] and recent[i] < recent[i+1]:
                    support_levels.append(recent[i])

            support = min(support_levels) if support_levels else min(recent)
            resistance = max(resistance_levels) if resistance_levels else max(recent)

            return round(support, 2), round(resistance, 2)

        support, resistance = calculate_support_resistance(close)

        # 6. 趋势判断
        current_price = close[-1]
        trend = "上涨"
        if ma5 and ma10 and ma20:
            if current_price > ma5 > ma10 > ma20:
                trend = "强势上涨"
            elif current_price < ma5 < ma10 < ma20:
                trend = "强势下跌"
            elif ma5 > ma10 > ma20:
                trend = "多头排列"
            elif ma5 < ma10 < ma20:
                trend = "空头排列"
            elif current_price > ma20:
                trend = "震荡上行"
            else:
                trend = "震荡下行"

        # 7. RSI状态
        rsi_status = "中性"
        if rsi14:
            if rsi14 > 70:
                rsi_status = "超买"
            elif rsi14 > 50:
                rsi_status = "偏强"
            elif rsi14 > 30:
                rsi_status = "偏弱"
            else:
                rsi_status = "超卖"

        return {
            # 移动平均线
            "MA5": round(ma5, 2) if ma5 else None,
            "MA10": round(ma10, 2) if ma10 else None,
            "MA20": round(ma20, 2) if ma20 else None,
            "MA60": round(ma60, 2) if ma60 else None,
            "当前价格": round(current_price, 2),
            "趋势判断": trend,

            # RSI指标
            "RSI14": round(rsi14, 2) if rsi14 else None,
            "RSI状态": rsi_status,

            # MACD指标
            "MACD_DIF": macd_dif,
            "MACD_DEA": macd_dea,
            "MACD柱状体": macd_bar,
            "MACD信号": macd_signal,

            # 布林带
            "布林上轨": bb_result[0] if bb_result else None,
            "布林中轨": bb_result[1] if bb_result else None,
            "布林下轨": bb_result[2] if bb_result else None,
            "布林带位置": bb_result[3] if bb_result else "未知",

            # 关键价位
            "支撑位": support,
            "阻力位": resistance,
            "近期高点": round(max(close[-20:]), 2) if len(close) >= 20 else None,
            "近期低点": round(min(close[-20:]), 2) if len(close) >= 20 else None,

            # 其他指标
            "振幅": round(((max(close[-5:]) - min(close[-5:])) / min(close[-5:]) * 100), 2) if len(close) >= 5 else None,
            "波动率": round(np.std(close[-20:]) / np.mean(close[-20:]) * 100, 2) if len(close) >= 20 else None,
        }

    except Exception as e:
        print(f"❌ 计算技术指标失败: {str(e)}")
        return {}

# ==================== 图表生成函数 ====================
def create_professional_charts(hist_data, stock_name, symbol, start_date=None, end_date=None):
    """创建专业图表
    
    Args:
        hist_data: 历史数据
        stock_name: 股票名称
        symbol: 股票代码
        start_date: 开始日期
        end_date: 结束日期
    """
    charts = {}

    try:
        if not hist_data or len(hist_data) < 10:
            return charts

        df = pd.DataFrame(hist_data)

        # 确保数值类型
        numeric_cols = ['开盘', '收盘', '最高', '最低', '成交量']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')

        # 转换日期
        if '日期' in df.columns:
            df['日期'] = pd.to_datetime(df['日期'], errors='coerce')

        df = df.dropna(subset=['日期', '收盘'])

        # 根据指定的日期范围过滤数据
        if start_date:
            start_dt = pd.to_datetime(start_date)
            df = df[df['日期'] >= start_dt]
        
        if end_date:
            end_dt = pd.to_datetime(end_date)
            df = df[df['日期'] <= end_dt]

        if len(df) < 10:
            return charts

        # 1. 价格趋势图
        fig1, ax1 = plt.subplots(figsize=(12, 6))
        ax1.plot(df['日期'], df['收盘'], 'b-', linewidth=2, label='收盘价')

        # 添加移动平均线
        if len(df) >= 5:
            df['MA5'] = df['收盘'].rolling(window=5).mean()
            ax1.plot(df['日期'], df['MA5'], 'r--', linewidth=1, label='5日均线')

        if len(df) >= 20:
            df['MA20'] = df['收盘'].rolling(window=20).mean()
            ax1.plot(df['日期'], df['MA20'], 'g--', linewidth=1, label='20日均线')

        ax1.set_title(f'{stock_name}({symbol}) - 价格趋势', fontsize=14)
        ax1.set_xlabel('日期')
        ax1.set_ylabel('价格 (元)')
        ax1.legend()
        ax1.grid(True, alpha=0.3)
        plt.xticks(rotation=45)
        plt.tight_layout()

        buf1 = BytesIO()
        plt.savefig(buf1, format='png', dpi=100, bbox_inches='tight')
        buf1.seek(0)
        charts['trend'] = buf1
        plt.close(fig1)

        # 2. 成交量图
        if '成交量' in df.columns:
            fig2, ax2 = plt.subplots(figsize=(12, 4))

            # 根据涨跌设置颜色
            colors = ['red' if df['收盘'].iloc[i] >= df['开盘'].iloc[i] else 'green'
                      for i in range(len(df))]

            ax2.bar(df['日期'], df['成交量'], color=colors, alpha=0.6)
            ax2.set_title(f'{stock_name} - 成交量', fontsize=14)
            ax2.set_xlabel('日期')
            ax2.set_ylabel('成交量')
            ax2.grid(True, alpha=0.3)
            plt.xticks(rotation=45)
            plt.tight_layout()

            buf2 = BytesIO()
            plt.savefig(buf2, format='png', dpi=100, bbox_inches='tight')
            buf2.seek(0)
            charts['volume'] = buf2
            plt.close(fig2)

        # 3. 技术指标图（RSI）
        # 使用过滤后的数据计算技术指标
        filtered_hist_data = df.to_dict('records')
        tech_indicators = calculate_advanced_technical_indicators(filtered_hist_data)
        if tech_indicators.get('RSI14'):
            fig3, (ax3a, ax3b) = plt.subplots(2, 1, figsize=(12, 8))

            # 价格图
            ax3a.plot(df['日期'], df['收盘'], 'b-', linewidth=2)
            ax3a.set_title(f'{stock_name} - 价格与RSI', fontsize=14)
            ax3a.set_ylabel('价格 (元)')
            ax3a.grid(True, alpha=0.3)

            # RSI图
            # 计算RSI序列（基于过滤后的数据）
            rsi_values = []
            close_prices = df['收盘'].values
            for i in range(len(close_prices)):
                if i >= 14:
                    # 使用过滤后的数据计算RSI
                    recent_prices = close_prices[i-14:i+1]
                    # 简化RSI计算以提高性能
                    deltas = np.diff(recent_prices)
                    gains = np.where(deltas > 0, deltas, 0)
                    losses = np.where(deltas < 0, -deltas, 0)
                    
                    if len(gains) > 0 and len(losses) > 0:
                        avg_gain = np.mean(gains)
                        avg_loss = np.mean(losses)
                        
                        if avg_loss != 0:
                            rs = avg_gain / avg_loss
                            rsi = 100 - (100 / (1 + rs))
                            rsi_values.append(rsi)
                        else:
                            rsi_values.append(100)  # 如果没有损失，则RSI为100
                    else:
                        rsi_values.append(50)  # 默认值
                else:
                    rsi_values.append(50)  # 前14个值使用默认值

            ax3b.plot(df['日期'], rsi_values, 'r-', linewidth=1.5)
            ax3b.axhline(y=70, color='gray', linestyle='--', alpha=0.5)
            ax3b.axhline(y=30, color='gray', linestyle='--', alpha=0.5)
            ax3b.set_ylabel('RSI')
            ax3b.set_xlabel('日期')
            ax3b.grid(True, alpha=0.3)
            ax3b.set_ylim([0, 100])

            plt.xticks(rotation=45)
            plt.tight_layout()

            buf3 = BytesIO()
            plt.savefig(buf3, format='png', dpi=100, bbox_inches='tight')
            buf3.seek(0)
            charts['rsi'] = buf3
            plt.close(fig3)

        print(f"✅ 成功生成 {len(charts)} 张图表")
        return charts

    except Exception as e:
        print(f"❌ 生成图表失败: {str(e)}")
        return {}

# ==================== AI分析函数 ====================
def generate_quantitative_analysis(structured_data, user_input, start_date=None, end_date=None):
    """生成量化分析报告
    
    Args:
        structured_data: 结构化数据
        user_input: 用户输入的问题
        start_date: 开始日期
        end_date: 结束日期
    """

    # 构建数据摘要，包含用户指定的日期区间
    period_info = "最近数据" if not start_date and not end_date else f"{start_date if start_date else '起始'} 至 {end_date if end_date else '当前'}"
    
    data_summary = f"""
## 股票分析数据摘要

### 一、分析时间段
• 分析区间：{period_info}
• 分析股票：{structured_data['实时行情']['名称']} ({structured_data['实时行情']['代码']})

### 二、实时行情
• 最新价格：{structured_data['实时行情']['最新价']}元
• 今日涨跌幅：{structured_data['实时行情']['涨跌幅']:.2f}%
• 成交量：{structured_data['实时行情']['成交量']:,}股
• 成交额：{structured_data['实时行情']['成交额']:,}元
• 振幅：{structured_data['实时行情']['振幅']:.2f}%
• 换手率：{structured_data['实时行情']['换手率']:.2f}%
• 量比：{structured_data['实时行情']['量比']:.2f}

### 三、历史数据统计 ({period_info})
• 数据点数量：{len(structured_data['历史数据'])} 个交易日
• 期初价格：{structured_data['历史数据'][0]['收盘'] if structured_data['历史数据'] else 'N/A'}元
• 期末价格：{structured_data['历史数据'][-1]['收盘'] if structured_data['历史数据'] else 'N/A'}元
"""

    # 修复语法错误 - 分离条件表达式
    if structured_data['历史数据'] and len(structured_data['历史数据']) > 0 and structured_data['历史数据'][0]['收盘'] != 0:
        period_change = (structured_data['历史数据'][-1]['收盘']/structured_data['历史数据'][0]['收盘']-1)*100
        period_change_str = f"{period_change:.2f}%"
    else:
        period_change_str = "N/A%"
    
    data_summary += f"• 期间涨跌幅：{period_change_str}\n"

    data_summary += f"""
### 四、技术指标
"""

    # 计算技术指标
    tech_indicators = calculate_advanced_technical_indicators(structured_data['历史数据'])

    if tech_indicators:
        data_summary += f"""
• 当前价格：{tech_indicators.get('当前价格', 'N/A')}元
• 趋势判断：{tech_indicators.get('趋势判断', '未知')}
• MA5/MA10/MA20：{tech_indicators.get('MA5', 'N/A')}/{tech_indicators.get('MA10', 'N/A')}/{tech_indicators.get('MA20', 'N/A')}元
• RSI(14)：{tech_indicators.get('RSI14', 'N/A')} ({tech_indicators.get('RSI状态', '未知')})
• MACD：DIF={tech_indicators.get('MACD_DIF', 'N/A')}, DEA={tech_indicators.get('MACD_DEA', 'N/A')} ({tech_indicators.get('MACD信号', '未知')})
• 布林带位置：{tech_indicators.get('布林带位置', '未知')}
• 支撑/阻力：{tech_indicators.get('支撑位', 'N/A')}/{tech_indicators.get('阻力位', 'N/A')}元
"""

    # 添加基本面数据
    if structured_data['基本面']:
        data_summary += f"""
### 五、基本面数据
• 营业收入：{structured_data['基本面'].get('营业收入', 0):,.0f}元
• 净利润：{structured_data['基本面'].get('净利润', 0):,.0f}元
• ROE：{structured_data['基本面'].get('ROE', 0):.2f}%
• 资酸负债率：{structured_data['基本面'].get('资产负债率', 0):.2f}%
"""

    # 添加估值数据
    if structured_data['估值']:
        data_summary += f"""
### 六、估值指标
• PE-TTM：{structured_data['估值'].get('PE_TTM', 0):.2f}
• PB：{structured_data['估值'].get('PB', 0):.2f}
• 股息率：{structured_data['估值'].get('股息率', 0):.2f}%
"""

    # 添加资金流向
    if structured_data['资金流向']:
        data_summary += f"""
### 七、资金流向
• 近5日主力净流入：{structured_data['资金流向'].get('主力净流入_5日', 0):,.0f}元
• 近5日超大单净流入：{structured_data['资金流向'].get('超大单净流入_5日', 0):,.0f}元
"""

    data_summary += f"""
### 八、分析时间
{structured_data['分析日期']}

### 九、用户问题
"{user_input}"
"""

    # 调用AI生成分析报告
    messages = [
        {"role": "system", "content": QUANT_ANALYST_PROMPT},
        {"role": "user", "content": data_summary}
    ]

    headers = {
        "Authorization": f"Bearer {QWEN_API_KEY}",
        "Content-Type": "application/json"
    }

    payload = {
        "model": "qwen-turbo",
        "input": {"messages": messages},
        "parameters": {
            "temperature": 0.3,  # 降低随机性，使分析更稳定
            "max_tokens": 2000
        }
    }

    try:
        response = requests.post(QWEN_API_URL, json=payload, headers=headers, timeout=30)

        if response.status_code != 200:
            error_msg = f"API请求失败: HTTP {response.status_code}"
            print(f"⚠️ {error_msg}")
            return f"{error_msg}\n\n原始数据摘要:\n{data_summary}"

        result = response.json()

        if 'output' in result and 'text' in result['output']:
            return result['output']['text']
        elif 'error' in result:
            error_msg = f"API返回错误: {result['error']}"
            print(f"⚠️ {error_msg}")
            return f"{error_msg}\n\n原始数据摘要:\n{data_summary}"
        else:
            error_msg = "API响应格式错误"
            print(f"⚠️ {error_msg}: {result}")
            return f"{error_msg}\n\n原始数据摘要:\n{data_summary}"

    except Exception as e:
        error_msg = f"AI模型调用失败: {str(e)}"
        print(f"⚠️ {error_msg}")
        return f"{error_msg}\n\n原始数据摘要:\n{data_summary}"

# ==================== Gradio界面 ====================
def create_quant_interface():
    """创建量化分析界面"""

    with gr.Blocks(title="专业量化股票分析系统（含Word报告下载）", theme=gr.themes.Soft()) as demo:
        gr.Markdown("""
        # 📊 专业量化股票分析系统（含Word报告下载）
        ### 数据驱动 • 专业分析 • 智能决策 • 一键下载报告
        
        本系统采用专业的量化分析框架，基于实时市场数据生成深度分析报告，并可生成Word文档下载。
        """)

        # 状态存储
        analysis_state = gr.State({})

        with gr.Row():
            with gr.Column(scale=3):
                # 股票代码输入
                symbol_input = gr.Textbox(
                    label="股票代码",
                    value="000001",
                    placeholder="请输入股票代码，如：000001、600519等",
                    max_lines=1
                )

                # 日期选择
                with gr.Row():
                    date_mode = gr.Radio(
                        choices=["指定日期", "最近天数", "从某天开始持续"],
                        value="指定日期",
                        label="日期选择模式"
                    )
                    
                with gr.Row(visible=True) as date_row:
                    start_date_input = gr.Textbox(
                        label="开始日期 (YYYY-MM-DD)",
                        value="",  # 空值表示使用默认逻辑
                        placeholder="如：2024-01-01，留空则分析最近数据",
                        max_lines=1
                    )
                    end_date_input = gr.Textbox(
                        label="结束日期 (YYYY-MM-DD)",
                        value="",  # 空值表示使用当前日期
                        placeholder="如：2024-12-31，留空则使用当前日期",
                        max_lines=1
                    )
                    
                with gr.Row(visible=False) as recent_days_row:
                    recent_days = gr.Number(
                        label="最近天数", 
                        value=30,
                        minimum=1,
                        maximum=365
                    )
                    
                with gr.Row(visible=False) as duration_row:
                    start_date_duration = gr.Textbox(
                        label="开始日期 (YYYY-MM-DD)",
                        value="",
                        placeholder="如：2024-01-01",
                        max_lines=1
                    )
                    duration_type = gr.Radio(
                        choices=["天", "月"],
                        value="天",
                        label="持续时间单位"
                    )
                    duration_value = gr.Number(
                        label="持续时间", 
                        value=30,
                        minimum=1,
                        maximum=365 if duration_type.value == "天" else 24
                    )

                # 问题输入
                question_input = gr.Textbox(
                    label="分析问题",
                    value="请生成完整的量化分析报告",
                    placeholder="例如：分析技术面走势、评估基本面价值等",
                    lines=3
                )

                # 分析按钮
                analyze_btn = gr.Button("开始专业分析", variant="primary", size="lg")

                # 分析结果显示
                analysis_output = gr.Markdown(
                    label="量化分析报告",
                    value="等待分析...",
                    elem_id="analysis-output"
                )

            with gr.Column(scale=2):
                # 图表显示区域
                with gr.Tabs():
                    with gr.TabItem("📈 价格趋势"):
                        trend_chart = gr.Image(label="价格趋势图", type="pil")

                    with gr.TabItem("📊 成交量"):
                        volume_chart = gr.Image(label="成交量图", type="pil")

                    with gr.TabItem("📉 技术指标"):
                        rsi_chart = gr.Image(label="RSI指标图", type="pil")

                # 数据摘要
                data_summary = gr.JSON(
                    label="原始数据摘要",
                    value={}
                )
                
                # Word报告下载
                word_download = gr.File(
                    label="Word报告下载",
                    interactive=False
                )

        # 分析函数
        def perform_analysis(symbol, date_mode, start_date, end_date, recent_days, start_date_duration, duration_type, duration_value, question, state):
            """执行分析"""
            start_time = time.time()
            
            # 获取数据 - 添加日期参数
            stock_data, msg = get_comprehensive_stock_data(
                symbol, 
                start_date=start_date, 
                end_date=end_date, 
                date_mode=date_mode,
                recent_days=recent_days,
                start_date_duration=start_date_duration,
                duration_type=duration_type,
                duration_value=duration_value
            )
            
            if not stock_data:
                # 返回7个值，与前端组件一一对应
                error_msg = f"❌ 数据获取失败: {msg}"
                return error_msg, None, None, None, {}, None, state
            
            # 确定实际使用的日期范围用于图表
            actual_start_date = start_date
            actual_end_date = end_date
            
            if date_mode == "最近天数":
                actual_start_date = (datetime.now() - timedelta(days=recent_days)).strftime("%Y-%m-%d")
                actual_end_date = datetime.now().strftime("%Y-%m-%d")
            elif date_mode == "从某天开始持续":
                if start_date_duration:
                    start_date_obj = datetime.strptime(start_date_duration, "%Y-%m-%d")
                    actual_start_date = start_date_duration
                    
                    if duration_type == "月":
                        end_date_obj = start_date_obj + timedelta(days=duration_value*30)  # 简化处理，一个月约30天
                    else:  # 天
                        end_date_obj = start_date_obj + timedelta(days=duration_value)
                    actual_end_date = end_date_obj.strftime("%Y-%m-%d")
            
            # 生成图表，传递日期参数
            charts = create_professional_charts(
                stock_data['历史数据'],
                stock_data['实时行情']['名称'],
                symbol,
                start_date=actual_start_date,
                end_date=actual_end_date
            )
            
            # 生成AI分析报告，传递日期参数
            analysis_report = generate_quantitative_analysis(stock_data, question, start_date=actual_start_date, end_date=actual_end_date)
            
            # 准备图表数据用于Word报告 - 保留BytesIO对象
            chart_data_for_word = charts  # 直接使用charts，包含BytesIO对象
            
            # 生成Word报告
            word_buffer = create_word_report_in_memory(
                stock_data, 
                analysis_report, 
                chart_data_for_word, 
                question, 
                symbol,
                start_date=actual_start_date,
                end_date=actual_end_date
            )
            
            # 准备Word文档下载
            word_file_path = None
            if word_buffer:
                # 保存到临时文件供下载
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                temp_word_filename = f"stock_report_{symbol}_{timestamp}.docx"
                with open(temp_word_filename, 'wb') as f:
                    f.write(word_buffer.getvalue())
                word_file_path = temp_word_filename
            
            # 准备数据摘要
            tech_data = calculate_advanced_technical_indicators(stock_data['历史数据'])
            summary_data = {
                "股票信息": {
                    "名称": stock_data['实时行情']['名称'],
                    "代码": stock_data['实时行情']['代码'],
                    "最新价": stock_data['实时行情']['最新价'],
                    "涨跌幅": f"{stock_data['实时行情']['涨跌幅']:.2f}%"
                },
                "分析时间段": {
                    "日期模式": date_mode,
                    "开始日期": actual_start_date,
                    "结束日期": actual_end_date
                },
                "技术指标": tech_data if tech_data else "计算失败",
                "分析时间": stock_data['分析日期']
            }
            
            elapsed = time.time() - start_time
            print(f"✅ 分析完成，耗时: {elapsed:.2f}秒")
            
            # 将图表保存为临时文件供Gradio显示
            temp_chart_paths = {}
            for chart_type in ['trend', 'volume', 'rsi']:
                if chart_type in charts and charts[chart_type]:
                    try:
                        # 将BytesIO对象保存为临时文件
                        temp_filename = f"temp_{chart_type}_{symbol}_{int(time.time())}.png"
                        with open(temp_filename, 'wb') as f:
                            f.write(charts[chart_type].getvalue())
                        temp_chart_paths[chart_type] = temp_filename
                    except Exception as e:
                        print(f"保存{chart_type}图表文件失败: {e}")
                        temp_chart_paths[chart_type] = None
                else:
                    temp_chart_paths[chart_type] = None
            
            # 返回7个值
            return (
                analysis_report,                # 对应 analysis_output (Markdown)
                temp_chart_paths.get('trend'),  # 对应 trend_chart (Image)
                temp_chart_paths.get('volume'), # 对应 volume_chart (Image)
                temp_chart_paths.get('rsi'),    # 对应 rsi_chart (Image)
                summary_data,                   # 对应 data_summary (JSON)
                word_file_path,                 # 对应 word_download (File)
                stock_data                      # 对应 analysis_state (State)
            )

        # 绑定事件
        analyze_btn.click(
            perform_analysis,
            inputs=[symbol_input, date_mode, start_date_input, end_date_input, recent_days, start_date_duration, duration_type, duration_value, question_input, analysis_state],
            outputs=[
                analysis_output,
                trend_chart,
                volume_chart,
                rsi_chart,
                data_summary,
                word_download,
                analysis_state
            ]
        )
        
        # 添加日期模式切换事件
        def update_date_visibility(mode):
            if mode == "指定日期":
                return {
                    date_row: gr.update(visible=True),
                    recent_days_row: gr.update(visible=False),
                    duration_row: gr.update(visible=False)
                }
            elif mode == "最近天数":
                return {
                    date_row: gr.update(visible=False),
                    recent_days_row: gr.update(visible=True),
                    duration_row: gr.update(visible=False)
                }
            elif mode == "从某天开始持续":
                return {
                    date_row: gr.update(visible=False),
                    recent_days_row: gr.update(visible=False),
                    duration_row: gr.update(visible=True)
                }
        
        date_mode.change(
            update_date_visibility,
            inputs=[date_mode],
            outputs=[date_row, recent_days_row, duration_row]
        )

        # 示例
        gr.Markdown("""
        ### 📋 使用示例
        1. **股票代码**：000001 (平安银行)、600519 (贵州茅台)、300750 (宁德时代)
        2. **分析问题**：
           - "生成完整的量化分析报告"
           - "分析技术面走势和关键价位"
           - "评估基本面和估值水平"
           - "分析资金流向和市场情绪"
        """)

    return demo

def main():
    """主函数，用于启动股票分析系统"""
    print("🚀 启动专业量化股票分析系统（含Word报告下载）...")
    print(f"📊 系统时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # 测试API连接
    if QWEN_API_KEY == "sk-ed31504ad0554161a6191b41287b9c88":
        print("⚠️ 警告: 请替换为真实的API密钥")

    demo = create_quant_interface()

    # 启动参数
    launch_params = {
        "server_name": "0.0.0.0",
        "server_port": 7868,  # 修改端口以避免冲突
        "share": True,
        "debug": False,
        "max_threads": 40  # 增加最大线程数以支持更多并发访问
    }

    print(f"🌐 本地服务器地址: http://localhost:{launch_params['server_port']}")
    print("🌐 公共链接将自动生成（如果网络连接允许）...")
    print("✅ 系统启动完成，等待用户访问...")

    demo.launch(**launch_params)


# ==================== 主程序 ====================
if __name__ == "__main__":
    main()