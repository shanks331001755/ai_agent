#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
股票分析AI智能体 - 简化版Word报告生成
功能：通过命令行参数分析指定股票，使用AI模型生成分析报告，并生成包含图表的Word文档

依赖库安装命令：
    pip install akshare pandas numpy matplotlib python-docx
    或使用特定Python版本：
    py -3.14 -m pip install akshare pandas numpy matplotlib python-docx

使用方法：
    python stock_ai_agent_word_simple.py 000001 -q "请分析这只股票的近期走势"

API配置说明：
    请将 QWEN_API_KEY 替换为您的真实通义千问API密钥

文件生成逻辑：
    1. Word报告：每次运行都会生成一个以股票代码和时间戳命名的docx文件
    2. 报告文件命名格式：{股票代码}_analysis_report_{YYYYMMDD_HHMM}.docx
    3. 文件保存路径：与脚本相同的目录
"""
