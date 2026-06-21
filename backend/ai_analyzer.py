"""
AI Sales Data Analyzer
======================

使用 MiniMax 大模型进行智能销售数据分析。

核心功能：
1. 读取 Excel 文件并解析数据
2. 将结构化数据转换为 LLM 可理解的文本格式
3. 使用预设的 Prompt 模板调用 AI 分析
4. 解析返回的 JSON 结果
5. 生成可视化 HTML 报告

分析维度：
1. 营业额与利润统计（最近30天）
2. 价格区间业绩排名（7个区间）
3. 品类业绩对比（按小类分组）
4. 热销商品 TOP 8
5. 旧料回收明细 & 盈亏分析（珠宝行业专属）
6. 业务洞察与建议（AI增值）

作者: hanyang
日期: 2026-06-21
"""

import json
import logging
import uuid
from datetime import datetime, timedelta
from typing import Any, Dict, Optional, List, Tuple

import pandas as pd

from minimax_client import MiniMaxClient

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class AISalesAnalyzer:
    """AI 驱动的销售数据分析器"""

    # 默认金属价格
    DEFAULT_GOLD_PRICE = 930  # 元/克
    DEFAULT_SILVER_PRICE = 17  # 元/克

    # 价格区间定义
    PRICE_RANGES = [
        {"label": "0-100元", "min": 0, "max": 100},
        {"label": "100-300元", "min": 100, "max": 300},
        {"label": "300-500元", "min": 300, "max": 500},
        {"label": "500-1000元", "min": 500, "max": 1000},
        {"label": "1000-2000元", "min": 1000, "max": 2000},
        {"label": "2000-5000元", "min": 2000, "max": 5000},
        {"label": "5000元以上", "min": 5000, "max": float("inf")},
    ]

    # 列名别名映射（支持不同来源的Excel格式）
    COLUMN_ALIASES = {
        "业绩金额": ["总金额", "金额", "销售金额", "实收金额", "应收金额", "业绩"],
        "利润金额": ["利润", "毛利", "毛利润", "纯利", "净利润"],
        "商品名称": ["商品名", "品名", "商品", "货品名称", "商品条码", "条码", "货号"],
        "小类": ["类别", "分类", "品类", "子类", "子类别", "小类别"],
        "金重": ["金重(g)", "黄金重量", "黄金克重", "金料重量", "Au重量"],
        "银重": ["银重(g)", "白银重量", "白银克重", "银料重量", "Ag重量"],
        "销售类型": ["类型", "交易类型", "业务类型", "操作类型", "单据类型"],
    }

    def __init__(
        self,
        api_key: str,
        model: str = "MiniMax-M3",
        gold_price: float = DEFAULT_GOLD_PRICE,
        silver_price: float = DEFAULT_SILVER_PRICE,
    ):
        """
        初始化 AI 分析器

        Args:
            api_key: MiniMax API 密钥
            model: 模型名称（默认 MiniMax-M3）
            gold_price: 当前金价（元/克，默认930）
            silver_price: 当前银价（元/克，默认17）
        """
        self.client = MiniMaxClient(api_key=api_key, model=model)
        self.gold_price = gold_price
        self.silver_price = silver_price

        logger.info(
            f"AI 分析器初始化完成 | 金价: ¥{gold_price}/g | 银价: ¥{silver_price}/g"
        )

    def analyze_file(self, file_path: str) -> Dict[str, Any]:
        """
        分析 Excel 文件（主入口）

        Args:
            file_path: Excel 文件路径

        Returns:
            完整的分析结果字典，包含所有维度数据和报告HTML
        """
        logger.info(f"开始分析文件: {file_path}")

        # Step 1: 读取 Excel 数据
        df = self._read_excel(file_path)
        logger.info(f"成功读取 {len(df)} 条记录")

        # Step 2: 转换为文本格式
        data_text = self._format_data_for_ai(df)
        logger.info(f"数据格式化完成 | 文本长度: {len(data_text)} 字符")

        # Step 3: 构建 Prompt
        prompt = self._build_analysis_prompt(df, data_text)
        logger.info(f"Prompt 构建完成 | 长度: {len(prompt)} 字符")

        # Step 4: 调用 AI 分析（获取业务洞察）
        ai_result = self._call_ai_api(prompt)

        # Step 5: 用 Pandas 计算完整的结构化数据（核心数据源）
        pandas_stats = self._calculate_full_stats(df)
        analysis_result = pandas_stats

        # Step 6: 合并AI洞察（如果AI返回了有效数据）
        if ai_result and ai_result.get("insights"):
            analysis_result["insights"] = ai_result["insights"]
        # 如果Pandas计算的某些字段为空但AI有数据，用AI补充
        for key in ["range_stats", "cat_stats", "prod_stats", "recycle_items"]:
            if not analysis_result.get(key) and ai_result.get(key):
                analysis_result[key] = ai_result[key]

        # Step 7: 生成 HTML 报告
        report_html = self._generate_report(analysis_result)
        analysis_result["report_html"] = report_html

        # Step 8: 记录文件信息
        analysis_result["file_name"] = file_path.split("/")[-1]
        analysis_result["id"] = str(uuid.uuid4())
        analysis_result["created_at"] = datetime.now().isoformat()

        logger.info(f"分析完成 | ID: {analysis_result['id']}")

        return analysis_result

    def _normalize_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        标准化列名：将各种别名统一为标准列名

        Args:
            df: 原始DataFrame

        Returns:
            列名标准化后的DataFrame
        """
        col_mapping = {}
        for std_name, aliases in self.COLUMN_ALIASES.items():
            # 如果标准列名已存在，跳过
            if std_name in df.columns:
                continue
            # 查找别名
            for alias in aliases:
                for col in df.columns:
                    if col.strip() == alias or col.strip().lower() == alias.lower():
                        col_mapping[col] = std_name
                        logger.info(f"  列名映射: '{col}' → '{std_name}'")
                        break
                if std_name in [col_mapping.get(c) for c in df.columns]:
                    break

        if col_mapping:
            df = df.rename(columns=col_mapping)
            logger.info(f"已映射 {len(col_mapping)} 个列名: {col_mapping}")

        return df

    def _read_excel(self, file_path: str) -> pd.DataFrame:
        """
        读取 Excel 文件

        支持格式：.xls (xlrd), .xlsx (openpyxl)
        自动检测编码和表头行

        Args:
            file_path: 文件路径

        Returns:
            Pandas DataFrame
        """
        try:
            # 尝试读取 xlsx
            if file_path.endswith(".xlsx"):
                df = pd.read_excel(file_path, engine="openpyxl")
            else:
                # 尝试 xls
                df = pd.read_excel(file_path, engine="xlrd")

            # 清理列名（去除空格）
            df.columns = [str(col).strip() for col in df.columns]

            # 标准化列名映射
            df = self._normalize_columns(df)
            logger.info(f"标准化后列名: {list(df.columns)}")

            # 标准化日期列
            date_cols = ["销售时间", "日期", "time", "date", "Date"]
            for col in date_cols:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col], errors="coerce")
                    break

            return df

        except Exception as e:
            raise Exception(f"读取 Excel 文件失败: {e}")

    def _format_data_for_ai(self, df: pd.DataFrame) -> str:
        """
        将 DataFrame 格式化为适合 LLM 处理的文本

        采用表格化格式，保留关键列，限制样本数量以控制 Token 消耗

        Args:
            df: 原始 DataFrame

        Returns:
            格式化的文本字符串
        """
        # 选择关键列
        key_columns = [
            "销售时间",
            "销售类型",
            "小类",
            "商品名称",
            "业绩金额",
            "利润金额",
            "金重",
            "银重",
        ]

        available_cols = [col for col in key_columns if col in df.columns]

        if not available_cols:
            # 如果标准列名不匹配，使用所有列
            available_cols = list(df.columns)[:10]  # 最多取前10列

        # 限制显示数量（控制Token消耗）
        max_rows = min(len(df), 500)  # 最多显示500条
        display_df = df.head(max_rows)[available_cols].copy()

        # 格式化为文本表格
        lines = []
        lines.append("=" * 80)
        lines.append("销售数据明细")
        lines.append(f"总记录数: {len(df)} 条")
        lines.append(f"展示记录数: {max_rows} 条")
        lines.append("=" * 80)
        lines.append("")

        # 表头
        header = " | ".join(available_cols)
        lines.append(header)
        lines.append("-" * len(header))

        # 数据行
        for idx, row in display_df.iterrows():
            values = [str(row.get(col, ""))[:50] for col in available_cols]  # 截断长文本
            line = " | ".join(values)
            lines.append(line)

        lines.append("")
        lines.append("=" * 80)
        lines.append("数据列说明:")
        for col in available_cols:
            sample_val = df[col].iloc[0] if len(df) > 0 else "N/A"
            lines.append(f"  - {col}: 示例值 '{sample_val}'")
        lines.append("=" * 80)

        result_text = "\n".join(lines)
        logger.info(f"数据已格式化 | 列数: {len(available_cols)} | 行数: {max_rows}")

        return result_text

    def _build_analysis_prompt(self, df: pd.DataFrame, data_text: str) -> str:
        """
        构建完整的分析 Prompt

        整合系统提示词、用户数据、分析任务指令

        Args:
            df: 原始 DataFrame（用于获取统计信息）
            data_text: 已格式化的数据文本

        Returns:
            完整的 Prompt 字符串
        """
        # 统计信息摘要（帮助AI理解数据概况）
        stats_summary = self._get_data_summary(df)

        # 构建完整 Prompt
        prompt = """# 角色设定
你是一个资深的珠宝零售行业数据分析师，拥有10年以上的珠宝行业数据分析经验。
你精通Excel数据处理、统计分析、商业洞察。请基于提供的数据进行专业的多维度分析。

# 数据概览
{stats_summary}

# 销售明细数据
{data_text}

# 参数设置
- **金价**: ¥{self.gold_price}/g （用于计算回收金料当前市值）
- **银价**: ¥{self.silver_price}/g （用于计算回收银料当前市值）

# 分析任务（请按顺序执行以下6个维度分析）

## 维度1：营业额与利润统计
- 统计范围：以数据中最新的日期为基准，向前推算30天
- 输出：
  - total_revenue: 总营业额（业绩金额合计）
  - total_profit: 总利润（利润金额合计）
  - sales_count: 销售笔数（业绩金额>0的记录数）
  - time_range: {{start: 开始日期, end: 结束日期}}

## 维度2：价格区间业绩排名
将业绩金额>0的记录按以下7个价格区间分组：

| 区间标签 | 金额范围 |
|---------|---------|
| 0-100元 | [0, 100] |
| 100-300元 | (100, 300] |
| 300-500元 | (300, 500] |
| 500-1000元 | (500, 1000] |
| 1000-2000元 | (1000, 2000] |
| 2000-5000元 | (2000, 5000] |
| 5000元以上 | (5000, +∞) |

每个区间输出：
- count: 交易笔数
- total: 业绩金额合计
- avg: 平均客单价
- transactions: 明细列表（每项包含：日期、商品名称、小类、金额）【仅保留前20条】
按总金额降序排列

## 维度3：品类业绩对比
- 按「小类」字段分组
- 输出每个品类的：name(品类名), count(笔数), total(业绩合计)
- 按total降序排列

## 维度4：热销商品 TOP 8
- 按「商品名称」分组
- 输出前8名的：name(商品名), count(件数), total(业绩合计)
- 按total降序排列

## 维度5：旧料回收明细与盈亏分析 ⭐
筛选条件：销售类型 == '回收'

每条记录处理逻辑：
1. **金属类型判断**：
   - 商品名称包含['金','AU','au','素圈'] → 黄金(gold)
   - 否则 → 白银(silver)

2. **重量提取**：
   - 黄金：取「金重」字段的绝对值（如果该列为空则尝试从商品名提取数字）
   - 白银：取「银重」字段的绝对值

3. **支付金额**：取「业绩金额」的绝对值

4. **当前市值**：
   - 金料: 重量 × ¥{self.gold_price}
   - 银料: 重量 × ¥{self.silver_price}

5. **单笔盈亏**：当前市值 - 支付金额（正=盈利，负=亏损）

输出：
- recycle_items: 每笔回收的明细列表（date, name, metal, weight, paid, value, pnl, price）
- recycle_total_paid: 总支出
- recycle_total_value: 当前总市值
- recycle_total_pnl: 整体盈亏
- recycle_gold_g / _paid / _value / _count: 金料汇总
- recycle_silver_g / _paid / _value / _count: 银料汇总
- gold_price / silver_price: 使用的单价

## 维度6：业务洞察与建议（AI增值）
基于以上5个维度的分析结果，输出3-5条简短的业务洞察：
1. 最畅销的价格区间及占比
2. 主要贡献品类
3. 回收业务整体盈亏情况
4. 异常数据提示（如有）
5. 1-2条改进建议

将此部分放在 insights 数组中，每项为一段文字。

# 输出要求
请严格按照以下 JSON 格式输出结果（不要添加任何额外文字或markdown标记）：

{{
  "total_revenue": 营业总额,
  "total_profit": 总利润,
  "sales_count": 销售笔数,
  "time_range": {{"start": "...", "end": "..."}},
  "range_stats": [{{"range": "0-100元", "count": N, "total": X, "avg": Y}}],
  "range_details": {{"0-100元": {{"transactions": [...], ...}}},
  "cat_stats": [{{"name": "...", "count": N, "total": X}}],
  "prod_stats": [{{"name": "...", "count": N, "total": X}}],
  "recycle_items": [...],
  "recycle_total_paid": X,
  "recycle_total_value": X,
  "recycle_total_pnl": X,
  "recycle_gold_g": X,
  "recycle_gold_paid": X,
  "recycle_gold_value": X,
  "recycle_gold_count": N,
  "recycle_silver_g": X,
  "recycle_silver_paid": X,
  "recycle_silver_value": X,
  "recycle_silver_count": N,
  "gold_price": {self.gold_price},
  "silver_price": {self.silver_price},
  "insights": ["洞察1", "洞察2", ...]
}}

# 注意事项
1. 所有金额保留2位小数
2. 所有重量保留3位小数
3. 日期格式统一为 "YYYY-MM-DD HH:mm"
4. 如果某个维度没有数据（如无回收记录），对应字段设为null或空数组
5. 请确保输出是合法的JSON格式，不要包含任何非JSON字符
6. transactions数组每个元素最多保留20条记录以控制长度
"""

        return prompt

    def _get_data_summary(self, df: pd.DataFrame) -> str:
        """
        生成数据摘要信息

        Args:
            df: DataFrame

        Returns:
            数据摘要文本
        """
        lines = []

        # 基本信息
        lines.append(f"- 总记录数: {len(df)} 行")

        # 日期范围
        date_col = None
        for candidate in ["销售时间", "日期", "time", "date"]:
            if candidate in df.columns and pd.api.types.is_datetime64_any_dtype(df[candidate]):
                date_col = candidate
                break

        if date_col:
            valid_dates = df[date_col].dropna()
            if len(valid_dates) > 0:
                min_date = valid_dates.min().strftime("%Y-%m-%d %H:%M")
                max_date = valid_dates.max().strftime("%Y-%m-%d %H:%M")
                lines.append(f"- 时间跨度: {min_date} ~ {max_date}")
            else:
                lines.append("- 时间跨度: 无法解析日期")

        # 列信息
        lines.append(f"- 列数: {len(df.columns)}")
        lines.append(f"- 列名: {', '.join(df.columns.tolist()[:8])}...")  # 只显示前8列

        # 销售类型分布
        if "销售类型" in df.columns:
            type_counts = df["销售类型"].value_counts().head(5).to_dict()
            type_str = ", ".join([f"{k}({v})" for k, v in type_counts.items()])
            lines.append(f"- 销售类型分布: {type_str}")

        # 小类分布
        if "小类" in df.columns:
            cat_counts = df["小类"].value_counts().head(8).to_dict()
            cat_str = ", ".join([f"{k}({v})" for k, v in cat_counts.items()])
            lines.append(f"- 品类分布(TOP8): {cat_str}")

        # 业绩金额统计
        if "业绩金额" in df.columns:
            revenue_series = pd.to_numeric(df["业绩金额"], errors="coerce").dropna()
            if len(revenue_series) > 0:
                lines.append(f"- 业绩金额合计: ¥{revenue_series.sum():,.2f}")
                lines.append(f"- 平均单笔: ¥{revenue_series.mean():,.2f}")
                lines.append(f"- 最大值: ¥{revenue_series.max():,.2f}")

        return "\n".join(lines)

    def _call_ai_api(self, prompt: str) -> Dict[str, Any]:
        """
        调用 MiniMax API 进行分析

        Args:
            prompt: 完整的分析 Prompt

        Returns:
            解析后的 JSON 结果字典
        """
        system_prompt = """你是一个专业的数据分析助手。你只输出JSON格式的分析结果，不要输出任何其他内容。
请确保你的回复是完整、有效的JSON，可以直接被json.loads()解析。
如果数据中某些维度没有相关记录，对应字段设置为null或[]。"""

        logger.info("正在调用 MiniMax API 进行智能分析...")
        start_time = datetime.now()

        try:
            result = self.client.chat_with_json(
                prompt=prompt,
                system_prompt=system_prompt,
                temperature=0.3,  # 使用较低温度确保输出稳定
            )
        except Exception as e:
            logger.error(f"AI 分析失败: {e}")
            # 返回空结果而不是抛出异常，让系统可以继续运行
            result = self._get_empty_result()

        elapsed = (datetime.now() - start_time).total_seconds()
        logger.info(f"AI 分析完成 | 耗时: {elapsed:.1f}s | 返回键数: {len(result)}")

        return result

    def _calculate_full_stats(self, df: pd.DataFrame) -> Dict[str, Any]:
        """
        使用 Pandas 计算完整的数据统计（核心数据源）

        计算报告模板所需的所有字段，确保数据准确完整

        Args:
            df: DataFrame

        Returns:
            完整的统计字典
        """
        stats = {
            "generate_time": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "gold_price": self.gold_price,
            "silver_price": self.silver_price,
        }

        # ========== 日期处理 ==========
        date_col = None
        for candidate in ["销售时间", "日期", "time", "date"]:
            if candidate in df.columns:
                try:
                    df[candidate] = pd.to_datetime(df[candidate], errors="coerce")
                    date_col = candidate
                    break
                except Exception:
                    continue

        if not date_col:
            logger.warning("未找到日期列，使用全部数据")
            recent_df = df.copy()
            stats["time_range"] = {"start": "未知", "end": "未知"}
        else:
            valid_dates = df[date_col].dropna()
            if len(valid_dates) > 0:
                max_date = valid_dates.max()  # 最晚日期（基准日）
                # 近30天：从基准日往前推30天（不包含第30天当天）
                # 例如 06-18 为基准，则取 (05-19, 06-18]
                min_date = max_date - timedelta(days=30)
                stats["time_range"] = {
                    "start": (min_date + timedelta(days=1)).strftime("%Y-%m-%d"),
                    "end": max_date.strftime("%Y-%m-%d"),
                }
                recent_df = df[(df[date_col] > min_date) & (df[date_col] <= max_date)].copy()
                logger.info(
                    f"日期筛选: {stats['time_range']['start']} ~ {stats['time_range']['end']} "
                    f"| 筛选后 {len(recent_df)}/{len(df)} 条"
                )
            else:
                recent_df = df.copy()
                stats["time_range"] = {"start": "未知", "end": "未知"}

        # 确保数值列是数字类型
        for num_col in ["业绩金额", "利润金额", "销售金额", "金重", "银重"]:
            if num_col in recent_df.columns:
                recent_df[num_col] = pd.to_numeric(recent_df[num_col], errors="coerce")

        # ========== 有效业绩列 ==========
        # 部分记录(如赠品、玉器)的业绩金额=0但销售金额>0
        # 参考报告对这类记录使用「销售金额」作为有效业绩
        def _effective_revenue(row):
            yj = row.get("业绩金额", 0) or 0
            xs = row.get("销售金额", 0) or 0
            return yj if yj > 0 else max(xs, 0)

        recent_df["eff_rev"] = recent_df.apply(_effective_revenue, axis=1)

        # ========== 维度1：基础KPI ==========
        eff_rev_series = recent_df["eff_rev"]
        profit_series = recent_df["利润金额"].dropna() if "利润金额" in recent_df.columns else pd.Series([])

        # 营业总额 = 有效业绩 > 0 的合计（含所有类型的正向记录）
        pos_eff = eff_rev_series[eff_rev_series > 0]
        stats["total_revenue"] = round(float(pos_eff.sum()), 2) if len(pos_eff) > 0 else 0
        stats["sales_count"] = int(len(pos_eff))
        stats["sales_revenue"] = stats["total_revenue"]

        # 总利润 = 全部记录的利润金额合计（含退换亏损、赠品成本等，与参考报告一致）
        stats["total_profit"] = round(float(profit_series.sum()), 2) if len(profit_series) > 0 else 0

        logger.info(
            f"基础KPI | 营业总额: ¥{stats['total_revenue']:,} | "
            f"总利润: ¥{stats['total_profit']:,} | 笔数: {stats['sales_count']}"
        )

        # ========== 维度2：价格区间业绩排名 ==========
        stats["range_stats"] = []
        stats["range_details"] = {}

        if "eff_rev" in recent_df.columns:
            sales_only = recent_df[recent_df["eff_rev"] > 0].copy()
            for range_def in self.PRICE_RANGES:
                mask = (sales_only["eff_rev"] > range_def["min"]) & (sales_only["eff_rev"] <= range_def["max"])
                range_df = sales_only[mask]
                count = len(range_df)
                total = round(float(range_df["eff_rev"].sum()), 2) if count > 0 else 0
                avg = round(total / count, 2) if count > 0 else 0

                # 收集交易明细（最多20条）
                transactions = []
                for _, row in range_df.head(20).iterrows():
                    tx_date = row.get(date_col, "")
                    if hasattr(tx_date, "strftime"):
                        tx_date = tx_date.strftime("%Y-%m-%d %H:%M")
                    transactions.append({
                        "date": str(tx_date),
                        "name": str(row.get("商品名称", ""))[:30],
                        "cat": str(row.get("小类", ""))[:20],
                        "amount": round(float(row.get("eff_rev", 0)), 2),
                    })

                stats["range_stats"].append({
                    "range": range_def["label"],
                    "count": count,
                    "total": total,
                    "avg": avg,
                })
                stats["range_details"][range_def["label"]] = {
                    "total_amount": total,
                    "count": count,
                    "transactions": transactions,
                }

            # 按总金额降序排列
            stats["range_stats"].sort(key=lambda x: x["total"], reverse=True)

            # 计算占比（基于全部正向业绩）
            total_count = sum(r["count"] for r in stats["range_stats"])
            total_rev = sum(r["total"] for r in stats["range_stats"])
            for r in stats["range_stats"]:
                r["count_pct"] = round(r["count"] / total_count * 100, 1) if total_count > 0 else 0
                r["revenue_pct"] = round(r["total"] / total_rev * 100, 1) if total_rev > 0 else 0

        logger.info(f"价格区间 | {len(stats['range_stats'])} 个区间有数据")

        # ========== 维度3：品类业绩对比 ==========
        stats["cat_stats"] = []
        if "小类" in recent_df.columns and "eff_rev" in recent_df.columns:
            cat_grouped = recent_df.groupby("小类")["eff_rev"].agg(["count", "sum"]).reset_index()
            cat_grouped.columns = ["name", "count", "total"]
            # 只保留有效业绩>0的品类
            cat_grouped = cat_grouped[cat_grouped["total"] > 0]
            cat_grouped["total"] = cat_grouped["total"].apply(lambda x: round(abs(float(x)), 2))
            cat_grouped = cat_grouped.sort_values("total", ascending=False)
            for _, row in cat_grouped.iterrows():
                stats["cat_stats"].append({
                    "name": str(row["name"])[:20],
                    "count": int(row["count"]),
                    "total": float(row["total"]),
                })
            logger.info(f"品类统计 | {len(stats['cat_stats'])} 个品类")

        # ========== 维度4：热销商品 TOP8 ==========
        stats["prod_stats"] = []
        if "商品名称" in recent_df.columns and "eff_rev" in recent_df.columns:
            prod_grouped = recent_df.groupby("商品名称")["eff_rev"].agg(["count", "sum"]).reset_index()
            prod_grouped.columns = ["name", "count", "total"]
            prod_grouped = prod_grouped[prod_grouped["total"] > 0]
            prod_grouped["total"] = prod_grouped["total"].apply(lambda x: round(abs(float(x)), 2))
            prod_grouped = prod_grouped.sort_values("total", ascending=False).head(8)
            for _, row in prod_grouped.iterrows():
                stats["prod_stats"].append({
                    "name": str(row["name"])[:30],
                    "count": int(row["count"]),
                    "total": float(row["total"]),
                })
            logger.info(f"商品TOP{len(stats['prod_stats'])} | 已统计")

        # ========== 维度5：旧料回收明细与盈亏分析 ==========
        stats.update(self._calculate_recycle_stats(recent_df, date_col))

        # 默认洞察（如果AI没有返回）
        if "insights" not in stats:
            stats["insights"] = self._generate_default_insights(stats)

        return stats

    def _calculate_recycle_stats(self, df: pd.DataFrame, date_col: Optional[str]) -> Dict[str, Any]:
        """
        计算旧料回收相关统计数据

        Args:
            df: 已筛选的DataFrame
            date_col: 日期列名

        Returns:
            回收统计字典
        """
        recycle_stats = {
            "recycle_items": [],
            "recycle_total_paid": 0,
            "recycle_total_value": 0,
            "recycle_total_pnl": 0,
            "recycle_gold_g": 0,
            "recycle_gold_paid": 0,
            "recycle_gold_value": 0,
            "recycle_gold_count": 0,
            "recycle_silver_g": 0,
            "recycle_silver_paid": 0,
            "recycle_silver_value": 0,
            "recycle_silver_count": 0,
        }

        # 筛选回收类型记录
        if "销售类型" not in df.columns:
            return recycle_stats

        recycle_df = df[df["销售类型"] == "回收"].copy()
        if len(recycle_df) == 0:
            logger.info("无旧料回收数据")
            return recycle_stats

        # 确保数值列
        for col in ["业绩金额", "金重", "银重"]:
            if col in recycle_df.columns:
                recycle_df[col] = pd.to_numeric(recycle_df[col], errors="coerce")

        gold_g_total = 0
        gold_paid_total = 0
        gold_value_total = 0
        gold_count = 0
        silver_g_total = 0
        silver_paid_total = 0
        silver_value_total = 0
        silver_count = 0

        def _get_metal_weight(row, metal_type):
            """
            获取回收记录的金属克重。
            兼容POS系统数据特点：银料重量可能写在「金重」或「件重」列（而非「银重」列）。

            Args:
                row: DataFrame 行
                metal_type: "gold" 或 "silver"

            Returns:
                克重（绝对值）
            """
            if metal_type == "gold":
                # 黄金：优先取「金重」，fallback 到「件重」
                w = abs(float(row.get("金重", 0) or 0))
                if w == 0:
                    w = abs(float(row.get("件重", 0) or 0))
                return w
            else:
                # 白银：优先取「银重」，为0时 fallback 到「金重」「件重」（POS常见写法）
                w = abs(float(row.get("银重", 0) or 0))
                if w == 0:
                    w = abs(float(row.get("金重", 0) or 0))
                if w == 0:
                    w = abs(float(row.get("件重", 0) or 0))
                return w

        def _guess_metal_type(row):
            """
            智能判断金属类型（当商品名称为空或无明确关键词时）。
            通过重量列数据特征推断：如果「银重」>0 则为银，否则看「金重」。
            同时参考商品名称关键词。
            """
            name = str(row.get("商品名称", ""))
            # 关键词匹配（优先）
            gold_keywords = ['金', 'AU', 'au', '素圈', '黄金', '足金', 'K金', 'k金']
            silver_keywords = ['银', 'AG', 'ag', '足银', 'S925', 's925', '纯银', '白银', '旧银', '银饰']

            has_gold_kw = any(kw in name for kw in gold_keywords)
            has_silver_kw = any(kw in name for kw in silver_keywords)

            if has_gold_kw and not has_silver_kw:
                return "gold"
            if has_silver_kw and not has_gold_kw:
                return "silver"

            # 名称无法判断 → 用重量列数据推断
            silver_w = float(row.get("银重", 0) or 0)
            gold_w = float(row.get("金重", 0) or 0)

            if silver_w != 0 and gold_w == 0:
                return "silver"
            if gold_w != 0 and silver_w == 0:
                # 有金重无银重：需要进一步区分
                # 如果金重值较小(<100g)且业绩金额也较小(可能是银料按金价列写的)
                paid = abs(float(row.get("业绩金额", 0) or 0))
                if paid > 0 and gold_w > 0:
                    unit_price = paid / gold_w
                    # 黄金价通常 >400元/g，银价通常 <30元/g
                    if unit_price < 50:
                        return "silver"
                return "gold"

            # 默认按黄金处理（回收中黄金更常见）
            return "gold"

        for idx, row in recycle_df.iterrows():
            name = str(row.get("商品名称", ""))
            # 智能判断金属类型
            metal_type = _guess_metal_type(row)
            weight = _get_metal_weight(row, metal_type)

            if metal_type == "gold":
                price = self.gold_price
                gold_count += 1
                gold_g_total += weight
                paid = abs(float(row.get("业绩金额", 0) or 0))
                gold_paid_total += paid
                value = weight * price
                gold_value_total += value
            else:
                price = self.silver_price
                silver_count += 1
                silver_g_total += weight
                paid = abs(float(row.get("业绩金额", 0) or 0))
                silver_paid_total += paid
                value = weight * price
                silver_value_total += value

            pnl = value - paid

            # 日期格式化
            tx_date = row.get(date_col, "")
            if hasattr(tx_date, 'strftime'):
                tx_date = tx_date.strftime("%Y-%m-%d %H:%M")

            recycle_stats["recycle_items"].append({
                "date": str(tx_date),
                "name": name[:30],
                "metal": metal_type,
                "weight": round(weight, 3),
                "paid": round(paid, 2),
                "value": round(value, 2),
                "pnl": round(pnl, 2),
                "price": price,
            })

        # 汇总
        recycle_stats["recycle_total_paid"] = round(gold_paid_total + silver_paid_total, 2)
        recycle_stats["recycle_total_value"] = round(gold_value_total + silver_value_total, 2)
        recycle_stats["recycle_total_pnl"] = round(recycle_stats["recycle_total_value"] - recycle_stats["recycle_total_paid"], 2)
        recycle_stats["recycle_gold_g"] = round(gold_g_total, 3)
        recycle_stats["recycle_gold_paid"] = round(gold_paid_total, 2)
        recycle_stats["recycle_gold_value"] = round(gold_value_total, 2)
        recycle_stats["recycle_gold_count"] = gold_count
        recycle_stats["recycle_silver_g"] = round(silver_g_total, 3)
        recycle_stats["recycle_silver_paid"] = round(silver_paid_total, 2)
        recycle_stats["recycle_silver_value"] = round(silver_value_total, 2)
        recycle_stats["recycle_silver_count"] = silver_count

        logger.info(f"回收统计 | 金:{gold_count}笔/{gold_g_total:.3f}g | 银:{silver_count}笔/{silver_g_total:.3f}g | 盈亏:¥{recycle_stats['recycle_total_pnl']:+,.2f}")
        return recycle_stats

    def _generate_default_insights(self, stats: Dict[str, Any]) -> List[str]:
        """生成默认的业务洞察"""
        insights = []

        # 价格区间洞察
        if stats.get("range_stats"):
            top_range = max(stats["range_stats"], key=lambda x: x["total"])
            total_rev = stats.get("total_revenue", 1)
            pct = round(top_range["total"] / total_rev * 100, 1) if total_rev > 0 else 0
            insights.append(f"最畅销的价格区间是「{top_range['range']}」，贡献了{pct}%的营业额（共{top_range['count']}笔交易）")

        # 品类洞察
        if stats.get("cat_stats"):
            top_cat = stats["cat_stats"][0] if stats["cat_stats"] else None
            if top_cat:
                insights.append(f"核心品类为「{top_cat['name']}」，业绩达¥{top_cat['total']:,}，占总业绩比重最高")

        # 回收业务洞察
        if stats.get("recycle_total_pnl") is not None:
            pnl = stats["recycle_total_pnl"]
            if pnl > 0:
                insights.append(f"旧料回收业务整体盈利 ¥{pnl:+,}，当前金属价格有利")
            elif pnl < 0:
                insights.append(f"旧料回收业务账面亏损 ¥{pnl:,}，属正常持仓浮亏（随金价波动）")
            else:
                insights.append("旧料回收业务暂无盈亏")

        # 商品洞察
        if stats.get("prod_stats"):
            top_prod = stats["prod_stats"][0] if stats["prod_stats"] else None
            if top_prod:
                insights.append(f"热销商品TOP1：「{top_prod['name']}」，销量{top_prod['count']}件，业绩¥{top_prod['total']:,}")

        if not insights:
            insights.append("数据已加载，可进一步分析各维度趋势")

        return insights

    def _get_empty_result(self) -> Dict[str, Any]:
        """返回空的结果模板"""
        return {
            "total_revenue": 0,
            "total_profit": 0,
            "sales_count": 0,
            "time_range": None,
            "range_stats": [],
            "range_details": {},
            "cat_stats": [],
            "prod_stats": [],
            "recycle_items": [],
            "recycle_total_paid": 0,
            "recycle_total_value": 0,
            "recycle_total_pnl": 0,
            "recycle_gold_g": 0,
            "recycle_gold_paid": 0,
            "recycle_gold_value": 0,
            "recycle_gold_count": 0,
            "recycle_silver_g": 0,
            "recycle_silver_paid": 0,
            "recycle_silver_value": 0,
            "recycle_silver_count": 0,
            "gold_price": self.gold_price,
            "silver_price": self.silver_price,
            "insights": ["AI分析服务暂时不可用，请稍后重试"],
        }

    def _generate_report(self, data: Dict[str, Any]) -> str:
        """
        生成 HTML 报告

        从 report_template.html 读取模板并填充数据

        Args:
            data: 分析结果数据

        Returns:
            HTML 报告字符串
        """
        try:
            # 读取模板
            template_path = __file__.replace("ai_analyzer.py", "report_template.html")
            with open(template_path, "r", encoding="utf-8") as f:
                template = f.read()

            # 替换占位符
            html = template.replace("{{DATA_JSON}}", json.dumps(data, ensure_ascii=False, indent=2))

            return html

        except Exception as e:
            logger.error(f"生成HTML报告失败: {e}")
            return f"<html><body><h1>报告生成失败</h1><p>错误: {e}</p></body></html>"


def analyze_sales_data_with_ai(
    file_path: str,
    api_key: str,
    gold_price: float = 930.0,
    silver_price: float = 17.0,
    model: str = "MiniMax-M3",
) -> Dict[str, Any]:
    """
    便捷函数：使用 AI 分析销售数据

    Args:
        file_path: Excel 文件路径
        api_key: MiniMax API 密钥
        gold_price: 金价（元/克）
        silver_price: 银价（元/克）
        model: 模型名称

    Returns:
        完整的分析结果字典

    Example:
        >>> result = analyze_sales_data_with_ai(
        ...     "sales.xls",
        ...     api_key="sk-xxx",
        ...     gold_price=930,
        ...     silver_price=17
        ... )
        >>> print(result['total_revenue'])
    """
    analyzer = AISalesAnalyzer(
        api_key=api_key,
        model=model,
        gold_price=gold_price,
        silver_price=silver_price,
    )
    return analyzer.analyze_file(file_path)


if __name__ == "__main__":
    # 测试入口
    import sys

    print("=" * 60)
    print("AI 销售数据分析器测试")
    print("=" * 60)

    # 检查参数
    if len(sys.argv) < 2:
        print("\n用法: python ai_analyzer.py <excel文件路径>")
        print("示例: python ai_analyzer.py sales_data.xls")
        exit(1)

    file_path = sys.argv[1]

    # 检查 API Key
    api_key = os.getenv("MINIMAX_API_KEY")
    if not api_key:
        print("\n❌ 请设置环境变量 MINIMAX_API_KEY")
        exit(1)

    print(f"\n📂 分析文件: {file_path}")
    print(f"🔑 API Key: {api_key[:10]}...{api_key[-6:]}\n")

    # 执行分析
    try:
        result = analyze_sales_data_with_ai(
            file_path=file_path,
            api_key=api_key,
        )

        print("\n✅ 分析完成！")
        print(f"\n📊 关键指标:")
        print(f"  - 营业额: ¥{result.get('total_revenue', 0):,.2f}")
        print(f"  - 利润: ¥{result.get('total_profit', 0):,.2f}")
        print(f"  - 销售笔数: {result.get('sales_count', 0)}")
        print(f"\n📝 业务洞察:")
        for insight in result.get("insights", []):
            print(f"  • {insight}")

        print(f"\n💾 报告ID: {result.get('id')}")
        print(f"📄 报告长度: {len(result.get('report_html', ''))} 字符")

    except Exception as e:
        print(f"\n❌ 分析失败: {e}")
        import traceback

        traceback.print_exc()
