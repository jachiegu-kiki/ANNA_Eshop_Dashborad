"""
电商店铺数据看板 — 数据清洗与分类计算管道 v2
=============================================
v2 更新：
  - 每行 = 1件商品（忽略"商品数量"字段）
  - 销售价按订单维度分摊：SKU标价占比 × 实付金额 → TWD → 按当日汇率折算CNY
  - 汇率通过API自动获取，按采购日期匹配当日汇率，带本地缓存
  - 毛利率使用加权平均（Σ利润 / ΣGMV）

运行方式：
  python ecom_pipeline.py --input ./data/202503.xlsx --output ./output/dashboard_data.json
  python ecom_pipeline.py --input ./data/ --output ./output/dashboard_data.json   # 读取目录下所有xlsx/csv

依赖：pip install pandas openpyxl xlrd requests

架构位置：N8N定时触发 → 本脚本 → JSON → HTML看板读取
"""

import sys
import subprocess as _sp
import pandas as pd
import json
import argparse
import os
import glob
import logging
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional, List, Dict

# ── 自动安装缺失依赖（兼容 venv 环境） ──
def _ensure(pkg, import_as=None):
    try:
        __import__(import_as or pkg)
    except ImportError:
        _sp.check_call([sys.executable, "-m", "pip", "install", pkg, "--quiet"])

_ensure("dataclasses")   # Python 3.6 backport（3.7+ 内置，安装无害）
_ensure("openpyxl")
_ensure("xlrd")
_ensure("requests")

from dataclasses import dataclass

try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    import urllib.request
    import urllib.error
    HAS_REQUESTS = False


# ============================================================
# 1. 配置区
# ============================================================

@dataclass
class StoreConfig:
    """单店铺配置，由老板指定目标毛利率"""
    store_name: str
    target_margin: float  # 店铺目标毛利率（%），如 35 表示 35%


# 默认店铺配置（后续可从数据库/API读取）
DEFAULT_STORE_CONFIGS = [
    StoreConfig("23号店【VFU瑜珈服】", 35.0),
    StoreConfig("19号店-【健身女孩瑜珈服】", 35.0),
]

# 各定位产品的毛利率警告阈值（低于此值触发告警）
CATEGORY_MARGIN_THRESHOLDS = {
    "利润款": 80.0,
    "基础款": 60.0,
    "流量款": 30.0,
    "调整款": 40.0,
}

# 流量款销量占比阈值
TRAFFIC_QTY_SHARE_THRESHOLD = 5.0  # %

# 全局毛利率告警线（KPI卡片用）
GLOBAL_MARGIN_WARNING = 35.0

# 默认数据年份（日期格式"3月2日"缺少年份，此处补充）
DEFAULT_YEAR = 2026

# 默认兜底汇率（所有API都失败时使用）
# 注意方向：1 TWD = 0.22 CNY（不是4.55，那是反向的 1 CNY = 4.55 TWD）
FALLBACK_EXCHANGE_RATE = 0.22  # 1 TWD → ? CNY

# 汇率合理性区间（防止API返回异常值）
FX_RATE_MIN = 0.15   # 1 TWD 最低 0.15 CNY
FX_RATE_MAX = 0.35   # 1 TWD 最高 0.35 CNY

# 日志
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)


# ============================================================
# 2. 汇率模块 — TWD → CNY 每日汇率
# ============================================================

class ExchangeRateService:
    """
    获取 TWD→CNY 历史汇率，带本地缓存。
    优先从缓存读取，未命中时调用API，结果写入缓存文件。

    API来源（按优先级）：
      1. fawazahmed0 currency-api（免费，无需API Key）
      2. 兜底使用 FALLBACK_EXCHANGE_RATE
    """

    def __init__(self, cache_dir: Optional[str] = None):
        if cache_dir is None:
            cache_dir = str(Path(os.path.abspath(__file__)).parent / "cache")
        self.cache_file = Path(cache_dir) / "exchange_rates.json"
        self.cache = self._load_cache()
        self._dirty = False  # 是否有新数据需要写入

    def _load_cache(self) -> Dict[str, float]:
        """从本地文件加载缓存"""
        if self.cache_file.exists():
            try:
                with open(self.cache_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                logger.info(f"汇率缓存已加载: {len(data)} 条记录")
                return data
            except (json.JSONDecodeError, IOError):
                logger.warning("汇率缓存文件损坏，重新创建")
        return {}

    def _save_cache(self):
        """将缓存写入本地文件"""
        if not self._dirty:
            return
        self.cache_file.parent.mkdir(parents=True, exist_ok=True)
        with open(self.cache_file, "w", encoding="utf-8") as f:
            json.dump(self.cache, f, indent=2)
        logger.info(f"汇率缓存已保存: {len(self.cache)} 条记录")
        self._dirty = False

    def _validate_rate(self, rate: float, source: str = "") -> Optional[float]:
        """
        校验汇率是否在合理区间内。
        如果拿到的是反向汇率（如4.55），自动取倒数修正。
        """
        if rate is None or rate <= 0:
            return None

        # 如果在合理区间，直接返回
        if FX_RATE_MIN <= rate <= FX_RATE_MAX:
            return round(rate, 6)

        # 如果值 > 1，很可能是反向汇率（1 CNY = X TWD），取倒数
        if rate > 1:
            inverted = 1.0 / rate
            if FX_RATE_MIN <= inverted <= FX_RATE_MAX:
                logger.warning(f"  汇率方向修正{source}: {rate} → 1/{rate} = {inverted:.6f}")
                return round(inverted, 6)

        logger.warning(f"  汇率异常{source}: {rate}，不在合理区间 [{FX_RATE_MIN}, {FX_RATE_MAX}]，丢弃")
        return None

    def _fetch_from_api(self, date_str: str) -> Optional[float]:
        """
        从 API 获取 TWD→CNY 汇率。
        date_str 格式: "2025-03-02"
        返回: 1 TWD = ? CNY，或 None（失败）
        """
        # API 1: fawazahmed0 currency-api
        urls = [
            f"https://{date_str}.currency-api.pages.dev/v1/currencies/twd.json",
            f"https://cdn.jsdelivr.net/npm/@fawazahmed0/currency-api@{date_str}/v1/currencies/twd.json",
        ]

        for url in urls:
            try:
                if HAS_REQUESTS:
                    resp = requests.get(url, timeout=5)
                    if resp.status_code == 200:
                        data = resp.json()
                        rate = data.get("twd", {}).get("cny")
                        if rate:
                            validated = self._validate_rate(float(rate), f"(API: {url})")
                            if validated:
                                return validated
                else:
                    req = urllib.request.Request(url)
                    with urllib.request.urlopen(req, timeout=5) as resp:
                        data = json.loads(resp.read().decode("utf-8"))
                        rate = data.get("twd", {}).get("cny")
                        if rate:
                            validated = self._validate_rate(float(rate), f"(API: {url})")
                            if validated:
                                return validated
            except Exception as e:
                logger.debug(f"API请求失败 {url}: {e}")
                continue

        return None

    def get_rate(self, date_str: str) -> float:
        """
        获取指定日期的 TWD→CNY 汇率（1 TWD = ? CNY）。
        date_str 格式: "2025-03-02"
        查找策略：先查缓存 → 调API → 尝试前几天 → 兜底值
        """
        # 1. 查缓存（同时校验缓存值是否合理）
        if date_str in self.cache:
            cached = self.cache[date_str]
            if FX_RATE_MIN <= cached <= FX_RATE_MAX:
                return cached
            else:
                logger.warning(f"  缓存汇率异常 {date_str}={cached}，重新获取")
                del self.cache[date_str]

        # 2. 调API（当天）
        rate = self._fetch_from_api(date_str)
        if rate:
            self.cache[date_str] = rate
            self._dirty = True
            logger.info(f"  汇率 {date_str}: 1 TWD = {rate} CNY (API)")
            return rate

        # 3. 当天失败（周末/节假日），往前找最多7天
        base_date = datetime.strptime(date_str, "%Y-%m-%d")
        for delta in range(1, 8):
            prev_date = (base_date - timedelta(days=delta)).strftime("%Y-%m-%d")
            # 先看缓存
            if prev_date in self.cache and FX_RATE_MIN <= self.cache[prev_date] <= FX_RATE_MAX:
                self.cache[date_str] = self.cache[prev_date]
                self._dirty = True
                logger.info(f"  汇率 {date_str}: 使用 {prev_date} 的缓存值 {self.cache[prev_date]}")
                return self.cache[prev_date]
            # 再调API
            rate = self._fetch_from_api(prev_date)
            if rate:
                self.cache[prev_date] = rate
                self.cache[date_str] = rate
                self._dirty = True
                logger.info(f"  汇率 {date_str}: 使用 {prev_date} 的值 {rate} (API)")
                return rate

        # 4. 全部失败，使用兜底值
        logger.warning(f"  汇率 {date_str}: 所有API失败，使用兜底值 {FALLBACK_EXCHANGE_RATE}")
        self.cache[date_str] = FALLBACK_EXCHANGE_RATE
        self._dirty = True
        return FALLBACK_EXCHANGE_RATE

    def _is_network_available(self) -> bool:
        """
        快速检测外网是否可达（单次最多等待5秒）。
        用于 batch_fetch 前的快速判断，网络不通则跳过所有 API 调用。
        """
        test_url = "https://cdn.jsdelivr.net/npm/@fawazahmed0/currency-api@latest/v1/currencies.json"
        try:
            if HAS_REQUESTS:
                resp = requests.head(test_url, timeout=5)
                return resp.status_code < 500
            else:
                urllib.request.urlopen(test_url, timeout=5)
                return True
        except Exception:
            return False

    def batch_fetch(self, date_list: List[str]):
        """
        批量预取汇率（减少逐条查询的等待）。
        date_list: ["2025-03-02", "2025-03-03", ...]
        """
        unique_dates = sorted(set(date_list))
        need_fetch = [d for d in unique_dates if d not in self.cache]

        if not need_fetch:
            logger.info(f"汇率: 全部 {len(unique_dates)} 个日期已有缓存")
            return

        logger.info(f"汇率: 需要获取 {len(need_fetch)} 个日期的汇率，检测网络...")

        # ── 网络预检：不通则立即用兜底值，避免每个日期等待几分钟 ──
        if not self._is_network_available():
            logger.warning(
                f"⚠️  外网不可达，{len(need_fetch)} 个日期全部使用兜底汇率 "
                f"{FALLBACK_EXCHANGE_RATE}（1 TWD = {FALLBACK_EXCHANGE_RATE} CNY）"
            )
            for d in need_fetch:
                self.cache[d] = FALLBACK_EXCHANGE_RATE
                self._dirty = True
            self._save_cache()
            return

        logger.info(f"网络正常，开始获取 {len(need_fetch)} 个日期的汇率...")

        for i, date_str in enumerate(need_fetch):
            self.get_rate(date_str)
            # 礼貌性延迟，避免被限流
            if i < len(need_fetch) - 1:
                time.sleep(0.1)

        self._save_cache()
        logger.info("汇率批量获取完成")

    def close(self):
        """确保缓存写入磁盘"""
        self._save_cache()


# ============================================================
# 3. 数据读取与清洗
# ============================================================

def _detect_header_row(df_raw: pd.DataFrame, max_scan: int = 10) -> int:
    """
    自动检测真正的表头行。
    策略：遍历前 max_scan 行，找匹配 COLUMN_MAP 必要字段最多的那行。
    没有任何匹配时默认 row=0（避免 skiprows=1 的错误假设）。
    """
    required_names = set(COLUMN_MAP.keys())
    best_row, best_score = 0, 0
    for i in range(min(max_scan, len(df_raw))):
        row_vals = {str(v).strip() for v in df_raw.iloc[i] if pd.notna(v)}
        score = len(row_vals & required_names)
        if score > best_score:
            best_score, best_row = score, i
    logger.info(f"  自动检测表头行: row={best_row}（匹配 {best_score} 个字段）")
    return best_row


def read_file(filepath: str) -> pd.DataFrame:
    """
    读取单个数据文件，支持 xlsx / xls / csv。
    自动检测表头行，不依赖 skiprows=1 的固定假设。
    """
    ext = Path(filepath).suffix.lower()
    name = Path(filepath).name

    if ext in (".xlsx", ".xls"):
        engine = "openpyxl" if ext == ".xlsx" else "xlrd"
        # 先无表头读取前10行，自动检测表头位置
        df_raw = pd.read_excel(filepath, engine=engine, header=None, nrows=10)
        header_row = _detect_header_row(df_raw)
        df = pd.read_excel(filepath, engine=engine, header=header_row)
        logger.info(f"成功读取 {name}（{ext[1:]}），表头行={header_row + 1}，数据行数={len(df)}")
        return df

    elif ext == ".csv":
        for enc in ["utf-8-sig", "utf-8", "gbk", "gb2312", "gb18030"]:
            try:
                # CSV 同样自动检测表头行
                df_raw = pd.read_csv(filepath, encoding=enc, header=None, nrows=10)
                header_row = _detect_header_row(df_raw)
                df = pd.read_csv(filepath, encoding=enc, header=header_row)
                logger.info(f"成功读取 {name}（CSV, {enc}），表头行={header_row + 1}，数据行数={len(df)}")
                return df
            except (UnicodeDecodeError, UnicodeError):
                continue
        raise ValueError(f"无法解码CSV文件: {filepath}")

    else:
        raise ValueError(f"不支持的文件格式: {ext}，仅支持 .xlsx / .xls / .csv")


def load_data(input_path: str) -> pd.DataFrame:
    """
    加载数据，支持单文件或目录批量加载。
    目录模式下自动扫描所有 xlsx / xls / csv 文件。
    """
    path = Path(input_path)

    if path.is_file():
        frames = [read_file(str(path))]

    elif path.is_dir():
        # 扫描目录下所有支持的文件格式
        patterns = ["*.xlsx", "*.xls", "*.csv"]
        files = []
        for pat in patterns:
            files.extend(glob.glob(str(path / pat)))
        files = sorted(set(files))  # 去重+排序

        if not files:
            raise FileNotFoundError(
                f"目录 {input_path} 下未找到 xlsx/xls/csv 文件"
            )

        logger.info(f"扫描目录 {input_path}，发现 {len(files)} 个文件：")
        for f in files:
            logger.info(f"  → {Path(f).name}")

        frames = [read_file(f) for f in files]

    else:
        raise FileNotFoundError(f"路径不存在: {input_path}")

    df = pd.concat(frames, ignore_index=True)
    logger.info(f"合并后总行数: {len(df)}")
    return df


# 必要字段映射（原始列名 → 内部列名）
COLUMN_MAP = {
    "采购日期":     "date",           # 格式: "3月2日"
    "订单号":       "order_id",       # 订单维度分摊用
    "店铺名":       "store",
    "商品ID":       "style_id",       # 款式维度
    "货品编码":     "sku_code",       # 规格编码（保留备用）
    "2--商品价格":  "item_price_twd", # SKU标价（台币）
    "实付金额":     "paid_amount_twd",# 订单实付总额（台币）
    "实际采购价":   "cost_price",     # 采购成本（人民币）
}

def clean_data(df: pd.DataFrame, fx_service: ExchangeRateService) -> pd.DataFrame:
    """
    数据清洗（v2）：
    1. 校验字段，重命名
    2. 基础过滤（空店铺、标价≤0）
    3. 每行 = 1件商品
    4. 按订单分摊销售价（用全部行做分母，保证分摊准确）
    5. 过滤成本价为0的行（在分摊之后，避免混合订单分母偏差）
    6. 解析日期，获取汇率，折算CNY
    7. 重新计算毛利润
    """
    # ── 校验字段 ──
    missing = [c for c in COLUMN_MAP if c not in df.columns]
    if missing:
        raise KeyError(f"缺少必要字段: {missing}。实际字段: {list(df.columns)}")

    df = df.rename(columns=COLUMN_MAP).copy()

    # ── 过滤空店铺 ──
    before = len(df)
    df = df[df["store"].notna() & (df["store"].str.strip() != "")]
    logger.info(f"过滤空店铺: {before} → {len(df)} 行")

    # ── 数值字段修正 ──
    for col in ["item_price_twd", "paid_amount_twd", "cost_price"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # ── 过滤标价/实付异常 ──
    before = len(df)
    df = df[(df["item_price_twd"] > 0) & (df["paid_amount_twd"] > 0)].copy()
    logger.info(f"过滤标价/实付≤0: {before} → {len(df)} 行")

    # ── 每行 = 1件商品 ──
    df["qty"] = 1

    # ── 按订单分摊销售价（TWD）——必须在过滤成本为0之前 ──
    # 先用全部行计算订单标价合计，保证混合订单的分母正确
    order_total_price = df.groupby("order_id")["item_price_twd"].transform("sum")
    df["sale_price_twd"] = (df["item_price_twd"] / order_total_price * df["paid_amount_twd"]).round(2)

    # ── 过滤实际成本价为0的行 ──
    # 放在分摊计算之后，这样成本为0的行已参与分母计算
    # 有效行只拿属于自己的那份，不会多分
    before = len(df)
    df = df[df["cost_price"] > 0].copy()
    excluded = before - len(df)
    logger.info(f"过滤成本价为0: {before} → {len(df)} 行（剔除 {excluded} 行）")

    # ── 解析日期（兼容 datetime / 字符串 两种格式） ──
    if pd.api.types.is_datetime64_any_dtype(df["date"]):
        # Excel 原生日期格式：Timestamp 对象
        df["date"] = pd.to_datetime(df["date"])
        df["year"]  = df["date"].dt.year
        df["month"] = df["date"].dt.month
        df["day"]   = df["date"].dt.day
        logger.info("日期列为 datetime 类型，直接提取年月日")
    else:
        # 文本格式："3月2日" 等
        df["date"] = df["date"].astype(str)
        df["month"] = df["date"].str.extract(r"(\d+)月")[0].astype(int)
        df["day"]   = df["date"].str.extract(r"月(\d+)日")[0].astype(int)
        df["year"]  = DEFAULT_YEAR
        logger.info("日期列为字符串类型，通过正则提取月日")

    df["full_date"] = df.apply(
        lambda r: f"{int(r['year'])}-{int(r['month']):02d}-{int(r['day']):02d}", axis=1
    )
    df["year_month"] = df.apply(
        lambda r: f"{int(r['year'])}-{int(r['month']):02d}", axis=1
    )

    # ── style_id 转字符串 ──
    df["style_id"] = df["style_id"].apply(
        lambda x: str(int(x)) if pd.notna(x) and x == x else ""
    )

    # ── 批量获取汇率 ──
    unique_dates = df["full_date"].unique().tolist()
    logger.info(f"涉及 {len(unique_dates)} 个不同日期，开始获取汇率...")
    fx_service.batch_fetch(unique_dates)

    df["fx_rate"] = df["full_date"].apply(lambda d: fx_service.get_rate(d))

    # ── 汇率诊断日志 ──
    rate_summary = df.groupby("full_date")["fx_rate"].first().sort_index()
    logger.info("各日期汇率 (1 TWD = ? CNY):")
    for date, rate in rate_summary.items():
        flag = "✓" if FX_RATE_MIN <= rate <= FX_RATE_MAX else "⚠️ 异常"
        logger.info(f"  {date}: {rate:.6f} {flag}")

    # ── TWD → CNY ──
    df["sale_price"] = (df["sale_price_twd"] * df["fx_rate"]).round(2)

    # ── 重新计算毛利润 ──
    df["profit"] = (df["sale_price"] - df["cost_price"]).round(2)

    # ── 数据截止日 ──
    max_month = df["month"].max()
    month_data = df[df["month"] == max_month]
    max_day = month_data["day"].max()
    df.attrs["data_cutoff"] = f"{DEFAULT_YEAR}年{max_month}月{max_day}日"
    logger.info(f"数据截止日: {df.attrs['data_cutoff']}")

    # ── 日志 ──
    logger.info(f"清洗完成: {len(df)} 行, "
                f"总GMV(CNY)=¥{df['sale_price'].sum():,.2f}, "
                f"总利润=¥{df['profit'].sum():,.2f}")

    return df


# ============================================================
# 4. 款式聚合与分类
# ============================================================

def aggregate_styles(df: pd.DataFrame) -> pd.DataFrame:
    """
    按 (店铺, 款式ID, 年月) 聚合到款式粒度。
    毛利率 = 加权平均 = Σ利润 / ΣGMV × 100
    """
    agg = df.groupby(["store", "style_id", "year_month"]).agg(
        gmv=("sale_price", "sum"),
        qty=("qty", "sum"),
        profit=("profit", "sum"),
        cost=("cost_price", "sum"),
    ).reset_index()

    # 加权平均毛利率
    agg["margin_rate"] = agg.apply(
        lambda r: round(r["profit"] / r["gmv"] * 100, 2) if r["gmv"] > 0 else 0,
        axis=1,
    )

    # 款式销量占比（相对于同店铺同月总销量）
    store_month_qty = agg.groupby(["store", "year_month"])["qty"].transform("sum")
    agg["qty_share"] = round(agg["qty"] / store_month_qty * 100, 2)

    logger.info(f"聚合后款式数: {len(agg)}")
    return agg


def get_store_target_margin(store_name: str, configs: List[StoreConfig]) -> float:
    """获取店铺目标毛利率，找不到则使用全局默认值"""
    for cfg in configs:
        if cfg.store_name == store_name:
            return cfg.target_margin
    logger.warning(f"未找到店铺 [{store_name}] 的配置，使用默认目标毛利率 {GLOBAL_MARGIN_WARNING}%")
    return GLOBAL_MARGIN_WARNING


def classify_styles(
    styles_df: pd.DataFrame,
    store_configs: List[StoreConfig],
) -> pd.DataFrame:
    """
    按业务规则对款式分类，基准线 = 店铺目标毛利率（由老板指定）。

    分类逻辑（优先级从高到低）：
      1. 利润款：款式毛利率 >= 目标毛利率 + 10%
      2. 流量款：款式毛利率 <= 目标毛利率 - 10%  且  销量占比 >= 5%
      3. 基础款：目标毛利率 - 10% <= 款式毛利率 <= 目标毛利率 + 10%
      4. 调整款：不满足以上任何条件
    """
    df = styles_df.copy()

    df["target_margin"] = df["store"].apply(
        lambda s: get_store_target_margin(s, store_configs)
    )

    upper = df["target_margin"] + 10
    lower = df["target_margin"] - 10

    # 按优先级：调整款 → 基础款 → 流量款 → 利润款（后赋值覆盖前）
    df["category"] = "调整款"

    mask_basic = (df["margin_rate"] >= lower) & (df["margin_rate"] <= upper)
    df.loc[mask_basic, "category"] = "基础款"

    mask_traffic = (df["margin_rate"] <= lower) & (
        df["qty_share"] >= TRAFFIC_QTY_SHARE_THRESHOLD
    )
    df.loc[mask_traffic, "category"] = "流量款"

    mask_profit = df["margin_rate"] >= upper
    df.loc[mask_profit, "category"] = "利润款"

    for cat in ["利润款", "流量款", "基础款", "调整款"]:
        count = (df["category"] == cat).sum()
        logger.info(f"  {cat}: {count} 款")

    return df


# ============================================================
# 5. 汇总计算（供看板消费）
# ============================================================

def compute_dashboard_data(
    df: pd.DataFrame,
    styles_df: pd.DataFrame,
    store_configs: List[StoreConfig],
) -> dict:
    """计算看板所需的全部汇总数据，输出完整JSON结构。"""
    result = {
        "meta": {
            "data_cutoff": df.attrs.get("data_cutoff", "未知"),
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "stores": sorted(df["store"].unique().tolist()),
            "months": sorted(df["year_month"].unique().tolist()),
        },
        "config": {
            "store_configs": {c.store_name: c.target_margin for c in store_configs},
            "category_thresholds": CATEGORY_MARGIN_THRESHOLDS,
            "global_margin_warning": GLOBAL_MARGIN_WARNING,
            "traffic_qty_share_threshold": TRAFFIC_QTY_SHARE_THRESHOLD,
        },
        "raw_styles": [],
        "summary": {},
    }

    # 款式明细
    for _, r in styles_df.iterrows():
        result["raw_styles"].append({
            "store": r["store"],
            "style_id": r["style_id"],
            "year_month": r["year_month"],
            "gmv": round(float(r["gmv"]), 2),
            "qty": int(r["qty"]),
            "profit": round(float(r["profit"]), 2),
            "margin_rate": round(float(r["margin_rate"]), 2),
            "qty_share": round(float(r["qty_share"]), 2),
            "category": r["category"],
            "target_margin": round(float(r["target_margin"]), 2),
        })

    # 通用汇总函数
    def summarize(subset: pd.DataFrame) -> dict:
        total_gmv = round(float(subset["gmv"].sum()), 2)
        total_qty = int(subset["qty"].sum())
        total_profit = round(float(subset["profit"].sum()), 2)
        total_margin = round(total_profit / total_gmv * 100, 2) if total_gmv > 0 else 0
        style_count = int(subset["style_id"].nunique())

        kpi = {
            "gmv": total_gmv,
            "qty": total_qty,
            "style_count": style_count,
            "margin_rate": total_margin,
            "margin_warning": bool(total_margin < GLOBAL_MARGIN_WARNING),
        }

        categories = {}
        for cat in ["利润款", "流量款", "基础款", "调整款"]:
            cat_data = subset[subset["category"] == cat]
            cat_gmv = round(float(cat_data["gmv"].sum()), 2)
            cat_qty = int(cat_data["qty"].sum())
            cat_profit = round(float(cat_data["profit"].sum()), 2)
            cat_margin = round(cat_profit / cat_gmv * 100, 2) if cat_gmv > 0 else 0
            cat_styles = int(cat_data["style_id"].nunique())
            threshold = CATEGORY_MARGIN_THRESHOLDS.get(cat, 0)

            categories[cat] = {
                "style_count": cat_styles,
                "gmv": cat_gmv,
                "gmv_share": round(cat_gmv / total_gmv * 100, 2) if total_gmv > 0 else 0,
                "qty": cat_qty,
                "qty_share": round(cat_qty / total_qty * 100, 2) if total_qty > 0 else 0,
                "margin_rate": cat_margin,
                "margin_threshold": threshold,
                "margin_warning": bool(cat_styles > 0 and cat_margin < threshold),
            }

        return {"kpi": kpi, "categories": categories}

    result["summary"]["全部"] = summarize(styles_df)

    result["summary"]["按店铺"] = {}
    for store in styles_df["store"].unique():
        result["summary"]["按店铺"][store] = summarize(
            styles_df[styles_df["store"] == store]
        )

    result["summary"]["按月份"] = {}
    for ym in sorted(styles_df["year_month"].unique()):
        result["summary"]["按月份"][ym] = summarize(
            styles_df[styles_df["year_month"] == ym]
        )

    return result


# ============================================================
# 6. 输出
# ============================================================

class NumpyEncoder(json.JSONEncoder):
    """处理numpy类型的JSON编码器"""
    def default(self, obj):
        import numpy as np
        if isinstance(obj, (np.bool_,)):
            return bool(obj)
        if isinstance(obj, (np.integer,)):
            return int(obj)
        if isinstance(obj, (np.floating,)):
            return float(obj)
        if isinstance(obj, (np.ndarray,)):
            return obj.tolist()
        return super().default(obj)


def save_json(data: dict, output_path: str):
    """输出JSON文件"""
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, cls=NumpyEncoder)
    size_kb = os.path.getsize(output_path) / 1024
    logger.info(f"JSON已保存: {output_path} ({size_kb:.1f} KB)")


# ============================================================
# 7. 主流程
# ============================================================

def run_pipeline(
    input_path: str,
    output_path: str,
    store_configs: Optional[List[StoreConfig]] = None,
):
    """完整管道：读取 → 清洗(含汇率) → 聚合 → 分类 → 汇总 → 输出JSON"""
    if store_configs is None:
        store_configs = DEFAULT_STORE_CONFIGS

    logger.info("=" * 60)
    logger.info("电商数据管道 v2 启动")
    logger.info("=" * 60)

    # Step 1: 读取
    raw_df = load_data(input_path)

    # Step 2: 清洗（含汇率获取）
    fx_service = ExchangeRateService()
    try:
        clean_df = clean_data(raw_df, fx_service)
    finally:
        fx_service.close()  # 确保汇率缓存写入磁盘

    # Step 3: 聚合
    styles_df = aggregate_styles(clean_df)

    # Step 4: 分类
    styles_df = classify_styles(styles_df, store_configs)

    # Step 5: 汇总
    dashboard_data = compute_dashboard_data(clean_df, styles_df, store_configs)

    # Step 6: 输出
    save_json(dashboard_data, output_path)

    # 摘要
    s = dashboard_data["summary"]["全部"]
    logger.info("-" * 40)
    logger.info("管道运行完成，数据摘要：")
    logger.info(f"  数据截止日: {dashboard_data['meta']['data_cutoff']}")
    logger.info(f"  店铺数: {len(dashboard_data['meta']['stores'])}")
    logger.info(f"  总GMV(CNY): ¥{s['kpi']['gmv']:,.2f}")
    logger.info(f"  总销量: {s['kpi']['qty']:,} 件")
    logger.info(f"  款式数: {s['kpi']['style_count']}")
    logger.info(f"  综合毛利率: {s['kpi']['margin_rate']}%")
    for cat in ["利润款", "流量款", "基础款", "调整款"]:
        c = s["categories"][cat]
        logger.info(
            f"  {cat}: {c['style_count']}款, "
            f"GMV=¥{c['gmv']:,.0f}({c['gmv_share']}%), "
            f"毛利率={c['margin_rate']}% "
            f"{'⚠️ 低于标准' if c['margin_warning'] else '✓'}"
        )
    logger.info("=" * 60)

    return dashboard_data


# ============================================================
# 8. CLI入口
# ============================================================

# ── 基础路径：自动定位到脚本所在目录 ──
SCRIPT_DIR = Path(os.path.abspath(__file__)).parent

# ── IDE 直接运行时的默认配置 ──
# 默认读取 data 目录下所有 xlsx/xls/csv 文件
DEFAULT_INPUT_PATH    = SCRIPT_DIR / "uploads_raw"
DEFAULT_OUTPUT_PATH   = SCRIPT_DIR / "output" / "dashboard_data.json"
DEFAULT_TARGET_MARGIN = 35.0


def main():
    parser = argparse.ArgumentParser(description="电商店铺数据看板 — 数据管道 v2")
    parser.add_argument("--input",  "-i", default=None, help="输入CSV路径或目录")
    parser.add_argument("--output", "-o", default=None, help="输出JSON路径")
    parser.add_argument("--target-margin", type=float, default=None, help="统一目标毛利率(%)")
    args = parser.parse_args()

    input_path    = Path(args.input)  if args.input  else DEFAULT_INPUT_PATH
    output_path   = Path(args.output) if args.output else DEFAULT_OUTPUT_PATH
    target_margin = args.target_margin if args.target_margin is not None else DEFAULT_TARGET_MARGIN

    logger.info(f"脚本目录: {SCRIPT_DIR}")
    logger.info(f"输入路径: {input_path}")
    logger.info(f"输出路径: {output_path}")

    if not input_path.exists():
        logger.error(
            f"输入路径不存在: {input_path}\n"
            f"  请确认目录结构：\n"
            f"  {SCRIPT_DIR}\n"
            f"    ├── ecom_pipeline.py\n"
            f"    └── data/\n"
            f"        └── 你的CSV文件.csv"
        )
        return

    configs = DEFAULT_STORE_CONFIGS
    if target_margin is not None:
        configs = [StoreConfig(c.store_name, target_margin) for c in configs]
        logger.info(f"目标毛利率: {target_margin}%")

    run_pipeline(str(input_path), str(output_path), configs)


if __name__ == "__main__":
    main()
