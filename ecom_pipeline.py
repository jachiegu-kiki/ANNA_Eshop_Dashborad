"""
电商店铺数据看板 — 数据清洗与分类计算管道 v3
=============================================
v3 更新（基于v2）：
- 新增: 补货备货成本回填模块，解决台湾发货/福州仓库出/厂家出货成本为0的问题
- 新增: 统一补货备货明细表生成（cost_collector）
- 新增: 加权平均法回填采购成本（cost_filler）
- 新增: 待人工处理清单输出（未匹配SKU + 补货表漏填采购价）
- 改进: 支持A/B组采购表列名差异自动适配
- 保留: v2 全部逻辑不变（分摊、汇率、分类、看板JSON）

数据流:
  补货备货申请表(A组+B组)
       ↓ cost_collector
  统一补货明细表
       ↓ cost_filler (加权平均)
  采购表(成本已回填) → clean_data → 聚合 → 分类 → JSON看板

运行方式:
  python ecom_pipeline_v3.py --input ./uploads_raw --output ./output/dashboard_data.json \
      --replen-a ./data/A组备货补货申请表.xlsx --replen-b ./data/B组补货申请表.xlsx

依赖：pip install pandas openpyxl xlrd requests
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
from typing import Optional, List, Dict, Tuple

# ── 自动安装缺失依赖 ──
def _ensure(pkg, import_as=None):
    try:
        __import__(import_as or pkg)
    except ImportError:
        _sp.check_call([sys.executable, "-m", "pip", "install", pkg, "--quiet", "--break-system-packages"])

_ensure("dataclasses")
_ensure("openpyxl")
_ensure("xlrd")
_ensure("requests")

from dataclasses import dataclass, field

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
    store_name: str
    target_margin: float

DEFAULT_STORE_CONFIGS = [
    StoreConfig("23号店【VFU瑜珈服】", 35.0),
    StoreConfig("19号店-【健身女孩瑜珈服】", 35.0),
    StoreConfig("39号店 --孕妇装", 35.0),
    StoreConfig("37号店-【十月结晶孕妇装】", 35.0),
    StoreConfig("18号店 -【童装】", 35.0),
]

CATEGORY_MARGIN_THRESHOLDS = {
    "利润款": 80.0, "基础款": 60.0, "流量款": 30.0, "调整款": 40.0,
}
TRAFFIC_QTY_SHARE_THRESHOLD = 5.0
GLOBAL_MARGIN_WARNING = 35.0
DEFAULT_YEAR = 2026
FALLBACK_EXCHANGE_RATE = 0.22
FX_RATE_MIN = 0.15
FX_RATE_MAX = 0.35

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

# ============================================================
# 2. 汇率模块 (与v2完全一致，此处省略重复注释)
# ============================================================

class ExchangeRateService:
    def __init__(self, cache_dir: Optional[str] = None):
        if cache_dir is None:
            cache_dir = str(Path(os.path.abspath(__file__)).parent / "cache")
        self.cache_file = Path(cache_dir) / "exchange_rates.json"
        self.cache = self._load_cache()
        self._dirty = False

    def _load_cache(self) -> Dict[str, float]:
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
        if not self._dirty:
            return
        self.cache_file.parent.mkdir(parents=True, exist_ok=True)
        with open(self.cache_file, "w", encoding="utf-8") as f:
            json.dump(self.cache, f, indent=2)
        self._dirty = False

    def _validate_rate(self, rate: float, source: str = "") -> Optional[float]:
        if rate is None or rate <= 0:
            return None
        if FX_RATE_MIN <= rate <= FX_RATE_MAX:
            return round(rate, 6)
        if rate > 1:
            inverted = 1.0 / rate
            if FX_RATE_MIN <= inverted <= FX_RATE_MAX:
                return round(inverted, 6)
        return None

    def _fetch_from_api(self, date_str: str) -> Optional[float]:
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
                            validated = self._validate_rate(float(rate))
                            if validated:
                                return validated
                else:
                    req = urllib.request.Request(url)
                    with urllib.request.urlopen(req, timeout=5) as resp:
                        data = json.loads(resp.read().decode("utf-8"))
                        rate = data.get("twd", {}).get("cny")
                        if rate:
                            validated = self._validate_rate(float(rate))
                            if validated:
                                return validated
            except Exception:
                continue
        return None

    def get_rate(self, date_str: str) -> float:
        if date_str in self.cache:
            cached = self.cache[date_str]
            if FX_RATE_MIN <= cached <= FX_RATE_MAX:
                return cached
            else:
                del self.cache[date_str]

        rate = self._fetch_from_api(date_str)
        if rate:
            self.cache[date_str] = rate
            self._dirty = True
            return rate

        base_date = datetime.strptime(date_str, "%Y-%m-%d")
        for delta in range(1, 8):
            prev_date = (base_date - timedelta(days=delta)).strftime("%Y-%m-%d")
            if prev_date in self.cache and FX_RATE_MIN <= self.cache[prev_date] <= FX_RATE_MAX:
                self.cache[date_str] = self.cache[prev_date]
                self._dirty = True
                return self.cache[prev_date]
            rate = self._fetch_from_api(prev_date)
            if rate:
                self.cache[prev_date] = rate
                self.cache[date_str] = rate
                self._dirty = True
                return rate

        self.cache[date_str] = FALLBACK_EXCHANGE_RATE
        self._dirty = True
        return FALLBACK_EXCHANGE_RATE

    def _is_network_available(self) -> bool:
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
        unique_dates = sorted(set(date_list))
        need_fetch = [d for d in unique_dates if d not in self.cache]
        if not need_fetch:
            return
        if not self._is_network_available():
            for d in need_fetch:
                self.cache[d] = FALLBACK_EXCHANGE_RATE
            self._dirty = True
            self._save_cache()
            return
        for i, date_str in enumerate(need_fetch):
            self.get_rate(date_str)
            if i < len(need_fetch) - 1:
                time.sleep(0.1)
        self._save_cache()

    def close(self):
        self._save_cache()


# ============================================================
# 3. 补货备货成本回填模块 (v3 新增)
# ============================================================

# --- 3a. 统一补货备货明细表 (原 cost_collector.py) ---

A_REPLEN_SHEET_CONFIG = {
    "2025年12月": {"sku_col": "货号（可后补）", "date_col": "日期"},
    "1月":        {"sku_col": "货品编码",       "date_col": "日期"},
    "2月":        {"sku_col": "货品编码",       "date_col": "日期"},
    "3月":        {"sku_col": "货品编码",       "date_col": "日期"},
    "4月":        {"sku_col": "货品编码",       "date_col": "日期"},
}
B_REPLEN_SHEET_CONFIG = {"sheet_name": "補貨表", "sku_col": "货品编码", "date_col": "采购日期"}
REPLEN_COMMON_COLS = {"price_col": "采购价", "qty_col": "实际采购", "amount_col": "合计金额"}


def _norm_sku(val) -> str:
    """标准化货品编码：转字符串、去空格、去 .0 后缀"""
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def _parse_excel_date(val):
    """解析日期：兼容 Excel 序列号(如46113)和正常日期"""
    if pd.isna(val):
        return pd.NaT
    if isinstance(val, (int, float)) and val > 40000:
        return pd.Timestamp(datetime(1899, 12, 30) + timedelta(days=int(val)))
    return pd.to_datetime(val, errors="coerce")


def _parse_replen_sheet(filepath, sheet_name, sku_col, date_col) -> pd.DataFrame:
    """解析单个补货 sheet，返回统一格式"""
    df = pd.read_excel(filepath, sheet_name=sheet_name, header=2)
    required = [sku_col, date_col, REPLEN_COMMON_COLS["price_col"],
                REPLEN_COMMON_COLS["qty_col"], REPLEN_COMMON_COLS["amount_col"]]
    missing = [c for c in required if c not in df.columns]
    if missing:
        logger.warning(f"补货表 [{sheet_name}] 缺少列 {missing}，跳过")
        return pd.DataFrame()
    return pd.DataFrame({
        "采购日期": pd.to_datetime(df[date_col], errors="coerce"),
        "货品编码": df[sku_col].apply(_norm_sku),
        "采购单价": pd.to_numeric(df[REPLEN_COMMON_COLS["price_col"]], errors="coerce"),
        "数量":     pd.to_numeric(df[REPLEN_COMMON_COLS["qty_col"]], errors="coerce"),
        "合计金额": pd.to_numeric(df[REPLEN_COMMON_COLS["amount_col"]], errors="coerce"),
        "来源":     sheet_name,
    })


def build_replenishment_master(
    a_filepath: Optional[str] = None,
    b_filepath: Optional[str] = None,
) -> pd.DataFrame:
    """
    构建统一补货备货明细表。
    筛选: 货品编码非空 AND 实际采购>0。
    采购价=0 标记为漏填（不参与加权计算，提醒人工补充）。
    """
    frames = []

    if a_filepath and Path(a_filepath).exists():
        for sheet_name, cfg in A_REPLEN_SHEET_CONFIG.items():
            try:
                sub = _parse_replen_sheet(a_filepath, sheet_name, cfg["sku_col"], cfg["date_col"])
                sub["来源"] = f"A组-{sheet_name}"
                if len(sub) > 0:
                    frames.append(sub)
                    logger.info(f"补货表 A组[{sheet_name}]: {len(sub)} 行")
            except Exception as e:
                logger.warning(f"补货表 A组[{sheet_name}] 失败: {e}")

    if b_filepath and Path(b_filepath).exists():
        cfg = B_REPLEN_SHEET_CONFIG
        try:
            sub = _parse_replen_sheet(b_filepath, cfg["sheet_name"], cfg["sku_col"], cfg["date_col"])
            sub["来源"] = "B组"
            if len(sub) > 0:
                frames.append(sub)
                logger.info(f"补货表 B组: {len(sub)} 行")
        except Exception as e:
            logger.warning(f"补货表 B组失败: {e}")

    if not frames:
        logger.warning("未提供补货备货表或无法读取，成本回填将跳过")
        return pd.DataFrame()

    master = pd.concat(frames, ignore_index=True)

    # 筛选: 货品编码非空
    master = master[master["货品编码"].notna() & (~master["货品编码"].isin(["", "nan", "None"]))].copy()
    # 筛选: 实际采购 > 0
    master = master[master["数量"].fillna(0) > 0].copy()

    # 标记漏填
    master["漏填标记"] = master["采购单价"].fillna(0) <= 0
    missing_count = master["漏填标记"].sum()
    if missing_count > 0:
        logger.warning(f"⚠️ 补货表有 {missing_count} 条记录采购价为0（人为漏填），需人工补充！")
        for _, row in master[master["漏填标记"]].iterrows():
            logger.warning(f"  漏填 → {row['来源']}, 货品编码={row['货品编码']}, 数量={row['数量']}")

    logger.info(f"统一补货明细表: {len(master)} 行, 有效={len(master[~master['漏填标记']])} 行, "
                f"唯一SKU={master['货品编码'].nunique()} 个")
    return master


# --- 3b. 加权平均成本回填 (原 cost_filler.py) ---

FILL_METHODS = {"台湾发货", "福州仓库出", "福州仓出货", "厂家出货", "厂家采购",
                "台灣發貨", "福州倉庫出", "廠家出貨"}


def _compute_weighted_avg(replen_df, sku, as_of_date) -> float:
    """计算某SKU截至as_of_date的加权平均采购单价 = Σ合计金额/Σ数量"""
    mask = (
        (replen_df["货品编码"] == sku)
        & (~replen_df["漏填标记"])
        & (replen_df["采购日期"].notna())
        & (replen_df["采购日期"] <= as_of_date)
    )
    subset = replen_df[mask]
    if subset.empty:
        return 0.0
    total_amount = subset["合计金额"].fillna(0).sum()
    total_qty = subset["数量"].fillna(0).sum()
    return round(total_amount / total_qty, 2) if total_qty > 0 else 0.0


def fill_purchase_costs(
    purchase_df: pd.DataFrame,
    replen_df: pd.DataFrame,
    method_col: str = "采购方式",
    cost_col: str = "cost_price",
    sku_col: str = "sku_code",
    date_col: str = "date",
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    回填采购表中缺失的成本价（在 COLUMN_MAP 重命名之后的内部列名上操作）。
    返回 (回填后的df, 回填报告df)。
    """
    df = purchase_df.copy()

    # 标准化 SKU
    df[sku_col] = df[sku_col].apply(lambda v: _norm_sku(v) if pd.notna(v) else "")

    # 解析日期
    df[date_col] = df[date_col].apply(_parse_excel_date)

    # 找需要回填的行
    need_fill = (
        df[method_col].isin(FILL_METHODS)
        & (df[cost_col] <= 0)
        & (df[sku_col] != "")
    )
    fill_indices = df[need_fill].index
    logger.info(f"成本回填: 需要处理 {len(fill_indices)} 行")

    if len(fill_indices) == 0 or replen_df.empty:
        return df, pd.DataFrame()

    cache = {}
    report_rows = []
    filled = 0

    for idx in fill_indices:
        sku = df.at[idx, sku_col]
        pdate = df.at[idx, date_col]

        if pd.isna(pdate):
            report_rows.append({"行号": idx, "货品编码": sku, "状态": "跳过-无采购日期", "回填单价": 0})
            continue

        key = (sku, pdate)
        if key not in cache:
            cache[key] = _compute_weighted_avg(replen_df, sku, pdate)
        avg = cache[key]

        if avg > 0:
            df.at[idx, cost_col] = avg
            filled += 1
            status = "已回填"
        else:
            status = "未匹配"

        report_rows.append({
            "行号": idx, "货品编码": sku, "采购方式": df.at[idx, method_col],
            "商品ID": df.at[idx, "style_id"] if "style_id" in df.columns else "",
            "采购日期": pdate, "状态": status, "回填单价": avg,
        })

    report_df = pd.DataFrame(report_rows)
    logger.info(f"成本回填完成: 成功={filled}, 未匹配={len(fill_indices)-filled}, 总={len(fill_indices)}")
    return df, report_df


# ============================================================
# 4. 数据读取与清洗
# ============================================================

def _detect_header_row(df_raw: pd.DataFrame, max_scan: int = 10) -> int:
    required_names = set(COLUMN_MAP.keys())
    best_row, best_score = 0, 0
    for i in range(min(max_scan, len(df_raw))):
        row_vals = {str(v).strip() for v in df_raw.iloc[i] if pd.notna(v)}
        score = len(row_vals & required_names)
        if score > best_score:
            best_score, best_row = score, i
    return best_row


def read_file(filepath: str) -> pd.DataFrame:
    ext = Path(filepath).suffix.lower()
    name = Path(filepath).name
    if ext in (".xlsx", ".xls"):
        engine = "openpyxl" if ext == ".xlsx" else "xlrd"
        df_raw = pd.read_excel(filepath, engine=engine, header=None, nrows=10)
        header_row = _detect_header_row(df_raw)
        df = pd.read_excel(filepath, engine=engine, header=header_row)
        logger.info(f"读取 {name}，表头行={header_row + 1}，数据行={len(df)}")
        return df
    elif ext == ".csv":
        for enc in ["utf-8-sig", "utf-8", "gbk", "gb2312", "gb18030"]:
            try:
                df_raw = pd.read_csv(filepath, encoding=enc, header=None, nrows=10)
                header_row = _detect_header_row(df_raw)
                df = pd.read_csv(filepath, encoding=enc, header=header_row)
                logger.info(f"读取 {name}（CSV, {enc}），表头行={header_row + 1}，数据行={len(df)}")
                return df
            except (UnicodeDecodeError, UnicodeError):
                continue
        raise ValueError(f"无法解码CSV: {filepath}")
    else:
        raise ValueError(f"不支持的格式: {ext}")


def load_data(input_path: str) -> pd.DataFrame:
    path = Path(input_path)
    if path.is_file():
        frames = [read_file(str(path))]
    elif path.is_dir():
        files = []
        for pat in ["*.xlsx", "*.xls", "*.csv"]:
            files.extend(glob.glob(str(path / pat)))
        files = sorted(set(files))
        if not files:
            raise FileNotFoundError(f"目录 {input_path} 下未找到数据文件")
        frames = [read_file(f) for f in files]
    else:
        raise FileNotFoundError(f"路径不存在: {input_path}")
    df = pd.concat(frames, ignore_index=True)
    logger.info(f"合并后总行数: {len(df)}")
    return df


# ── 必要字段映射（v3: 增加采购方式列） ──
# 注意: A组成本列='1-实际成本价', B组='商品单价'或'实际采购价'
# 通过 _adapt_columns 在映射前统一
COLUMN_MAP = {
    "采购日期":     "date",
    "订单号":       "order_id",
    "店铺名":       "store",
    "商品ID":       "style_id",
    "货品编码":     "sku_code",
    "2--商品价格":  "item_price_twd",
    "实付金额":     "paid_amount_twd",
    "实际采购价":   "cost_price",         # 统一后的成本列名
    "采购方式":     "purchase_method",     # v3新增: 回填用
}


def _adapt_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    适配A/B组列名差异，统一为 COLUMN_MAP 可识别的列名。
    当A/B组concat后，同一语义有多个列（如 '1-实际成本价' 和 '商品单价'），
    需要逐行合并（coalesce）到统一列名。
    """
    df = df.copy()

    # ── 成本列合并 → '实际采购价' ──
    cost_sources = ["实际采购价", "1-实际成本价", "商品单价"]
    existing_cost = [c for c in cost_sources if c in df.columns]
    if existing_cost:
        # coalesce: 按优先级取第一个非空值
        merged = df[existing_cost[0]]
        for col in existing_cost[1:]:
            merged = merged.fillna(df[col])
        df["实际采购价"] = merged
        # 清理冗余列
        for col in existing_cost:
            if col != "实际采购价" and col in df.columns:
                df = df.drop(columns=[col])
        logger.info(f"成本列合并: {existing_cost} → '实际采购价'")

    # ── 采购方式合并 → '采购方式' ──
    method_sources = ["采购方式", "3-采购方式"]
    existing_method = [c for c in method_sources if c in df.columns]
    if existing_method:
        merged = df[existing_method[0]]
        for col in existing_method[1:]:
            merged = merged.fillna(df[col])
        df["采购方式"] = merged
        for col in existing_method:
            if col != "采购方式" and col in df.columns:
                df = df.drop(columns=[col])
        logger.info(f"采购方式合并: {existing_method} → '采购方式'")

    return df


def clean_data(
    df: pd.DataFrame,
    fx_service: ExchangeRateService,
    replen_master: Optional[pd.DataFrame] = None,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    数据清洗（v3）：
    1. 列名适配 → 字段映射
    2. 基础过滤
    3. 按订单分摊销售价(TWD)
    4. ★ 成本回填（补货备货加权平均） ★  ← v3新增
    5. 过滤成本价仍为0的行
    6. 日期解析 → 汇率 → CNY → 毛利润

    Returns: (清洗后df, 回填报告df)
    """
    # ── 列名适配 ──
    df = _adapt_columns(df)

    # ── 校验 & 重命名 ──
    missing = [c for c in COLUMN_MAP if c not in df.columns]
    if missing:
        raise KeyError(f"缺少必要字段: {missing}。实际字段: {list(df.columns)}")
    df = df.rename(columns=COLUMN_MAP).copy()

    # ── 过滤空店铺 ──
    before = len(df)
    df = df[df["store"].notna() & (df["store"].str.strip() != "")]
    logger.info(f"过滤空店铺: {before} → {len(df)} 行")

    # ── 数值修正 ──
    for col in ["item_price_twd", "paid_amount_twd", "cost_price"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # ── 过滤标价/实付异常 ──
    before = len(df)
    df = df[(df["item_price_twd"] > 0) & (df["paid_amount_twd"] > 0)].copy()
    logger.info(f"过滤标价/实付≤0: {before} → {len(df)} 行")

    df["qty"] = 1

    # ── 按订单分摊销售价(TWD) ──
    order_total_price = df.groupby("order_id")["item_price_twd"].transform("sum")
    df["sale_price_twd"] = (df["item_price_twd"] / order_total_price * df["paid_amount_twd"]).round(2)

    # ── ★ v3: 成本回填 ★ ──
    fill_report = pd.DataFrame()
    if replen_master is not None and not replen_master.empty:
        logger.info("=" * 40)
        logger.info("开始成本回填（补货备货加权平均法）")
        df, fill_report = fill_purchase_costs(
            df, replen_master,
            method_col="purchase_method",
            cost_col="cost_price",
            sku_col="sku_code",
            date_col="date",
        )
        logger.info("=" * 40)
    else:
        logger.warning("未提供补货明细表，跳过成本回填")

    # ── 过滤成本价仍为0的行（回填后） ──
    before = len(df)
    df = df[df["cost_price"] > 0].copy()
    excluded = before - len(df)
    logger.info(f"过滤成本价为0: {before} → {len(df)} 行（剔除 {excluded} 行）")

    # ── 日期解析 ──
    df["date"] = df["date"].apply(_parse_excel_date)

    if pd.api.types.is_datetime64_any_dtype(df["date"]):
        df["year"] = df["date"].dt.year
        df["month"] = df["date"].dt.month
        df["day"] = df["date"].dt.day
    else:
        df["date"] = df["date"].astype(str)
        df["month"] = df["date"].str.extract(r"(\d+)月")[0].astype(int)
        df["day"] = df["date"].str.extract(r"月(\d+)日")[0].astype(int)
        df["year"] = DEFAULT_YEAR

    df["full_date"] = df.apply(
        lambda r: f"{int(r['year'])}-{int(r['month']):02d}-{int(r['day']):02d}", axis=1
    )
    df["year_month"] = df.apply(
        lambda r: f"{int(r['year'])}-{int(r['month']):02d}", axis=1
    )

    df["style_id"] = df["style_id"].apply(
        lambda x: str(int(x)) if pd.notna(x) and x == x else ""
    )

    # ── 汇率 ──
    unique_dates = df["full_date"].unique().tolist()
    fx_service.batch_fetch(unique_dates)
    df["fx_rate"] = df["full_date"].apply(lambda d: fx_service.get_rate(d))
    df["sale_price"] = (df["sale_price_twd"] * df["fx_rate"]).round(2)
    df["profit"] = (df["sale_price"] - df["cost_price"]).round(2)

    max_month = df["month"].max()
    max_day = df[df["month"] == max_month]["day"].max()
    df.attrs["data_cutoff"] = f"{DEFAULT_YEAR}年{max_month}月{max_day}日"

    logger.info(f"清洗完成: {len(df)} 行, GMV(CNY)=¥{df['sale_price'].sum():,.2f}, "
                f"利润=¥{df['profit'].sum():,.2f}")

    return df, fill_report


# ============================================================
# 5. 款式聚合与分类 (与v2一致)
# ============================================================

def aggregate_styles(df: pd.DataFrame) -> pd.DataFrame:
    agg = df.groupby(["store", "style_id", "year_month"]).agg(
        gmv=("sale_price", "sum"), qty=("qty", "sum"),
        profit=("profit", "sum"), cost=("cost_price", "sum"),
    ).reset_index()
    agg["margin_rate"] = agg.apply(
        lambda r: round(r["profit"] / r["gmv"] * 100, 2) if r["gmv"] > 0 else 0, axis=1
    )
    store_month_qty = agg.groupby(["store", "year_month"])["qty"].transform("sum")
    agg["qty_share"] = round(agg["qty"] / store_month_qty * 100, 2)
    return agg


def get_store_target_margin(store_name, configs):
    for cfg in configs:
        if cfg.store_name == store_name:
            return cfg.target_margin
    return GLOBAL_MARGIN_WARNING


def classify_styles(styles_df, store_configs):
    df = styles_df.copy()
    df["target_margin"] = df["store"].apply(lambda s: get_store_target_margin(s, store_configs))
    upper = df["target_margin"] + 10
    lower = df["target_margin"] - 10
    df["category"] = "调整款"
    df.loc[(df["margin_rate"] >= lower) & (df["margin_rate"] <= upper), "category"] = "基础款"
    df.loc[(df["margin_rate"] <= lower) & (df["qty_share"] >= TRAFFIC_QTY_SHARE_THRESHOLD), "category"] = "流量款"
    df.loc[df["margin_rate"] >= upper, "category"] = "利润款"
    return df


# ============================================================
# 6. 汇总 & 输出 (与v2一致)
# ============================================================

class NumpyEncoder(json.JSONEncoder):
    def default(self, obj):
        import numpy as np
        if isinstance(obj, (np.bool_,)):   return bool(obj)
        if isinstance(obj, (np.integer,)): return int(obj)
        if isinstance(obj, (np.floating,)):return float(obj)
        if isinstance(obj, (np.ndarray,)): return obj.tolist()
        return super().default(obj)


def compute_dashboard_data(df, styles_df, store_configs):
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
        },
        "raw_styles": [],
        "summary": {},
    }

    for _, r in styles_df.iterrows():
        result["raw_styles"].append({
            "store": r["store"], "style_id": r["style_id"], "year_month": r["year_month"],
            "gmv": round(float(r["gmv"]), 2), "qty": int(r["qty"]),
            "profit": round(float(r["profit"]), 2), "margin_rate": round(float(r["margin_rate"]), 2),
            "qty_share": round(float(r["qty_share"]), 2), "category": r["category"],
            "target_margin": round(float(r["target_margin"]), 2),
        })

    def summarize(subset):
        total_gmv = round(float(subset["gmv"].sum()), 2)
        total_qty = int(subset["qty"].sum())
        total_profit = round(float(subset["profit"].sum()), 2)
        total_margin = round(total_profit / total_gmv * 100, 2) if total_gmv > 0 else 0
        kpi = {"gmv": total_gmv, "qty": total_qty, "style_count": int(subset["style_id"].nunique()),
               "margin_rate": total_margin, "margin_warning": bool(total_margin < GLOBAL_MARGIN_WARNING)}
        categories = {}
        for cat in ["利润款", "流量款", "基础款", "调整款"]:
            cd = subset[subset["category"] == cat]
            cg = round(float(cd["gmv"].sum()), 2)
            cq = int(cd["qty"].sum())
            cp = round(float(cd["profit"].sum()), 2)
            cm = round(cp / cg * 100, 2) if cg > 0 else 0
            th = CATEGORY_MARGIN_THRESHOLDS.get(cat, 0)
            categories[cat] = {
                "style_count": int(cd["style_id"].nunique()), "gmv": cg,
                "gmv_share": round(cg / total_gmv * 100, 2) if total_gmv > 0 else 0,
                "qty": cq, "qty_share": round(cq / total_qty * 100, 2) if total_qty > 0 else 0,
                "margin_rate": cm, "margin_threshold": th,
                "margin_warning": bool(int(cd["style_id"].nunique()) > 0 and cm < th),
            }
        return {"kpi": kpi, "categories": categories}

    result["summary"]["全部"] = summarize(styles_df)
    result["summary"]["按店铺"] = {s: summarize(styles_df[styles_df["store"] == s]) for s in styles_df["store"].unique()}
    result["summary"]["按月份"] = {ym: summarize(styles_df[styles_df["year_month"] == ym]) for ym in sorted(styles_df["year_month"].unique())}
    return result


def save_json(data, output_path):
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, cls=NumpyEncoder)
    logger.info(f"JSON已保存: {output_path}")


def _save_reports(report_dir: Path, fill_report: pd.DataFrame, replen_master: pd.DataFrame):
    """输出回填报告 + 待人工处理清单，无论是否有回填数据都尝试生成"""
    report_dir.mkdir(parents=True, exist_ok=True)

    # 回填报告
    if not fill_report.empty:
        fill_report.to_excel(str(report_dir / "回填报告.xlsx"), index=False)
        logger.info(f"回填报告已保存: {report_dir / '回填报告.xlsx'}")

    # 待人工处理清单（始终生成）
    unmatched = pd.DataFrame(columns=["货品编码", "采购方式", "商品ID"])
    if not fill_report.empty:
        unmatched = fill_report[fill_report["状态"] == "未匹配"][
            ["货品编码", "采购方式", "商品ID"]
        ].drop_duplicates()

    missing_price = pd.DataFrame(columns=["来源", "采购日期", "货品编码", "数量"])
    if not replen_master.empty and "漏填标记" in replen_master.columns:
        mp = replen_master[replen_master["漏填标记"]]
        if not mp.empty:
            missing_price = mp[["来源", "采购日期", "货品编码", "数量"]]

    with pd.ExcelWriter(str(report_dir / "待人工处理清单.xlsx")) as w:
        unmatched.to_excel(w, sheet_name="未匹配货品编码", index=False)
        missing_price.to_excel(w, sheet_name="补货表漏填采购价", index=False)
    logger.info(f"待人工处理清单已保存: {report_dir / '待人工处理清单.xlsx'}")


# ============================================================
# 7. 主流程 (v3)
# ============================================================

SCRIPT_DIR = Path(os.path.abspath(__file__)).parent
DEFAULT_INPUT_PATH = SCRIPT_DIR / "uploads_raw"
DEFAULT_OUTPUT_PATH = SCRIPT_DIR / "output" / "dashboard_data.json"
DEFAULT_TARGET_MARGIN = 35.0


@dataclass
class PipelineResult:
    """管道运行结果"""
    dashboard_data: dict
    fill_report: pd.DataFrame
    replen_master: pd.DataFrame


def run_pipeline(
    input_path: str,
    output_path: str,
    replen_a_path: Optional[str] = None,
    replen_b_path: Optional[str] = None,
    store_configs: Optional[List[StoreConfig]] = None,
) -> PipelineResult:
    """
    完整管道 v3:
    读取 → 构建补货明细 → 清洗(含成本回填+汇率) → 聚合 → 分类 → 汇总 → JSON + 报告
    """
    if store_configs is None:
        store_configs = DEFAULT_STORE_CONFIGS

    logger.info("=" * 60)
    logger.info("电商数据管道 v3 启动")
    logger.info("=" * 60)

    # Step 1: 构建统一补货明细表
    replen_master = build_replenishment_master(replen_a_path, replen_b_path)

    # Step 2: 读取采购表
    raw_df = load_data(input_path)

    # Step 3: 清洗（含成本回填 + 汇率）
    fx_service = ExchangeRateService()
    try:
        clean_df, fill_report = clean_data(raw_df, fx_service, replen_master)
    finally:
        fx_service.close()

    # Step 4: 聚合
    styles_df = aggregate_styles(clean_df)

    # Step 5: 分类
    styles_df = classify_styles(styles_df, store_configs)

    # Step 6: 汇总 & 输出
    dashboard_data = compute_dashboard_data(clean_df, styles_df, store_configs)
    save_json(dashboard_data, output_path)

    # Step 7: 输出回填报告和待处理清单
    report_dir = Path(output_path).parent
    _save_reports(report_dir, fill_report, replen_master)

    # 摘要
    s = dashboard_data["summary"]["全部"]
    logger.info("-" * 40)
    logger.info("管道运行完成：")
    logger.info(f"  数据截止日: {dashboard_data['meta']['data_cutoff']}")
    logger.info(f"  总GMV(CNY): ¥{s['kpi']['gmv']:,.2f}")
    logger.info(f"  总销量: {s['kpi']['qty']:,} 件")
    logger.info(f"  综合毛利率: {s['kpi']['margin_rate']}%")
    if not fill_report.empty:
        filled_count = (fill_report["状态"] == "已回填").sum()
        total_fill = len(fill_report)
        logger.info(f"  成本回填: {filled_count}/{total_fill} 成功")
    logger.info("=" * 60)

    return PipelineResult(dashboard_data, fill_report, replen_master)


def main():
    parser = argparse.ArgumentParser(description="电商数据管道 v3（含补货成本回填）")
    parser.add_argument("--input", "-i", default=None, help="采购表路径或目录")
    parser.add_argument("--output", "-o", default=None, help="输出JSON路径")
    parser.add_argument("--replen-a", default=None, help="A组备货补货申请表路径")
    parser.add_argument("--replen-b", default=None, help="B组补货申请表路径")
    parser.add_argument("--target-margin", type=float, default=None, help="统一目标毛利率(%)")
    args = parser.parse_args()

    input_path = Path(args.input) if args.input else DEFAULT_INPUT_PATH
    output_path = Path(args.output) if args.output else DEFAULT_OUTPUT_PATH
    target_margin = args.target_margin if args.target_margin is not None else DEFAULT_TARGET_MARGIN

    if not input_path.exists():
        logger.error(f"输入路径不存在: {input_path}")
        return

    configs = DEFAULT_STORE_CONFIGS
    if target_margin is not None:
        configs = [StoreConfig(c.store_name, target_margin) for c in configs]

    run_pipeline(
        str(input_path), str(output_path),
        replen_a_path=args.replen_a,
        replen_b_path=args.replen_b,
        store_configs=configs,
    )


if __name__ == "__main__":
    main()
