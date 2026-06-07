"""
电商店铺数据看板 — 数据清洗与分类计算管道 v4
=============================================
v4 更新（基于v3）：
- 新增: 商品资料库价格表查找层（货品编码 → 成本价），作为补货加权平均之后、人工台账之前的自动兜底
- 新增: 商品资料库价格表持久化缓存（output/商品资料库价格表.json），replace 模式归档后仍可用
- 保留: v3 全部逻辑不变（补货加权平均、人工台账、待人工清单、看板JSON）

成本回填优先级（自动在前，人工在后，使流到人工的行最少）：
  ① 采购表原生成本 → ② 补货表加权平均 → ③ 商品资料库 → ④ 人工台账 → 待人工处理

数据流:
  补货备货申请表(A组+B组)        商品资料库(货品编码→成本价)
       ↓ cost_collector              ↓ build_catalog_costmap
  统一补货明细表 ──┐          ┌── 价格字典 {sku: cost}
                   ↓          ↓
  采购表 → clean_data(②加权 → ③资料库 → ④人工) → 聚合 → 分类 → JSON看板
                                  ↓
                            待人工处理清单(仍未匹配的行)

运行方式:
  python ecom_pipeline.py --input ./uploads_raw --output ./output/dashboard_data.json \
      --replen-a ./data/A组备货补货申请表.xlsx --replen-b ./data/B组补货申请表.xlsx \
      --catalog ./data/A组商品资料库.xlsx
  （以上路径均可省略；不传时自动在 uploads_raw 目录按文件名关键词发现）

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
    df = pd.read_excel(filepath, sheet_name=sheet_name, header=2, engine="openpyxl")
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
    """计算某SKU截至as_of_date的加权平均采购单价 = Σ(采购单价×数量)/Σ数量
    ★ v3.1 修正：原用Σ合计金额/Σ数量，但补货表有7条合计金额=0（人工漏填），
      导致均价被拉低50%。改用采购单价×数量避免依赖可能漏填的合计金额字段。
    """
    mask = (
        (replen_df["货品编码"] == sku)
        & (~replen_df["漏填标记"])
        & (replen_df["采购日期"].notna())
        & (replen_df["采购日期"] <= as_of_date)
    )
    subset = replen_df[mask]
    if subset.empty:
        return 0.0
    total_amount = (subset["采购单价"].fillna(0) * subset["数量"].fillna(0)).sum()
    total_qty = subset["数量"].fillna(0).sum()
    return round(total_amount / total_qty, 2) if total_qty > 0 else 0.0


# --- 3c. 商品资料库价格表 (v4 新增) ---
#
# 设计（第一性原理 / 奥卡姆）：
#   商品资料库 = 一本「价格字典」，唯一职责是 货品编码 → 成本价。
#   它没有采购批次/数量语义（不像补货表要加权平均），所以最简表示就是
#   {货品编码: 成本价}，不引入任何多余结构。
#
#   插入位置（逆向：让最少的行流到人工）：
#     ① 采购表原生 → ② 补货加权平均 → ③ 商品资料库(本模块) → ④ 人工台账 → 待人工处理
#   资料库命中的行状态=「资料库匹配」(≠未匹配)，故被现有清单逻辑自动排除，
#   无需改动待人工清单生成代码。
#
#   持久化（沿用台账设计语言）：
#     解析后把价格表缓存到 output/商品资料库价格表.json。app.py 的 replace 模式
#     每次上传会归档 uploads_raw，资料库 3641 行若每次都要重传必丢；有缓存兜底后，
#     没传资料库文件时自动读缓存，跨上传不丢。
#
#   重复货品编码：同一编码多条不同成本 → 取最新日期那条（最贴近当前成本）；
#                无日期则后出现者覆盖前者。确定性、无歧义。

CATALOG_KEYWORDS       = ("商品资料", "商品資料", "资料库", "資料庫")
CATALOG_SKU_COL        = "货品编码"
CATALOG_COST_COL       = "成本价"
CATALOG_DATE_COL       = "日期"
CATALOG_CACHE_FILENAME = "商品资料库价格表.json"


def _read_catalog_sheet(filepath: str) -> pd.DataFrame:
    """读取商品资料库 sheet：优先含关键词的 sheet，表头默认第3行(header=2)，
    缺列时自动扫描前10行兜底定位表头。"""
    xl = pd.ExcelFile(filepath, engine="openpyxl")
    sheet = next((s for s in xl.sheet_names
                  if any(k in s for k in CATALOG_KEYWORDS)), xl.sheet_names[0])
    df = pd.read_excel(filepath, sheet_name=sheet, header=2, engine="openpyxl")
    if CATALOG_SKU_COL in df.columns and CATALOG_COST_COL in df.columns:
        return df
    # 兜底：表头不在第3行 → 扫描定位
    raw = pd.read_excel(filepath, sheet_name=sheet, header=None, nrows=10, engine="openpyxl")
    for i in range(len(raw)):
        names = {str(v).strip() for v in raw.iloc[i] if pd.notna(v)}
        if CATALOG_SKU_COL in names and CATALOG_COST_COL in names:
            return pd.read_excel(filepath, sheet_name=sheet, header=i, engine="openpyxl")
    return df


def build_catalog_costmap(filepath: Optional[str], cache_dir: Optional[Path] = None) -> Dict[str, float]:
    """
    构建商品资料库价格表: {货品编码(标准化): 成本价}。
    优先级：传了文件→解析并刷新缓存；没传文件→读持久化缓存；都没有→空。
    任何异常都降级为空表（绝不让资料库问题阻断主流程）。
    """
    cache_path = (cache_dir / CATALOG_CACHE_FILENAME) if cache_dir else None

    # ── 无文件：尝试读持久化缓存（replace 模式下资料库被归档后仍可用）──
    if not filepath or not Path(filepath).exists():
        if cache_path and cache_path.exists():
            try:
                with open(cache_path, "r", encoding="utf-8") as f:
                    raw = json.load(f)
                cm = {_norm_sku(k): round(float(v), 2)
                      for k, v in raw.items() if float(v) > 0}
                logger.info(f"商品资料库价格表(缓存)已载入: {len(cm)} 个货品编码")
                return cm
            except Exception as e:
                logger.warning(f"商品资料库缓存损坏，按空处理: {e}")
        logger.info("未提供商品资料库，跳过资料库查找")
        return {}

    # ── 有文件：解析 ──
    try:
        df = _read_catalog_sheet(filepath)
    except Exception as e:
        logger.warning(f"商品资料库读取失败，跳过: {e}")
        return {}

    if CATALOG_SKU_COL not in df.columns or CATALOG_COST_COL not in df.columns:
        logger.warning(f"商品资料库缺少必要列({CATALOG_SKU_COL}/{CATALOG_COST_COL})，跳过。"
                       f"实际列: {list(df.columns)}")
        return {}

    out = pd.DataFrame({
        "sku":  df[CATALOG_SKU_COL].apply(_norm_sku),
        "cost": pd.to_numeric(df[CATALOG_COST_COL], errors="coerce"),
    })
    out["date"] = (df[CATALOG_DATE_COL].apply(_parse_excel_date)
                   if CATALOG_DATE_COL in df.columns else pd.NaT)

    # 只留有效行：编码非空 + 成本>0
    out = out[(out["sku"] != "") & (out["cost"].fillna(0) > 0)].copy()
    if out.empty:
        logger.warning("商品资料库无有效记录（编码空或成本≤0）")
        return {}

    # 重复编码：按日期升序，最新(末位)覆盖前者
    out = out.sort_values("date", na_position="first")
    dup = int(out["sku"].duplicated(keep="last").sum())
    costmap = {sku: round(float(c), 2) for sku, c in zip(out["sku"], out["cost"])}
    if dup:
        logger.info(f"商品资料库去重: {dup} 条重复货品编码，取最新日期成本")
    logger.info(f"商品资料库价格表: {len(costmap)} 个唯一货品编码")

    # 持久化缓存（跨 replace 上传保留）
    if cache_path is not None:
        try:
            cache_path.parent.mkdir(parents=True, exist_ok=True)
            with open(cache_path, "w", encoding="utf-8") as f:
                json.dump(costmap, f, ensure_ascii=False)
            logger.info(f"商品资料库价格表已缓存: {cache_path}")
        except Exception as e:
            logger.warning(f"商品资料库缓存写入失败(不影响本轮): {e}")

    return costmap


# ── 人工成本台帐（服务器端单一事实来源）──
#
# 设计（第一性原理 / 逆向）：
#   人工成本以 (订单号, 货品编码) 为唯一键 —— 同一货品在不同订单/时间成本可能不同，
#   必须精确到「采购实例」，避免一个价错配到所有同款行。
#
#   台帐是【单一事实来源】，持久化在 output/人工成本台账.json，跨上传保留
#   （app.py 的 replace 模式只归档 uploads_raw/data，不触碰 output，故天然持久）。
#   每轮流程：载入台帐 → 合并本轮上传清单的改动 → 存回 → 套用整份台帐。
#
#   员工回传的「待人工处理清单」用两个 sheet 与台帐交互：
#     · 『未匹配货品编码』：本轮【新填】(成本单价>0 → upsert 进台帐)。
#     · 『已填成本台账』  ：历史台帐镜像，供【修正】——
#                          成本单价>0 → upsert（改价）；留空或<=0 → delete（撤销该条）。
#   于是「过滤已填」(需求3) 与「填错可修正」(需求3) 互不冲突：
#     未匹配 sheet 永远只列没价的；要改已填的就在台帐 sheet 改。
#   且即使员工上传的清单漏掉某些行，台帐仍在服务器端保留，数据不丢失。

LEDGER_FILENAME = "人工成本台账.json"
MANUAL_SHEET    = "未匹配货品编码"
LEDGER_SHEET    = "已填成本台账"
PRICE_COLS      = ["成本单价", "成本單價", "采购单价", "采购价", "成本价"]


def load_ledger(ledger_path) -> Dict[Tuple[str, str], dict]:
    """从磁盘载入台帐。返回 {(订单号, 货品编码): {成本单价, 更新时间, 来源}}。损坏/不存在 → 空。"""
    p = Path(ledger_path)
    if not p.exists():
        return {}
    try:
        with open(p, "r", encoding="utf-8") as f:
            recs = json.load(f)
    except (json.JSONDecodeError, IOError) as e:
        logger.warning(f"人工成本台账损坏，按空处理: {e}")
        return {}

    ledger: Dict[Tuple[str, str], dict] = {}
    for r in recs if isinstance(recs, list) else []:
        oid = _norm_sku(r.get("订单号"))
        sku = _norm_sku(r.get("货品编码"))
        cost = pd.to_numeric(r.get("成本单价"), errors="coerce")
        if oid and sku and pd.notna(cost) and cost > 0:
            ledger[(oid, sku)] = {
                "成本单价": round(float(cost), 2),
                "更新时间": str(r.get("更新时间", "")),
                "来源":     str(r.get("来源", "")),
            }
    logger.info(f"人工成本台账已载入: {len(ledger)} 条")
    return ledger


def save_ledger(ledger_path, ledger: Dict[Tuple[str, str], dict]):
    """将台帐写回磁盘（list of records，按键排序，便于人工稽核与 diff）。"""
    recs = [{
        "订单号": k[0], "货品编码": k[1],
        "成本单价": v["成本单价"], "更新时间": v.get("更新时间", ""), "来源": v.get("来源", ""),
    } for k, v in sorted(ledger.items())]
    Path(ledger_path).parent.mkdir(parents=True, exist_ok=True)
    with open(ledger_path, "w", encoding="utf-8") as f:
        json.dump(recs, f, ensure_ascii=False, indent=2)
    logger.info(f"人工成本台账已保存: {ledger_path} ({len(recs)} 条)")


def _iter_list_rows(filepath: str, sheet: str):
    """逐行产出 (订单号, 货品编码, 成本单价|None)。price=None 表示该行留空。"""
    try:
        df = pd.read_excel(filepath, sheet_name=sheet, engine="openpyxl", dtype=str)
    except Exception as e:
        logger.warning(f"读取 {Path(filepath).name}[{sheet}] 失败: {e}")
        return
    price_col = next((c for c in PRICE_COLS if c in df.columns), None)
    if price_col is None or "货品编码" not in df.columns or "订单号" not in df.columns:
        logger.warning(f"{Path(filepath).name}[{sheet}] 缺少必要列(订单号/货品编码/成本单价)，跳过")
        return
    for _, row in df.iterrows():
        sku = _norm_sku(row.get("货品编码"))
        oid = _norm_sku(row.get("订单号"))
        if not sku or not oid:
            continue
        price = pd.to_numeric(row.get(price_col), errors="coerce")
        yield oid, sku, (None if pd.isna(price) else float(price))


def read_uploaded_manual(input_path: str):
    """
    读取本轮上传的「待人工处理清单」，解析出对台帐的改动。
    返回 (upserts: {(oid,sku): price}, deletes: set[(oid,sku)], src_name: str)。
    无文件/无 sheet → 空改动（绝不报错）。
    """
    path = Path(input_path)
    files: List[str] = []
    if path.is_dir():
        files = [str(f) for f in path.glob("*.xlsx")
                 if ("处理清单" in f.name or "處理清單" in f.name)]
    elif path.is_file() and ("处理清单" in path.name or "處理清單" in path.name):
        files = [str(path)]

    upserts: Dict[Tuple[str, str], float] = {}
    deletes: set = set()
    src_name = ""

    for fp in files:
        src_name = Path(fp).name
        try:
            sheets = pd.ExcelFile(fp, engine="openpyxl").sheet_names
        except Exception as e:
            logger.warning(f"读取人工清单失败 {src_name}: {e}")
            continue

        # 『未匹配货品编码』：仅新填（>0 → upsert）
        if MANUAL_SHEET in sheets:
            for oid, sku, price in _iter_list_rows(fp, MANUAL_SHEET):
                if price is not None and price > 0:
                    upserts[(oid, sku)] = round(price, 2)

        # 『已填成本台账』：修正（>0 → upsert 改价；空/<=0 → delete 撤销）
        if LEDGER_SHEET in sheets:
            for oid, sku, price in _iter_list_rows(fp, LEDGER_SHEET):
                if price is not None and price > 0:
                    upserts[(oid, sku)] = round(price, 2)
                else:
                    deletes.add((oid, sku))

    if upserts or deletes:
        logger.info(f"本轮人工清单改动: 新增/修改 {len(upserts)} 条, 撤销 {len(deletes)} 条")
    return upserts, deletes, src_name


def apply_ledger_changes(ledger: Dict[Tuple[str, str], dict], upserts, deletes, src="") -> Tuple[int, int]:
    """
    将本轮改动合并进台帐（原地修改）。
    冲突处理：先 delete 再 upsert —— 若同键既被撤销又给了正值，以正值为准（修正优先于撤销）。
    返回 (变更条数, 撤销条数)。
    """
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    n_del = 0
    for k in deletes:
        if k in ledger and k not in upserts:   # 给了新值的不算撤销
            del ledger[k]
            n_del += 1
    n_up = 0
    for k, price in upserts.items():
        price = round(float(price), 2)
        old = ledger.get(k, {}).get("成本单价")
        if old != price:                        # 仅在变化时更新时间戳，保留历史填写时间
            ledger[k] = {"成本单价": price, "更新时间": ts, "来源": src}
            n_up += 1
    if n_up or n_del:
        logger.info(f"台账合并完成: 实际变更 {n_up} 条, 撤销 {n_del} 条, 现存 {len(ledger)} 条")
    return n_up, n_del


def ledger_to_costmap(ledger: Dict[Tuple[str, str], dict]) -> Dict[Tuple[str, str], float]:
    """台帐 → 回填用 {(订单号,货品编码): 成本单价}。"""
    return {k: v["成本单价"] for k, v in ledger.items()}


def fill_purchase_costs(
    purchase_df: pd.DataFrame,
    replen_df: pd.DataFrame,
    manual_costs: Optional[Dict[Tuple[str, str], float]] = None,
    catalog_costs: Optional[Dict[str, float]] = None,
    method_col: str = "采购方式",
    cost_col: str = "cost_price",
    sku_col: str = "sku_code",
    date_col: str = "date",
    order_col: str = "order_id",
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    回填采购表中缺失的成本价（在 COLUMN_MAP 重命名之后的内部列名上操作）。

    成本优先级（第一性原理 / 奥卡姆：单一职责，只填 cost<=0 的行）：
      ① 采购表原生成本价 (cost>0)      —— 不动，最高优先（需求4：采购表补上即以它为主）
      ② 补货备货表加权平均             —— 系统自动（按 货品编码 + 采购日期）
      ③ 商品资料库查找 (catalog_costs) —— 系统自动（按 货品编码），v4新增
      ④ 员工人工填写 (manual_costs)    —— 兜底，仅当 ①②③ 都拿不到价时生效

    逆向：自动来源(②③)排在人工(④)前，资料库越全，流到人工的行越少。
    catalog_costs: {货品编码: 成本单价}
    manual_costs:  {(订单号, 货品编码): 成本单价}
    返回 (回填后的df, 回填报告df)。
    """
    manual_costs = manual_costs or {}
    catalog_costs = catalog_costs or {}
    df = purchase_df.copy()

    # 标准化 SKU
    df[sku_col] = df[sku_col].apply(lambda v: _norm_sku(v) if pd.notna(v) else "")

    # 解析日期
    df[date_col] = df[date_col].apply(_parse_excel_date)

    # 找需要回填的行（cost<=0 即 ① 不成立）
    need_fill = (
        df[method_col].isin(FILL_METHODS)
        & (df[cost_col] <= 0)
        & (df[sku_col] != "")
    )
    fill_indices = df[need_fill].index
    logger.info(f"成本回填: 需要处理 {len(fill_indices)} 行 "
                f"(加权平均源={'有' if not replen_df.empty else '无'}, "
                f"资料库={len(catalog_costs)} 条, 人工成本={len(manual_costs)} 条)")

    # 即使无补货源，只要有资料库或人工成本也要继续（员工二次回传 / 资料库兜底场景）
    if len(fill_indices) == 0:
        return df, pd.DataFrame()

    cache = {}
    report_rows = []
    filled_avg = 0
    filled_catalog = 0
    filled_manual = 0

    for idx in fill_indices:
        sku = df.at[idx, sku_col]
        pdate = df.at[idx, date_col]
        oid = _norm_sku(df.at[idx, order_col]) if order_col in df.columns else ""

        # ② 加权平均（无补货源 / 无采购日期 时为 0）
        avg = 0.0
        if not replen_df.empty and pd.notna(pdate):
            key = (sku, pdate)
            if key not in cache:
                cache[key] = _compute_weighted_avg(replen_df, sku, pdate)
            avg = cache[key]

        if avg > 0:
            df.at[idx, cost_col] = avg
            filled_avg += 1
            status, source, applied = "已回填", "加权平均", avg
        else:
            # ③ 商品资料库查找（按 货品编码）
            cat = catalog_costs.get(sku, 0)
            if cat and cat > 0:
                df.at[idx, cost_col] = cat
                filled_catalog += 1
                status, source, applied = "资料库匹配", "商品资料库", cat
            else:
                # ④ 人工填写兜底（精确匹配 订单号 + 货品编码）
                manual = manual_costs.get((oid, sku), 0)
                if manual and manual > 0:
                    df.at[idx, cost_col] = manual
                    filled_manual += 1
                    status, source, applied = "人工已填", "人工填写", manual
                else:
                    status, source, applied = "未匹配", "", 0

        report_rows.append({
            "行号": idx, "货品编码": sku, "采购方式": df.at[idx, method_col],
            "商品ID": df.at[idx, "style_id"] if "style_id" in df.columns else "",
            "来源": df.at[idx, "_source_file"] if "_source_file" in df.columns else "",
            "采购日期": pdate, "订单号": oid,
            "状态": status, "回填来源": source, "回填单价": applied,
        })

    report_df = pd.DataFrame(report_rows)
    unmatched_n = len(fill_indices) - filled_avg - filled_catalog - filled_manual
    logger.info(f"成本回填完成: 加权={filled_avg}, 资料库={filled_catalog}, 人工={filled_manual}, "
                f"未匹配={unmatched_n}, 总={len(fill_indices)}")
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


def load_data(input_path: str, exclude_keywords: Optional[List[str]] = None) -> pd.DataFrame:
    if exclude_keywords is None:
        # 排除补货表 + 商品资料库 + 系统生成的报表（尤其是员工回传的「待人工处理清单」），
        # 避免它们被当成采购表读取而污染数据 / 缺字段崩溃。
        exclude_keywords = ["备货", "补货", "補貨", "处理清单", "處理清單", "回填报告", "回填報告",
                            "资料库", "資料庫", "商品资料", "商品資料"]
    path = Path(input_path)
    if path.is_file():
        df = read_file(str(path))
        df["_source_file"] = path.name
        frames = [df]
    elif path.is_dir():
        files = []
        for pat in ["*.xlsx", "*.xls", "*.csv"]:
            files.extend(glob.glob(str(path / pat)))
        files = sorted(set(files))
        # 排除补货备货表
        before = len(files)
        files = [f for f in files if not any(kw in Path(f).name for kw in exclude_keywords)]
        if before > len(files):
            logger.info(f"已排除 {before - len(files)} 个补货备货文件，避免混入采购数据")
        if not files:
            raise FileNotFoundError(f"目录 {input_path} 下未找到采购表文件")
        for f in files:
            logger.info(f"  采购表: {Path(f).name}")
        frames = []
        for f in files:
            df = read_file(f)
            df["_source_file"] = Path(f).name
            frames.append(df)
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

    # ── 订单状态合并 → '订单状态' ──
    status_sources = ["订单状态", "訂單狀態", "4-订单状态"]
    existing_status = [c for c in status_sources if c in df.columns]
    if existing_status:
        merged = df[existing_status[0]]
        for col in existing_status[1:]:
            merged = merged.fillna(df[col])
        df["订单状态"] = merged
        for col in existing_status:
            if col != "订单状态" and col in df.columns:
                df = df.drop(columns=[col])
        logger.info(f"订单状态合并: {existing_status} → '订单状态'")

    return df


def clean_data(
    df: pd.DataFrame,
    fx_service: ExchangeRateService,
    replen_master: Optional[pd.DataFrame] = None,
    manual_costs: Optional[Dict[Tuple[str, str], float]] = None,
    catalog_costs: Optional[Dict[str, float]] = None,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    数据清洗（v3）：
    1. 列名适配 → 字段映射
    2. 基础过滤
    3. 按订单分摊销售价(TWD)
    4. ★ 成本回填（补货备货加权平均） ★  ← v3新增
    5. 过滤成本价仍为0的行
    6. 日期解析 → 汇率 → CNY → 毛利润

    Returns: (清洗后df, 回填报告df, 采购单价为零df)
    """
    # ── 列名适配 ──
    df = _adapt_columns(df)

    # ── 校验 & 重命名 ──
    missing = [c for c in COLUMN_MAP if c not in df.columns]
    if missing:
        raise KeyError(f"缺少必要字段: {missing}。实际字段: {list(df.columns)}")
    df = df.rename(columns=COLUMN_MAP).copy()

    # ── 过滤空店铺 ──
    # ── 商品ID / 订单号 统一字符串化 ──
    df["order_id"] = df["order_id"].apply(lambda x: _norm_sku(x) if pd.notna(x) else "")
    df["style_id"] = df["style_id"].apply(lambda x: _norm_sku(x) if pd.notna(x) else "")

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

    # ── 过滤无效记录：不采购/未采购/客户取消订单/已取消 ──
    # ★ v3.1 修正：必须在分摊之前过滤，否则被取消的商品会撑大分母，
    #   导致同订单内有效商品的分摊金额偏低（影响158个混合订单）
    EXCLUDE_METHODS = {"不采购/未采购", "客户取消订单"}
    before = len(df)
    mask_method = df["purchase_method"].astype(str).str.strip().isin(EXCLUDE_METHODS)
    mask_status = pd.Series(False, index=df.index)
    if "订单状态" in df.columns:
        CANCEL_KEYWORDS = {"已取消"}
        mask_status = df["订单状态"].astype(str).str.strip().isin(CANCEL_KEYWORDS)
    df = df[~(mask_method | mask_status)].copy()
    logger.info(f"过滤无效记录(不采购/取消): {before} → {len(df)} 行")

    # ── 按订单分摊销售价(TWD) ──
    order_total_price = df.groupby("order_id")["item_price_twd"].transform("sum")
    df["sale_price_twd"] = (df["item_price_twd"] / order_total_price * df["paid_amount_twd"]).round(2)

    # ── ★ v3: 成本回填（加权平均 + 人工兜底）★ ──
    # 注意：无论是否有补货源/人工成本，都要跑回填 —— 因为「识别未匹配行」本身依赖它，
    #       否则首次上传(两者皆无)时候选行会被静默丢弃，无法生成待人工清单。
    manual_costs = manual_costs or {}
    catalog_costs = catalog_costs or {}
    replen_for_fill = replen_master if replen_master is not None else pd.DataFrame()
    logger.info("=" * 40)
    logger.info(f"开始成本回填（加权平均源={'有' if not replen_for_fill.empty else '无'}, "
                f"资料库={len(catalog_costs)} 条, 人工成本={len(manual_costs)} 条）")
    df, fill_report = fill_purchase_costs(
        df, replen_for_fill,
        manual_costs=manual_costs,
        catalog_costs=catalog_costs,
        method_col="purchase_method",
        cost_col="cost_price",
        sku_col="sku_code",
        date_col="date",
        order_col="order_id",
    )
    logger.info("=" * 40)

    # ── 捕获采购单价为0的记录（供待人工处理清单使用） ──
    zero_cost_df = df[df["cost_price"] <= 0][["_source_file", "sku_code", "order_id"]].copy()
    zero_cost_df.columns = ["来源", "货品编码", "订单号"]
    zero_cost_df = zero_cost_df.drop_duplicates()

    # ── 过滤成本价仍为0的行（回填后） ──
    before = len(df)
    df = df[df["cost_price"] > 0].copy()
    excluded = before - len(df)
    logger.info(f"过滤成本价为0: {before} → {len(df)} 行（剔除 {excluded} 行）")

    # ── 空表保护 ──
    # 若本轮所有行都待人工填写（过滤后为空），仍需返回结构完整的空表，
    # 让下游照常生成「待人工处理清单」，否则首月(成本全空)会在日期/汇率步骤崩溃，连清单都产不出。
    if df.empty:
        logger.warning("过滤后无可计算行（可能成本尚未填写），生成空看板 + 待人工清单")
        for col, default in [("year", DEFAULT_YEAR), ("month", 0), ("day", 0)]:
            df[col] = pd.Series(dtype="int64")
        for col in ["full_date", "year_month"]:
            df[col] = pd.Series(dtype="object")
        for col in ["fx_rate", "sale_price", "profit"]:
            df[col] = pd.Series(dtype="float64")
        df.attrs["data_cutoff"] = "暂无可计算数据"
        return df, fill_report, zero_cost_df

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

    return df, fill_report, zero_cost_df


# ============================================================
# 5. 款式聚合与分类 (与v2一致)
# ============================================================

def aggregate_styles(df: pd.DataFrame) -> pd.DataFrame:
    # ★ v3.1 修正：purchase_method 为 NaN 时 groupby 会静默丢弃该行
    df = df.copy()
    df["purchase_method"] = df["purchase_method"].fillna("未知")

    agg = df.groupby(["store", "style_id", "year_month", "purchase_method"]).agg(
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
            "purchase_methods": sorted([str(m) for m in df["purchase_method"].dropna().unique() if str(m).strip()]),
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

    for _, r in styles_df.iterrows():
        result["raw_styles"].append({
            "store": r["store"], "style_id": r["style_id"], "year_month": r["year_month"],
            "purchase_method": str(r.get("purchase_method", "")),
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
        kpi = {"gmv": total_gmv, "qty": total_qty, "total_profit": total_profit,
               "style_count": int(subset["style_id"].nunique()),
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


def _save_reports(report_dir: Path, fill_report: pd.DataFrame, replen_master: pd.DataFrame,
                  zero_cost_df: pd.DataFrame = None, ledger: Dict[Tuple[str, str], dict] = None):
    """输出回填报告 + 待人工处理清单，无论是否有回填数据都尝试生成"""
    report_dir.mkdir(parents=True, exist_ok=True)

    # 回填报告
    if not fill_report.empty:
        fill_report.to_excel(str(report_dir / "回填报告.xlsx"), index=False)
        logger.info(f"回填报告已保存: {report_dir / '回填报告.xlsx'}")

    # ── 待人工处理清单（始终生成）──
    # 「未匹配货品编码」：经 加权平均 + 人工填写 后仍无成本的采购行，需员工手动填写『成本单价』。
    #   · 新增『采购日期』『订单号』(来自采购表) —— 作为人工成本回写的精确匹配键（需求1）。
    #   · 已通过加权平均或人工填写补上成本的行（状态≠未匹配）会被自动过滤掉（需求3：过滤已填过的）。
    UNMATCHED_COLS = ["货品编码", "商品ID", "采购方式", "采购日期", "订单号", "来源", "成本单价"]
    unmatched = pd.DataFrame(columns=UNMATCHED_COLS)
    if not fill_report.empty and "状态" in fill_report.columns:
        um = fill_report[fill_report["状态"] == "未匹配"].copy()
        if not um.empty:
            um["成本单价"] = ""  # 留空，供员工填写
            for c in ["商品ID", "订单号", "来源", "采购方式"]:
                if c not in um.columns:
                    um[c] = ""
            if "采购日期" not in um.columns:
                um["采购日期"] = pd.NaT
            # 订单号 / 货品编码 强制文本，避免 Excel 把长订单号转成科学计数法导致回写失配
            um["订单号"] = um["订单号"].astype(str)
            um["货品编码"] = um["货品编码"].astype(str)
            unmatched = um[UNMATCHED_COLS].drop_duplicates(subset=["订单号", "货品编码"])

    # 未匹配的货品编码集合（用于排除重复呈现）
    unmatched_skus = set(unmatched["货品编码"].unique()) if not unmatched.empty else set()

    missing_price = pd.DataFrame(columns=["来源", "采购日期", "货品编码", "数量"])
    if not replen_master.empty and "漏填标记" in replen_master.columns:
        mp = replen_master[replen_master["漏填标记"]]
        if not mp.empty:
            missing_price = mp[["来源", "采购日期", "货品编码", "数量"]]

    # 采购金额为零：剔除已在"未匹配货品编码"中的记录
    zero_cost = pd.DataFrame(columns=["来源", "货品编码", "订单号"])
    if zero_cost_df is not None and not zero_cost_df.empty:
        zero_cost = zero_cost_df[~zero_cost_df["货品编码"].isin(unmatched_skus)].copy()

    # ── 已填成本台账（台帐镜像，供员工修正/撤销）──
    # 列『成本单价』可改价；留空或填 0/负数 → 下次上传时该条从台帐撤销。
    ledger_cols = ["货品编码", "订单号", "成本单价", "更新时间", "来源"]
    ledger_rows = []
    if ledger:
        for (oid, sku), v in sorted(ledger.items()):
            ledger_rows.append({
                "货品编码": str(sku), "订单号": str(oid),
                "成本单价": v.get("成本单价"), "更新时间": v.get("更新时间", ""), "来源": v.get("来源", ""),
            })
    ledger_df = pd.DataFrame(ledger_rows, columns=ledger_cols)
    if not ledger_df.empty:
        ledger_df["订单号"] = ledger_df["订单号"].astype(str)
        ledger_df["货品编码"] = ledger_df["货品编码"].astype(str)

    with pd.ExcelWriter(str(report_dir / "待人工处理清单.xlsx")) as w:
        unmatched.to_excel(w, sheet_name="未匹配货品编码", index=False)
        ledger_df.to_excel(w, sheet_name="已填成本台账", index=False)
        missing_price.to_excel(w, sheet_name="补货表漏填采购价", index=False)
        zero_cost.to_excel(w, sheet_name="采购金额为零", index=False)
    logger.info(f"待人工处理清单已保存: {report_dir / '待人工处理清单.xlsx'} "
                f"(未匹配 {len(unmatched)} 条, 台账 {len(ledger_df)} 条)")


def _save_purchase_detail(clean_df: pd.DataFrame, report_dir: Path):
    """输出整理完的采购明细表"""
    if clean_df.empty:
        return

    detail = clean_df[[
        "full_date", "order_id", "qty", "sale_price_twd", "store",
        "sku_code", "style_id", "purchase_method", "cost_price",
        "sale_price", "profit",
    ]].copy()

    detail.columns = [
        "采购日期", "订单号", "商品数量", "销售单价(台币)", "店铺名",
        "货品编码", "商品ID", "采购方式", "采购单价",
        "销售单价(人民币)", "毛利",
    ]
    detail["采购金额"] = (detail["采购单价"] * detail["商品数量"]).round(2)
    detail["毛利率"] = detail.apply(
        lambda r: round(r["毛利"] / r["销售单价(人民币)"] * 100, 2)
        if r["销售单价(人民币)"] > 0 else 0, axis=1
    )

    # 重新排列列顺序
    detail = detail[[
        "采购日期", "订单号", "商品数量", "销售单价(台币)", "销售单价(人民币)", "店铺名",
        "货品编码", "商品ID", "采购方式", "采购单价", "采购金额",
        "毛利", "毛利率",
    ]]

    out_path = report_dir / "采购明细表.xlsx"
    detail.to_excel(str(out_path), index=False)
    logger.info(f"采购明细表已保存: {out_path} ({len(detail)} 行)")


def _save_replen_detail(replen_master: pd.DataFrame, report_dir: Path):
    """输出整理过的补货备货明细表"""
    if replen_master.empty:
        return
    # 只取有效记录（非漏填）
    df = replen_master[~replen_master["漏填标记"]].copy() if "漏填标记" in replen_master.columns else replen_master.copy()
    detail = pd.DataFrame({
        "采购日期": df["采购日期"],
        "货品编码": df["货品编码"],
        "实际采购数量": df["数量"],
        "采购单价": df["采购单价"],
        "采购金额": (df["采购单价"].fillna(0) * df["数量"].fillna(0)).round(2),
    })
    out_path = report_dir / "补货备货明细表.xlsx"
    detail.to_excel(str(out_path), index=False)
    logger.info(f"补货备货明细表已保存: {out_path} ({len(detail)} 行)")


# ============================================================
# 7. 主流程 (v3)
# ============================================================

SCRIPT_DIR = Path(os.path.abspath(__file__)).parent
DEFAULT_INPUT_PATH = SCRIPT_DIR / "uploads_raw"
DEFAULT_OUTPUT_PATH = SCRIPT_DIR / "output" / "dashboard_data.json"
DEFAULT_TARGET_MARGIN = 35.0

# ── 补货备货表默认路径（相对于脚本目录） ──
DEFAULT_REPLEN_A_PATH = SCRIPT_DIR / "uploads_raw" / "A组备货补货申请表.xlsx"
DEFAULT_REPLEN_B_PATH = SCRIPT_DIR / "uploads_raw" / "B组补货申请表.xlsx"
# ── 商品资料库默认路径（v4） ──
DEFAULT_CATALOG_PATH = SCRIPT_DIR / "uploads_raw" / "A组商品资料库.xlsx"


def _auto_discover_replen(search_dir: Path) -> Tuple[Optional[str], Optional[str]]:
    """
    在指定目录及其父目录中自动查找补货备货申请表。
    匹配规则：文件名包含关键词。
    """
    replen_a, replen_b = None, None
    search_dirs = [search_dir, search_dir.parent, SCRIPT_DIR, SCRIPT_DIR / "uploads_raw"]

    for d in search_dirs:
        if not d.exists():
            continue
        for f in d.glob("*.xlsx"):
            name = f.name
            if "备货" in name or "補貨" in name or "补货" in name:
                if "A组" in name and not replen_a:
                    replen_a = str(f)
                    logger.info(f"自动发现A组补货表: {f}")
                elif "B组" in name and not replen_b:
                    replen_b = str(f)
                    logger.info(f"自动发现B组补货表: {f}")
        if replen_a and replen_b:
            break

    return replen_a, replen_b


def _auto_discover_catalog(search_dir: Path) -> Optional[str]:
    """在指定目录及其父目录中自动查找商品资料库（文件名含 资料库/商品资料）。"""
    search_dirs = [search_dir, search_dir.parent, SCRIPT_DIR, SCRIPT_DIR / "uploads_raw"]
    for d in search_dirs:
        if not d.exists():
            continue
        for f in d.glob("*.xlsx"):
            if any(k in f.name for k in CATALOG_KEYWORDS):
                logger.info(f"自动发现商品资料库: {f}")
                return str(f)
    return None


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
    catalog_path: Optional[str] = None,
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
    # 如果没有指定路径，先尝试默认路径，再自动扫描
    if not replen_a_path and DEFAULT_REPLEN_A_PATH.exists():
        replen_a_path = str(DEFAULT_REPLEN_A_PATH)
    if not replen_b_path and DEFAULT_REPLEN_B_PATH.exists():
        replen_b_path = str(DEFAULT_REPLEN_B_PATH)
    if not replen_a_path or not replen_b_path:
        auto_a, auto_b = _auto_discover_replen(Path(input_path))
        replen_a_path = replen_a_path or auto_a
        replen_b_path = replen_b_path or auto_b
    if not replen_a_path and not replen_b_path:
        logger.warning("⚠️ 未找到补货备货申请表，成本回填将跳过。"
                       "请通过 --replen-a / --replen-b 指定路径，"
                       "或将文件放在 uploads_raw 目录下。")

    replen_master = build_replenishment_master(replen_a_path, replen_b_path)

    # Step 1.5: 人工成本台帐（服务器端单一事实来源）
    #   载入磁盘台帐 → 合并本轮上传清单的改动(新填/修正/撤销) → 存回 → 套用整份台帐。
    report_dir = Path(output_path).parent
    report_dir.mkdir(parents=True, exist_ok=True)
    ledger_path = report_dir / LEDGER_FILENAME
    ledger = load_ledger(ledger_path)
    upserts, deletes, src_name = read_uploaded_manual(input_path)
    n_up, n_del = apply_ledger_changes(ledger, upserts, deletes, src=src_name)
    if n_up or n_del:
        save_ledger(ledger_path, ledger)
    manual_costs = ledger_to_costmap(ledger)

    # Step 1.6: 商品资料库价格表（v4）
    #   传了文件→解析并刷新缓存；没传→读 output 持久化缓存（replace 模式归档后仍可用）。
    if not catalog_path and DEFAULT_CATALOG_PATH.exists():
        catalog_path = str(DEFAULT_CATALOG_PATH)
    if not catalog_path:
        catalog_path = _auto_discover_catalog(Path(input_path))
    catalog_costs = build_catalog_costmap(catalog_path, cache_dir=report_dir)

    # Step 2: 读取采购表
    raw_df = load_data(input_path)

    # Step 3: 清洗（含成本回填 + 汇率）
    fx_service = ExchangeRateService()
    try:
        clean_df, fill_report, zero_cost_df = clean_data(
            raw_df, fx_service, replen_master,
            manual_costs=manual_costs, catalog_costs=catalog_costs,
        )
    finally:
        fx_service.close()

    # Step 4: 聚合
    styles_df = aggregate_styles(clean_df)

    # Step 5: 分类
    styles_df = classify_styles(styles_df, store_configs)

    # Step 6: 汇总 & 输出
    dashboard_data = compute_dashboard_data(clean_df, styles_df, store_configs)
    save_json(dashboard_data, output_path)

    # Step 7: 输出回填报告和待处理清单（report_dir / ledger 已在 Step 1.5 准备好）
    _save_reports(report_dir, fill_report, replen_master, zero_cost_df, ledger=ledger)

    # Step 8: 输出采购明细表
    _save_purchase_detail(clean_df, report_dir)

    # Step 9: 输出补货备货明细表
    _save_replen_detail(replen_master, report_dir)

    # 摘要
    s = dashboard_data["summary"]["全部"]
    logger.info("-" * 40)
    logger.info("管道运行完成：")
    logger.info(f"  数据截止日: {dashboard_data['meta']['data_cutoff']}")
    logger.info(f"  总GMV(CNY): ¥{s['kpi']['gmv']:,.2f}")
    logger.info(f"  总销量: {s['kpi']['qty']:,} 件")
    logger.info(f"  综合毛利率: {s['kpi']['margin_rate']}%")
    if not fill_report.empty:
        filled_count = (fill_report["状态"] != "未匹配").sum()
        total_fill = len(fill_report)
        catalog_count = (fill_report["回填来源"] == "商品资料库").sum() if "回填来源" in fill_report.columns else 0
        logger.info(f"  成本回填: {filled_count}/{total_fill} 成功（其中资料库 {catalog_count} 条）")
    logger.info("=" * 60)

    return PipelineResult(dashboard_data, fill_report, replen_master)


def main():
    parser = argparse.ArgumentParser(description="电商数据管道 v3（含补货成本回填）")
    parser.add_argument("--input", "-i", default=None, help="采购表路径或目录")
    parser.add_argument("--output", "-o", default=None, help="输出JSON路径")
    parser.add_argument("--replen-a", default=None, help="A组备货补货申请表路径")
    parser.add_argument("--replen-b", default=None, help="B组补货申请表路径")
    parser.add_argument("--catalog", default=None, help="商品资料库路径（货品编码→成本价 查找表）")
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
        global GLOBAL_MARGIN_WARNING
        GLOBAL_MARGIN_WARNING = target_margin
        configs = [StoreConfig(c.store_name, target_margin) for c in configs]
        logger.info(f"目标毛利率: {target_margin}%")

    run_pipeline(
        str(input_path), str(output_path),
        replen_a_path=args.replen_a,
        replen_b_path=args.replen_b,
        catalog_path=args.catalog,
        store_configs=configs,
    )


if __name__ == "__main__":
    main()
