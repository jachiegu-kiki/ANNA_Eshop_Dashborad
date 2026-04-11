"""
电商数据看板 — 轻量级 Web 服务 v2
====================================
完整流程：
  1. 老板上传 Excel 文件（.xlsx / .xls）
  2. 服务器自动转换为 CSV（智能匹配表头和字段）
  3. 运行 ecom_pipeline.py 生成 dashboard_data.json
  4. 看板页面自动显示最新数据

运行: python app.py
"""

import os
import sys
import json
import time
import shutil
import hashlib
import logging
import subprocess
import secrets
from pathlib import Path
from datetime import datetime
from functools import wraps

# ── Flask 检测与导入 ──
try:
    from flask import (
        Flask, request, jsonify, send_from_directory,
        render_template_string, abort, redirect, url_for
    )
except ImportError:
    print("正在安装 Flask...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "flask"])
    from flask import (
        Flask, request, jsonify, send_from_directory,
        render_template_string, abort, redirect, url_for
    )

# ── 确保依赖完整 ──
for pkg in ["pandas", "openpyxl", "xlrd", "requests"]:
    try:
        __import__(pkg)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])

# ── 本地模块（excel_converter 已移除，pipeline 直接读取原始文件） ──

# ── 路径配置 ──
BASE_DIR      = Path(os.path.abspath(__file__)).parent
UPLOAD_DIR    = BASE_DIR / "uploads_raw"    # 老板上传的原始 Excel
DATA_DIR      = BASE_DIR / "data"           # 转换后的 CSV（喂给管道）
OUTPUT_DIR    = BASE_DIR / "output"         # dashboard_data.json + dashboard.html
CACHE_DIR     = BASE_DIR / "cache"
ARCHIVE_DIR   = BASE_DIR / "archive"
LOG_DIR       = BASE_DIR / "logs"
CONFIG_FILE   = BASE_DIR / ".env"
PIPELINE_FILE = BASE_DIR / "ecom_pipeline_v3.py"

for d in [UPLOAD_DIR, DATA_DIR, OUTPUT_DIR, CACHE_DIR, ARCHIVE_DIR, LOG_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# ── 日志 ──
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(str(LOG_DIR / "server.log"), encoding="utf-8"),
    ]
)
logger = logging.getLogger(__name__)


# ============================================================
# 配置管理
# ============================================================

def load_config() -> dict:
    config = {
        "UPLOAD_TOKEN": "",
        "VIEW_PASSWORD": "",
        "PORT": "5000",
        "HOST": "0.0.0.0",
        "TARGET_MARGIN": "35",
    }
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    key, val = line.split("=", 1)
                    config[key.strip()] = val.strip()
    return config


def init_config():
    if CONFIG_FILE.exists():
        return
    token = secrets.token_urlsafe(32)
    password = secrets.token_urlsafe(16)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        f.write("# ============================================\n")
        f.write("# 电商看板 安全配置文件\n")
        f.write("# ⚠️ 此文件包含密钥，请勿分享给他人！\n")
        f.write("# ============================================\n\n")
        f.write(f"UPLOAD_TOKEN={token}\n\n")
        f.write(f"VIEW_PASSWORD={password}\n\n")
        f.write(f"PORT=5000\n")
        f.write(f"HOST=0.0.0.0\n")

    logger.info("=" * 50)
    logger.info("  首次运行，已自动生成安全配置")
    logger.info(f"  配置文件: {CONFIG_FILE}")
    logger.info(f"  上传密钥: {token}")
    logger.info(f"  看板密码: {password}")
    logger.info("  ⚠️ 请妥善保管以上信息！")
    logger.info("=" * 50)


# ============================================================
# Flask 应用
# ============================================================

init_config()
CONFIG = load_config()
app = Flask(__name__)
app.secret_key = secrets.token_hex(16)

ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".xlsm", ".csv"}


def require_token(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        token = request.headers.get("X-Upload-Token", "")
        if not token or token != CONFIG["UPLOAD_TOKEN"]:
            logger.warning(f"无效 Token: {request.remote_addr}")
            return jsonify({"success": False, "error": "无效的上传密钥"}), 403
        return f(*args, **kwargs)
    return decorated


def check_view_auth():
    password = CONFIG.get("VIEW_PASSWORD", "")
    if not password:
        return True
    session_token = request.cookies.get("view_auth", "")
    expected = hashlib.sha256(password.encode()).hexdigest()
    return session_token == expected


# ── 首页 ──
@app.route("/")
def index():
    if not check_view_auth():
        return redirect(url_for("login_page"))
    # 如果还没有数据，显示等待页面
    data_file = OUTPUT_DIR / "dashboard_data.json"
    if not data_file.exists():
        return render_template_string(WAITING_HTML)
    return send_from_directory(str(OUTPUT_DIR), "dashboard.html")


# ── 登录页 ──
WAITING_HTML = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>数据看板 — 等待数据</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@300;400;700&display=swap');
  *{margin:0;padding:0;box-sizing:border-box}
  body{font-family:'Noto Sans SC',sans-serif;background:#0a0e17;color:#e2e8f0;
       min-height:100vh;display:flex;align-items:center;justify-content:center}
  .box{background:#111827;border:1px solid #1e293b;border-radius:16px;
       padding:48px 40px;width:100%;max-width:480px;text-align:center}
  .box h1{font-size:28px;margin-bottom:16px}
  .box p{font-size:14px;color:#94a3b8;margin-bottom:12px;line-height:1.8}
  .steps{text-align:left;background:#0a0e17;border-radius:10px;padding:20px 24px;
         margin:20px 0;font-size:13px;color:#94a3b8;line-height:2}
  .steps b{color:#38bdf8}
  .refresh{display:inline-block;margin-top:16px;padding:10px 24px;
           background:linear-gradient(135deg,#38bdf8,#818cf8);border:none;
           border-radius:8px;color:#0a0e17;font-size:14px;font-weight:700;
           font-family:inherit;cursor:pointer;text-decoration:none}
</style>
<meta http-equiv="refresh" content="15">
</head>
<body>
  <div class="box">
    <h1>📊 数据看板</h1>
    <p>看板已就绪，等待第一次数据上传</p>
    <div class="steps">
      <b>①</b> 把 Excel 文件放入「数据文件夹」<br>
      <b>②</b> 双击「上传数据」脚本<br>
      <b>③</b> 此页面将自动刷新显示看板
    </div>
    <p style="font-size:12px;color:#64748b">页面每 15 秒自动刷新</p>
    <a href="/" class="refresh">刷新</a>
  </div>
</body>
</html>
"""

LOGIN_HTML = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>数据看板 — 登录</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@300;400;700&display=swap');
  *{margin:0;padding:0;box-sizing:border-box}
  body{font-family:'Noto Sans SC',sans-serif;background:#0a0e17;color:#e2e8f0;
       min-height:100vh;display:flex;align-items:center;justify-content:center}
  .box{background:#111827;border:1px solid #1e293b;border-radius:16px;
       padding:48px 40px;width:100%;max-width:400px;text-align:center}
  .box h1{font-size:22px;margin-bottom:8px;
          background:linear-gradient(135deg,#38bdf8,#818cf8);
          -webkit-background-clip:text;-webkit-text-fill-color:transparent}
  .box p{font-size:13px;color:#64748b;margin-bottom:28px}
  .box input{width:100%;padding:12px 16px;background:#0a0e17;border:1px solid #1e293b;
             border-radius:10px;color:#e2e8f0;font-size:15px;font-family:inherit;
             margin-bottom:16px;transition:border-color .2s}
  .box input:focus{outline:none;border-color:#38bdf8;box-shadow:0 0 0 3px rgba(56,189,248,0.1)}
  .box button{width:100%;padding:12px;background:linear-gradient(135deg,#38bdf8,#818cf8);
              border:none;border-radius:10px;color:#0a0e17;font-size:15px;font-weight:700;
              font-family:inherit;cursor:pointer;transition:opacity .2s}
  .box button:hover{opacity:.9}
  .err{color:#f87171;font-size:13px;margin-bottom:12px}
</style>
</head>
<body>
  <div class="box">
    <h1>📊 电商数据看板</h1>
    <p>请输入查看密码</p>
    {% if error %}<div class="err">{{ error }}</div>{% endif %}
    <form method="POST" action="/login">
      <input type="password" name="password" placeholder="密码" autofocus>
      <button type="submit">进入看板</button>
    </form>
  </div>
</body>
</html>
"""

@app.route("/login", methods=["GET"])
def login_page():
    return render_template_string(LOGIN_HTML, error=None)

@app.route("/login", methods=["POST"])
def login_submit():
    password = request.form.get("password", "")
    if password == CONFIG.get("VIEW_PASSWORD", ""):
        resp = redirect(url_for("index"))
        token = hashlib.sha256(password.encode()).hexdigest()
        resp.set_cookie("view_auth", token, max_age=86400*30, httponly=True)
        return resp
    return render_template_string(LOGIN_HTML, error="密码错误，请重试")


@app.route("/dashboard_data.json")
def serve_data():
    if not check_view_auth():
        abort(403)
    data_file = OUTPUT_DIR / "dashboard_data.json"
    if not data_file.exists():
        # 还没有数据时返回空结构，避免前端 404 或 JS 崩溃
        empty_cats = {}
        for cat in ["利润款", "流量款", "基础款", "调整款"]:
            empty_cats[cat] = {
                "style_count": 0, "gmv": 0, "gmv_share": 0,
                "qty": 0, "qty_share": 0, "margin_rate": 0,
                "margin_threshold": 0, "margin_warning": False,
            }
        empty_summary = {
            "kpi": {"gmv": 0, "qty": 0, "style_count": 0, "margin_rate": 0, "margin_warning": False},
            "categories": empty_cats,
        }
        empty = {
            "meta": {"data_cutoff": "暂无数据", "generated_at": "", "stores": [], "months": []},
            "config": {"store_configs": {}, "category_thresholds": {"利润款": 80, "基础款": 60, "流量款": 30, "调整款": 40}, "global_margin_warning": 35, "traffic_qty_share_threshold": 5},
            "raw_styles": [],
            "summary": {"全部": empty_summary, "按店铺": {}, "按月份": {}},
        }
        return jsonify(empty)
    return send_from_directory(str(OUTPUT_DIR), "dashboard_data.json")


@app.route("/api/health")
def health():
    data_file = OUTPUT_DIR / "dashboard_data.json"
    has_data = data_file.exists()
    last_update = ""
    if has_data:
        mtime = os.path.getmtime(str(data_file))
        last_update = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
    return jsonify({
        "status": "ok",
        "has_data": has_data,
        "last_update": last_update,
    })


@app.route("/api/target-margin", methods=["GET", "POST"])
def target_margin_api():
    """获取或设置目标毛利率，并持久化到 .env 配置"""
    if not check_view_auth():
        abort(403)

    if request.method == "GET":
        return jsonify({"target_margin": float(CONFIG.get("TARGET_MARGIN", 35))})

    # POST: 更新目标毛利率
    data = request.get_json(silent=True) or {}
    new_margin = data.get("target_margin")
    if new_margin is None:
        return jsonify({"success": False, "error": "缺少 target_margin"}), 400
    try:
        new_margin = float(new_margin)
        if not (0 <= new_margin <= 100):
            raise ValueError
    except (ValueError, TypeError):
        return jsonify({"success": False, "error": "无效的毛利率值"}), 400

    # 写入 CONFIG 并持久化
    CONFIG["TARGET_MARGIN"] = str(new_margin)
    _save_config()
    logger.info(f"目标毛利率已更新为 {new_margin}%")

    # 如果有数据，自动重跑管道
    pipeline_result = {"skipped": True}
    data_files = [f for f in UPLOAD_DIR.iterdir()
                  if f.is_file() and f.suffix.lower() in ALLOWED_EXTENSIONS]
    if data_files:
        pipeline_result = _run_pipeline()

    return jsonify({
        "success": True,
        "target_margin": new_margin,
        "pipeline": pipeline_result,
    })


def _save_config():
    """将当前 CONFIG 写回 .env 文件"""
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        f.write("# 电商看板 安全配置文件\n")
        for key, val in CONFIG.items():
            f.write(f"{key}={val}\n")


@app.route("/download/<filename>")
def download_report(filename):
    """下载回填报告 / 待人工处理清单等 output 目录下的文件"""
    if not check_view_auth():
        abort(403)
    SAFE_FILES = {"待人工处理清单.xlsx", "回填报告.xlsx", "采购明细表.xlsx"}
    if filename not in SAFE_FILES:
        abort(404)
    filepath = OUTPUT_DIR / filename
    if not filepath.exists():
        abort(404)
    return send_from_directory(str(OUTPUT_DIR), filename, as_attachment=True)


# ============================================================
# 核心：文件上传 → Excel转CSV → 运行管道
# ============================================================

@app.route("/api/upload", methods=["POST"])
@require_token
def upload_files():
    """
    老板上传 Excel → 服务器自动转 CSV → 自动跑管道 → 看板更新
    """
    if "files" not in request.files:
        return jsonify({"success": False, "error": "未找到文件"}), 400

    files = request.files.getlist("files")
    if not files or all(not f.filename for f in files):
        return jsonify({"success": False, "error": "文件列表为空"}), 400

    # ── Step 1: 归档旧数据 ──
    mode = request.form.get("mode", "replace")
    if mode == "replace":
        _archive_old_data()

    # ── Step 2: 保存上传文件到 uploads_raw/ ──
    saved_files = []
    errors = []

    for f in files:
        if not f.filename:
            continue
        ext = Path(f.filename).suffix.lower()
        if ext not in ALLOWED_EXTENSIONS:
            errors.append(
                f"跳过不支持的文件: {f.filename}"
                f"（仅支持 .xlsx .xls .csv）"
            )
            continue
        safe_name = f.filename.replace("/", "_").replace("\\", "_")
        save_path = UPLOAD_DIR / safe_name
        try:
            f.save(str(save_path))
            saved_files.append(safe_name)
            logger.info(f"文件已保存: {save_path}")
        except Exception as e:
            errors.append(f"保存失败 {f.filename}: {str(e)}")

    if not saved_files:
        return jsonify({
            "success": False,
            "error": "没有成功保存任何文件",
            "details": errors,
        }), 400

    # ── Step 3: 运行数据管道（pipeline 直接读取 uploads_raw 中的原始文件） ──
    logger.info("=" * 40)
    logger.info("开始运行数据管道...")
    pipeline_result = _run_pipeline()

    return jsonify({
        "success": pipeline_result["success"],
        "message": (
            f"已处理 {len(saved_files)} 个文件 → "
            f"{'看板已更新' if pipeline_result['success'] else '解析失败'}"
        ),
        "files": saved_files,
        "pipeline": pipeline_result,
        "errors": errors if errors else None,
    })


@app.route("/api/reprocess", methods=["POST"])
@require_token
def reprocess():
    pipeline_result = _run_pipeline()
    return jsonify({"pipeline": pipeline_result})


def _archive_old_data():
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    for src_dir, sub_name in [(UPLOAD_DIR, "uploads"), (DATA_DIR, "data")]:
        old_files = [
            f for f in src_dir.iterdir()
            if f.is_file() and not f.name.startswith(".")
        ]
        if old_files:
            dest = ARCHIVE_DIR / ts / sub_name
            dest.mkdir(parents=True, exist_ok=True)
            for f in old_files:
                shutil.move(str(f), str(dest / f.name))
            logger.info(f"已归档 {len(old_files)} 个文件到 {dest}")


def _run_pipeline() -> dict:
    if not PIPELINE_FILE.exists():
        return {"success": False, "error": "管道脚本不存在"}

    data_files = [
        f for f in UPLOAD_DIR.iterdir()
        if f.is_file() and f.suffix.lower() in ALLOWED_EXTENSIONS
    ]
    if not data_files:
        return {"success": False, "error": "无数据文件可处理"}

    output_file = OUTPUT_DIR / "dashboard_data.json"
    target_margin = CONFIG.get("TARGET_MARGIN", "35")
    cmd = [
        sys.executable, str(PIPELINE_FILE),
        "--input", str(UPLOAD_DIR),
        "--output", str(output_file),
        "--target-margin", str(target_margin),
    ]

    try:
        start = time.time()
        result = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=300,
            cwd=str(BASE_DIR),
        )
        # Python 3.6 兼容：bytes -> str 手动 decode
        stdout_txt = result.stdout.decode("utf-8", errors="replace") if result.stdout else ""
        stderr_txt = result.stderr.decode("utf-8", errors="replace") if result.stderr else ""
        elapsed = round(time.time() - start, 1)

        # 将管道自身的日志转发到服务器日志（方便排查）
        if stderr_txt:
            for line in stderr_txt.strip().splitlines():
                logger.info("[pipeline] %s", line)

        if result.returncode == 0:
            logger.info("管道运行成功，耗时 %ss", elapsed)
            return {
                "success": True,
                "message": "数据解析完成（%ss）" % elapsed,
                "file_count": len(data_files),
            }
        else:
            logger.error("管道运行失败 (exit=%s)", result.returncode)
            return {
                "success": False,
                "error": "数据解析失败",
                "detail": stderr_txt[-800:] if stderr_txt else "",
            }
    except subprocess.TimeoutExpired:
        logger.error("管道执行超时（>300s），已强制终止。请检查网络或数据量。")
        return {"success": False, "error": "解析超时（>300s）"}
    except Exception as e:
        logger.error(f"管道启动异常: {e}")
        return {"success": False, "error": str(e)}


@app.route("/api/status")
@require_token
def status():
    uploads = [f.name for f in UPLOAD_DIR.iterdir() if f.is_file() and not f.name.startswith(".")]
    csvs = [f.name for f in DATA_DIR.glob("*.csv")]
    data_file = OUTPUT_DIR / "dashboard_data.json"
    return jsonify({
        "upload_files": uploads,
        "csv_files": csvs,
        "has_dashboard_data": data_file.exists(),
        "last_update": datetime.fromtimestamp(
            os.path.getmtime(str(data_file))
        ).strftime("%Y-%m-%d %H:%M:%S") if data_file.exists() else None,
    })


if __name__ == "__main__":
    port = int(CONFIG.get("PORT", 5000))
    host = CONFIG.get("HOST", "0.0.0.0")
    logger.info("=" * 50)
    logger.info("  电商数据看板服务 v2 启动")
    logger.info(f"  支持: Excel (.xlsx .xls) + CSV")
    logger.info(f"  地址: http://YOUR_SERVER_IP:{port}/")
    logger.info("=" * 50)
    app.run(host=host, port=port, debug=False)
