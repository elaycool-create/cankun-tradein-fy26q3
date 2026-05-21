#!/usr/bin/env python3
"""
燦坤 Trade-in 舊換新週報 — 一鍵分析 + HTML 報告產生器
Author: Generated for ASC Elay
Usage:
    python3 analyze_and_report.py --input data.xlsx --output report.html
    python3 analyze_and_report.py --input old.xlsx new.xlsx --output report.html
"""

import argparse
import json
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    sys.exit("❌ 需要 pandas：pip3 install pandas openpyxl")


# ────────────────────────────────────────────────────────────────────
# 常數
# ────────────────────────────────────────────────────────────────────
WEEK_START_DEFAULT = "2026-03-29"  # FY26Q3W01 起始日
DEALER_DEFAULT = "燦坤"
NULL_PLACEHOLDERS = {"58", "60", "65", "70"}  # 各次匯出的 null 替代字（自動偵測）

# ── 達標標準（%） ──
RECYCLE_TARGET = 10.0   # 回收率：成交數 ÷ iPhone 銷量 × 100% ≥ 10%
PLUGIN_TARGET  = 15.0   # 接線率：接機數 ÷ iPhone 銷量 × 100% ≥ 15%

IMEI_PATTERN = re.compile(r"IMEI:\d{15}")


# ────────────────────────────────────────────────────────────────────
# 工具函式
# ────────────────────────────────────────────────────────────────────
def detect_null_placeholders(df: pd.DataFrame) -> set:
    """從多個應為 null 的欄位偵測該檔案使用的 null 替代字（可能多個）"""
    found = set()
    # 這些欄位在「未成交訂單」上理論應該是 null
    candidate_cols = ("VAD", "购新LOB", "购新设备型号", "新品序列号",
                      "旧手机IMEI", "dgCode", "implementationId")
    for col in candidate_cols:
        if col in df.columns:
            vals = df[col].dropna().astype(str).unique()
            for v in vals:
                if v in NULL_PLACEHOLDERS:
                    found.add(v)
    return found


def make_clean(null_vals: set):
    """產生清理函式：把空值或任何 null 替代字都歸為空字串"""
    def _clean(x):
        if pd.isna(x):
            return ""
        s = str(x).strip()
        return "" if (s == "" or s in null_vals) else s
    return _clean


def build_week_ranges(start_date: str, num_weeks: int):
    """生成 [(週次代碼, start_ts, end_ts, 顯示日期)…]"""
    base = pd.Timestamp(start_date)
    out = []
    for i in range(num_weeks):
        s = base + timedelta(days=7 * i)
        e = s + timedelta(days=6)
        # 用 W01, W02… 格式（FY/Q 由 start_date 推斷可選擴充）
        wcode = f"FY26Q3W{i+1:02d}"
        disp = f"{s.month}/{s.day}–{e.month}/{e.day}"
        out.append((wcode, s, e, disp))
    return out


def assign_week(ts, week_ranges):
    if pd.isna(ts):
        return None
    nt = ts.normalize()
    for w, s, e, _ in week_ranges:
        if s <= nt <= e:
            return w
    return None


# ────────────────────────────────────────────────────────────────────
# 核心分析
# ────────────────────────────────────────────────────────────────────
def analyze(files, dealer_kw: str, start_date: str, num_weeks: int):
    frames = []
    for f in files:
        print(f"📂 讀取 {f}…")
        df = pd.read_excel(f, dtype=str)
        null_vals = detect_null_placeholders(df)
        clean = make_clean(null_vals)
        df["订单创建时间"] = pd.to_datetime(df["订单创建时间"], errors="coerce")
        # 從 检测信息 抽取 IMEI 碼（15 位數）
        df["imei"] = df["检测信息"].astype(str).str.extract(r'IMEI:(\d{15})', expand=False)
        df["is_deal"] = df["回收类型"] == "以旧换新"
        df["operator"]  = df["XPOS操作人"].apply(clean)
        df["old_model"] = df["旧手机型号"].apply(clean)
        df["new_model"] = df["购新设备型号"].apply(clean)
        df["new_lob"]   = df["购新LOB"].apply(clean)
        # 成交價格：用於獎金計算（只有 is_deal=True 的列才有意義的數值）
        df["deal_price"] = pd.to_numeric(df["成交价格"], errors="coerce").fillna(0).astype(int)
        frames.append(df)
        print(f"   {len(df):,} 筆，null 替代字 = {null_vals or '無'}")

    df_all = pd.concat(frames, ignore_index=True)

    # 過濾經銷商
    mask = df_all["经销商名称"].str.contains(dealer_kw, na=False) | \
           df_all["门店名称"].str.contains(dealer_kw, na=False)
    ck = df_all[mask].copy()
    print(f"🎯 過濾 '{dealer_kw}'：{len(ck):,} 筆")

    # ── 接機數去重：同一 IMEI 只算一次（取最早出現） ────────────
    # 依時間排序後，標記每個 IMEI 第一次出現為 True，重複出現為 False
    ck = ck.sort_values("订单创建时间", kind="stable").reset_index(drop=True)
    ck["is_plugin"] = ck["imei"].notna() & ~ck["imei"].duplicated(keep="first")
    n_imei_records = ck["imei"].notna().sum()
    n_unique_imei  = int(ck["is_plugin"].sum())
    n_dup = n_imei_records - n_unique_imei
    print(f"🔌 IMEI 接機統計：{n_imei_records} 筆原始 → 去重後 {n_unique_imei} 筆（剔除 {n_dup} 筆重複）")

    # 自動推算週數（如使用者沒指定）
    if num_weeks <= 0:
        latest = ck["订单创建时间"].max()
        if pd.isna(latest):
            num_weeks = 6
        else:
            base = pd.Timestamp(start_date)
            num_weeks = max(1, ((latest.normalize() - base).days // 7) + 1)
    print(f"📅 計算 {num_weeks} 週")

    week_ranges = build_week_ranges(start_date, num_weeks)
    weeks = [w[0] for w in week_ranges]
    week_dates = {w[0]: w[3] for w in week_ranges}

    ck["week"] = ck["订单创建时间"].apply(lambda x: assign_week(x, week_ranges))
    ck = ck[ck["week"].notna()].copy()

    stores = sorted(ck["门店名称"].unique())

    # 各門市 × 各週統計
    results = {}
    for store in stores:
        sdf = ck[ck["门店名称"] == store]
        results[store] = {}
        for w in weeks:
            wdf = sdf[sdf["week"] == w]
            deal_df = wdf[wdf["is_deal"]]
            op_deal = {}
            for op in deal_df["operator"]:
                if op:
                    op_deal[op] = op_deal.get(op, 0) + 1
            old_top = [(str(m), int(c)) for m, c in wdf["old_model"].value_counts().head(5).items() if m]
            new_items = []
            if not deal_df.empty:
                nc = deal_df.groupby(["new_lob", "new_model"]).size().sort_values(ascending=False)
                for (lob, model), cnt in nc.items():
                    if model:
                        new_items.append((str(lob), str(model), int(cnt)))
            results[store][w] = {
                "exec":   int(len(wdf)),
                "plugin": int(wdf["is_plugin"].sum()),
                "deal":   int(wdf["is_deal"].sum()),
                "old_top":  old_top,
                "new_items": new_items,
                "op_deal":  sorted(op_deal.items(), key=lambda x: -x[1]),
            }

    weekly_totals = {}
    for w in weeks:
        wdf = ck[ck["week"] == w]
        weekly_totals[w] = {
            "exec":   int(len(wdf)),
            "plugin": int(wdf["is_plugin"].sum()),
            "deal":   int(wdf["is_deal"].sum()),
        }

    all_old = [(str(m), int(c)) for m, c in ck["old_model"].value_counts().head(10).items() if m]
    deal_all = ck[ck["is_deal"]]
    all_new = []
    if not deal_all.empty:
        for (lob, m), c in deal_all.groupby(["new_lob", "new_model"]).size().sort_values(ascending=False).head(10).items():
            if m:
                all_new.append((str(lob), str(m), int(c)))

    global_op_dict = {}
    for store in stores:
        for w in weeks:
            for op, cnt in results[store][w]["op_deal"]:
                global_op_dict[op] = global_op_dict.get(op, 0) + cnt
    global_op = sorted(global_op_dict.items(), key=lambda x: -x[1])

    # ── 獎金計算用：每位操作人的成交記錄 ─────────────────────────
    deal_df = ck[ck["is_deal"] & (ck["operator"] != "")].copy()
    bonus_records = {}
    for _, row in deal_df.iterrows():
        op = row["operator"]
        ts = row["订单创建时间"]
        if pd.isna(ts):
            continue
        bonus_records.setdefault(op, []).append({
            "date":  ts.strftime("%Y-%m-%d"),
            "month": ts.strftime("%Y-%m"),
            "week":  row["week"],
            "store": row["门店名称"],
            "price": int(row["deal_price"]),
        })

    return {
        "results": results, "stores": stores,
        "weeks": weeks, "week_dates": week_dates,
        "all_old": all_old, "all_new": all_new,
        "weekly_totals": weekly_totals, "global_op": global_op,
        "latest_week": weeks[-1] if weeks else None,
        "bonus_records": bonus_records,
    }


# ────────────────────────────────────────────────────────────────────
# HTML 渲染
# ────────────────────────────────────────────────────────────────────
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>{{title}}</title>
<style>
:root{--bg:#f5f5f7;--card:#fff;--text:#1d1d1f;--sub:#6e6e73;--accent:#0071e3;--green:#34c759;--orange:#ff9500;--purple:#af52de;--border:#e5e5ea;--teal:#32ade6;--new:#ff9f0a}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,'SF Pro Display','PingFang TC',sans-serif;background:var(--bg);color:var(--text);font-size:14px}
.header{background:linear-gradient(135deg,#1d1d1f 0%,#0071e3 100%);color:#fff;padding:32px 40px 28px}
.header h1{font-size:28px;font-weight:700;letter-spacing:-.5px}
.header p{color:rgba(255,255,255,.7);margin-top:6px;font-size:14px}
.header-meta{display:flex;gap:14px;margin-top:16px;flex-wrap:wrap}
.meta-badge{background:rgba(255,255,255,.15);border-radius:20px;padding:5px 14px;font-size:13px;font-weight:500}
.container{max-width:1500px;margin:0 auto;padding:32px 24px}
.two-col{display:grid;grid-template-columns:1fr 1fr;gap:20px}
@media(max-width:900px){.two-col{grid-template-columns:1fr}}
.section-title{font-size:20px;font-weight:700;margin:32px 0 14px}
.summary-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:14px;margin-bottom:26px}
.summary-card{background:var(--card);border-radius:16px;padding:20px;box-shadow:0 2px 8px rgba(0,0,0,.06);text-align:center}
.summary-card .label{color:var(--sub);font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.5px}
.summary-card .value{font-size:32px;font-weight:700;margin:6px 0}
.summary-card .sub{color:var(--sub);font-size:12px}
.blue{color:var(--accent)}.green{color:var(--green)}.purple{color:var(--purple)}.teal{color:var(--teal)}
.card{background:var(--card);border-radius:16px;padding:22px;box-shadow:0 2px 8px rgba(0,0,0,.06);margin-bottom:20px}
.card-title{font-weight:700;font-size:15px;margin-bottom:13px}
.data-table{width:100%;border-collapse:collapse;font-size:13px}
.data-table th{background:var(--bg);padding:9px 11px;text-align:left;font-weight:600;font-size:11px;color:var(--sub);border-bottom:1px solid var(--border);white-space:nowrap}
.data-table td{padding:8px 11px;border-bottom:1px solid var(--border);vertical-align:middle}
.data-table tr:last-child td{border-bottom:none}
.data-table tr:hover td{background:rgba(0,113,227,.03)}
.num{text-align:right;font-variant-numeric:tabular-nums;font-weight:600}
.plugin-cell{color:var(--purple)}.deal-yes{color:var(--green);font-weight:700}
.small-text{font-size:12px;color:var(--sub);max-width:190px}.zero-row td{color:#d1d1d6}
.new-week td{background:rgba(255,159,10,.06)!important;border-left:3px solid var(--new)}
.new-week-row{background:rgba(255,159,10,.05)}
.new-week-row td{border-left:2px solid var(--new)}
.new-badge{display:inline-block;background:var(--new);color:#fff;font-size:10px;font-weight:700;padding:1px 6px;border-radius:8px;margin-left:5px;vertical-align:middle}
.op-pill{display:inline-flex;align-items:center;gap:5px;background:rgba(50,173,230,.1);color:var(--teal);border-radius:20px;padding:3px 10px;margin:2px 3px 2px 0;font-size:12px;font-weight:500;white-space:nowrap}
.op-pill .op-cnt{background:var(--teal);color:#fff;border-radius:10px;padding:1px 6px;font-size:11px;font-weight:700}
.op-cell{max-width:300px;line-height:1.9}
.op-summary-bar{padding:9px 20px;background:rgba(50,173,230,.04);border-bottom:1px solid var(--border);display:flex;flex-wrap:wrap;gap:4px}
.op-label{font-size:12px;color:var(--sub);font-weight:600;margin-right:4px;padding-top:4px;white-space:nowrap}
.store-card{background:var(--card);border-radius:16px;box-shadow:0 2px 8px rgba(0,0,0,.06);margin-bottom:18px;overflow:hidden}
.store-header{padding:13px 20px;background:linear-gradient(to right,rgba(0,113,227,.06),transparent);border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:9px}
.store-title{font-size:15px;font-weight:700}
.store-kpis{display:flex;flex-wrap:wrap;gap:7px}
.kpi-chip{border-radius:20px;padding:4px 12px;font-size:12px;font-weight:600}
.exec-chip{background:rgba(0,113,227,.1);color:var(--accent)}
.plugin-chip{background:rgba(175,82,222,.1);color:var(--purple)}
.deal-chip{background:rgba(52,199,89,.1);color:var(--green)}
.sales-row{display:flex;align-items:center;gap:10px;flex-wrap:wrap;padding:11px 20px;background:linear-gradient(to right,rgba(255,204,0,.08),rgba(255,149,0,.04));border-bottom:1px solid var(--border)}
.sales-label{font-size:12px;font-weight:600;color:var(--sub);white-space:nowrap}
.sales-divider{color:#d1d1d6;font-size:14px}
.sales-input{width:120px;padding:6px 12px;border:1.5px solid var(--border);border-radius:20px;font-size:13px;font-weight:600;font-family:inherit;outline:none;text-align:center}
.sales-input:focus{border-color:var(--orange);box-shadow:0 0 0 3px rgba(255,149,0,.15)}
.rate-badge{min-width:64px;text-align:center;font-size:15px;font-weight:800;padding:4px 12px;border-radius:20px;background:#f5f5f7;color:#bbb;transition:all .25s}
.rate-badge.high{background:rgba(52,199,89,.15);color:#25a244}
.rate-badge.mid{background:rgba(255,149,0,.15);color:#c97200}
.rate-badge.low{background:rgba(255,59,48,.12);color:#d63030}
.rate-badge.high::after{content:" ✅";font-size:11px}
.rate-rank-badge{display:inline-block;width:26px;height:26px;line-height:26px;border-radius:50%;text-align:center;font-size:12px;font-weight:700;background:var(--bg);color:var(--sub)}
.rate-rank-badge.r1{background:#FFD700;color:#7a5c00}
.rate-rank-badge.r2{background:#C0C0C0;color:#555}
.rate-rank-badge.r3{background:#CD7F32;color:#fff}
.empty-hint{text-align:center;color:#bbb;padding:24px;font-size:13px}
/* ── 獎金計算機 ── */
.bonus-summary{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;margin-bottom:18px}
.bsum-card{background:#fafafc;border:1px solid var(--border);border-radius:12px;padding:14px;text-align:center}
.bsum-card.highlight{background:linear-gradient(135deg,rgba(255,204,0,.12),rgba(255,149,0,.06));border-color:rgba(255,149,0,.3)}
.bsum-label{color:var(--sub);font-size:11px;font-weight:600}
.bsum-val{font-size:18px;font-weight:700;margin-top:5px}
.bonus-month-card{border:1px solid var(--border);border-radius:14px;padding:16px;margin-bottom:14px;background:#fff}
.bonus-month-head{display:flex;align-items:center;gap:12px;flex-wrap:wrap;padding-bottom:12px;border-bottom:1px solid var(--border);margin-bottom:14px}
.rule-badge{padding:3px 10px;border-radius:14px;font-size:12px;font-weight:600}
.rule-badge.basic{background:rgba(52,199,89,.12);color:#1a7a32}
.rule-badge.adv{background:rgba(255,149,0,.15);color:#c97200}
.bonus-calc-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:14px}
.bcalc{background:#fafafc;border-radius:10px;padding:14px;display:flex;flex-direction:column;gap:6px}
.bcalc.highlight{background:linear-gradient(135deg,rgba(255,204,0,.1),rgba(255,149,0,.04));border:1px solid rgba(255,149,0,.25)}
.bcalc-title{font-weight:700;font-size:13px;margin-bottom:6px}
.bcalc-row{display:flex;justify-content:space-between;font-size:13px;color:var(--text)}
.bcalc-row.total{padding-top:8px;border-top:1px dashed var(--border);margin-top:4px}
.bonus-rules-bar{display:flex;justify-content:space-between;align-items:flex-start;gap:10px;margin-bottom:14px;flex-wrap:wrap}
.bonus-rules-text{display:flex;gap:18px;flex-wrap:wrap;font-size:12px;padding:10px 14px;background:#fff8e6;border-radius:10px;border-left:3px solid var(--orange);flex:1;min-width:260px}
.lock-btn{padding:6px 14px;border:1px solid var(--border);border-radius:20px;background:#fff;font-size:12px;cursor:pointer;color:var(--sub);white-space:nowrap}
.lock-btn:hover{background:#f5f5f7}
.rate-rules-bar{display:flex;gap:18px;flex-wrap:wrap;font-size:12px;margin-bottom:14px;padding:10px 14px;background:#f9f9fb;border-radius:10px}
.filter-bar{display:flex;gap:12px;margin-bottom:16px;flex-wrap:wrap;align-items:center}
.filter-bar input{flex:1;min-width:200px;padding:9px 16px;border:1px solid var(--border);border-radius:24px;font-size:14px;outline:none;font-family:inherit}
footer{text-align:center;color:var(--sub);font-size:12px;padding:28px}

/* ════════════════════════════════════════════════════════
   RWD：iPad / iPhone 響應式優化
   ════════════════════════════════════════════════════════ */
/* iPad portrait 以下 (≤ 900px)：兩欄變一欄 */
@media (max-width: 900px) {
  .container { padding: 24px 18px; }
  .header { padding: 26px 22px 22px; }
  .header h1 { font-size: 24px; }
}
/* iPhone Pro Max / iPad mini 以下 (≤ 768px) */
@media (max-width: 768px) {
  .container { padding: 18px 12px; }
  .header { padding: 22px 16px 18px; }
  .header h1 { font-size: 20px; line-height: 1.25; letter-spacing: 0; }
  .header p { font-size: 12px; }
  .header-meta { gap: 6px; }
  .meta-badge { font-size: 11px; padding: 4px 10px; }

  .section-title { font-size: 17px; margin: 24px 0 12px; }
  .card { padding: 14px 12px; border-radius: 12px; margin-bottom: 14px; }
  .card-title { font-size: 14px; }

  /* 摘要 KPI：iPhone 至少 2 欄不單欄 */
  .summary-grid { grid-template-columns: repeat(2, 1fr); gap: 10px; margin-bottom: 18px; }
  .summary-card { padding: 14px 8px; border-radius: 12px; }
  .summary-card .value { font-size: 24px; }
  .summary-card .label { font-size: 10px; }

  /* 門市卡片 */
  .store-card { border-radius: 12px; margin-bottom: 12px; }
  .store-header { padding: 11px 14px; gap: 6px; }
  .store-title { font-size: 14px; width: 100%; }
  .store-kpis { width: 100%; gap: 5px; }
  .kpi-chip { font-size: 11px; padding: 3px 10px; }

  /* 銷量輸入列：直立排列 */
  .sales-row { padding: 12px 14px; gap: 8px; }
  .sales-divider { display: none; }
  .sales-input { width: 100%; max-width: 200px; }

  /* 操作人列 */
  .op-summary-bar { padding: 8px 14px; }
  .op-pill { font-size: 11px; padding: 2px 9px; }

  /* 表格：橫向捲動，避免被遮蓋 */
  .store-card > table.data-table,
  .bonus-month-card > table.data-table,
  .card > table.data-table,
  details > table.data-table {
    display: block;
    overflow-x: auto;
    -webkit-overflow-scrolling: touch;
    white-space: nowrap;
  }
  .data-table th, .data-table td { padding: 8px 10px; }
  .data-table .small-text, .data-table .op-cell { white-space: normal; max-width: none; }

  /* 獎金計算機 */
  .bonus-summary { grid-template-columns: repeat(2, 1fr); gap: 8px; }
  .bsum-card { padding: 10px 8px; }
  .bsum-val { font-size: 15px; }
  .bonus-month-card { padding: 12px; border-radius: 12px; }
  .bonus-month-head { gap: 8px; padding-bottom: 10px; margin-bottom: 12px; }
  .bonus-month-head > span:last-child { margin-left: 0 !important; width: 100%; }
  .bonus-calc-grid { grid-template-columns: 1fr; gap: 10px; }
  .bcalc { padding: 12px; }

  /* 排行榜雙指標表：太寬時提示左右滑動 */
  #leaderboard-table th, #leaderboard-table td { font-size: 11px; padding: 6px 8px; }

  /* Filter / 搜尋列 */
  .filter-bar input[type=text] { width: 100%; min-width: 0; font-size: 13px; }

  /* footer */
  footer { padding: 20px 14px; font-size: 11px; }
}
/* iPhone SE / 小螢幕 (≤ 380px) */
@media (max-width: 380px) {
  .header h1 { font-size: 18px; }
  .summary-grid { grid-template-columns: 1fr; }
  .store-title { font-size: 13px; }
  .meta-badge { font-size: 10px; padding: 3px 8px; }
}
</style>
</head>
<body>
<div class="header">
  <h1>📱 {{dealer_kw}} Trade-in 舊換新分析報告</h1>
  <p>Apple FY 週次追蹤 · 自動產生</p>
  <div class="header-meta">
    <span class="meta-badge">📅 {{first_week}}–{{last_week}}</span>
    <span class="meta-badge">🏪 {{n_stores}} 門市</span>
    <span class="meta-badge">📊 {{total_exec}} 筆執行</span>
    <span class="meta-badge">👤 {{n_ops}} 位操作人</span>
  </div>
</div>
<div class="container">
  <div class="summary-grid">
    <div class="summary-card"><div class="label">總執行數</div><div class="value blue">{{total_exec}}</div><div class="sub">全期合計</div></div>
    <div class="summary-card"><div class="label">Plugin 接機🔌</div><div class="value purple">{{total_plugin}}</div><div class="sub">有 IMEI 接線記錄</div></div>
    <div class="summary-card"><div class="label">成交數</div><div class="value green">{{total_deal}}</div><div class="sub">完成以舊換新</div></div>
    <div class="summary-card"><div class="label">操作人數</div><div class="value teal">{{n_ops}}</div><div class="sub">有成交記錄的人員</div></div>
  </div>
  <div class="section-title">📊 全{{dealer_kw}}週次總覽</div>
  <div class="card"><table class="data-table">
    <thead><tr><th>週次</th><th>執行數</th><th>Plugin🔌</th><th>成交數</th></tr></thead>
    <tbody>{{weekly_overview_rows}}<tr style="background:#f5f5f7;font-weight:700"><td>合計</td><td class="num blue">{{total_exec}}</td><td class="num purple">{{total_plugin}}</td><td class="num green">{{total_deal}}</td></tr></tbody>
  </table></div>
  <div class="section-title">👤 全{{dealer_kw}}操作人成交排行</div>
  <div class="two-col">
    <div class="card"><div class="card-title">🏆 操作人成交總排行</div><table class="data-table"><thead><tr><th>#</th><th>操作人</th><th>主要門市</th><th>成交筆數</th></tr></thead><tbody>{{global_op_rows}}</tbody></table></div>
    <div style="display:flex;flex-direction:column;gap:20px">
      <div class="card" style="margin:0"><div class="card-title">📤 最常被換出舊機 Top 10</div><table class="data-table"><thead><tr><th>舊機型號</th><th>次數</th></tr></thead><tbody>{{old_rows}}</tbody></table></div>
      <div class="card" style="margin:0"><div class="card-title">📥 最常換購新品 Top 10</div><table class="data-table"><thead><tr><th>類別</th><th>新品型號</th><th>次數</th></tr></thead><tbody>{{new_rows}}</tbody></table></div>
    </div>
  </div>
  <div class="section-title">💰 個人舊換新獎金計算機</div>
  <div class="card" id="bonus-locked-card">
    <div style="text-align:center;padding:24px 20px">
      <div style="font-size:48px;margin-bottom:10px">🔒</div>
      <div style="font-size:16px;font-weight:700;margin-bottom:6px">此區僅限授權人員查看</div>
      <div style="font-size:13px;color:var(--sub);margin-bottom:18px">請輸入密碼以解鎖個人獎金資料</div>
      <div style="display:flex;gap:8px;justify-content:center;flex-wrap:wrap">
        <input type="password" id="bonusPwd" placeholder="密碼" onkeydown="if(event.key==='Enter')tryUnlock()" style="width:200px;padding:9px 16px;border:1.5px solid var(--border);border-radius:24px;font-size:14px;outline:none;font-family:inherit;text-align:center">
        <button onclick="tryUnlock()" style="padding:9px 22px;border:none;border-radius:24px;background:var(--accent);color:#fff;font-size:14px;font-weight:600;cursor:pointer">🔓 解鎖</button>
      </div>
      <div id="pwd-error" style="color:#d63030;font-size:12px;margin-top:10px;display:none">❌ 密碼錯誤，請重試</div>
    </div>
  </div>
  <div class="card" id="bonus-unlocked-card" style="display:none">
    <div class="bonus-rules-bar">
      <div class="bonus-rules-text">
        <div><strong>🟢 基本獎勵</strong>（月總回收 &lt; 20 萬）<br>
          高回收(≥$10,001) → 總額 × <strong>1%</strong> ・ 低回收(≤$10,000) → 每件 <strong>$100</strong></div>
        <div><strong>🟡 進階獎勵</strong>（月總回收 ≥ 20 萬）<br>
          高回收(≥$10,001) → 總額 × <strong>2%</strong> ・ 低回收(≤$10,000) → 每件 <strong>$200</strong></div>
      </div>
      <button onclick="lockBonus()" class="lock-btn">🔒 鎖定</button>
    </div>
    <div class="filter-bar" style="margin-bottom:18px">
      <select id="bonusOpSelect" onchange="renderBonus(this.value)" style="flex:1;min-width:240px;padding:9px 16px;border:1px solid var(--border);border-radius:24px;font-size:14px;font-family:inherit;background:#fff">
        <option value="">— 請選擇操作人（薪號／姓名）—</option>
      </select>
    </div>
    <div id="bonus-result"><div class="empty-hint">👈 從上方選一位操作人來查詢個人獎金</div></div>
  </div>
  <div class="section-title">📈 iPhone 回收率 / 接線率排行榜</div>
  <div class="card">
    <p style="font-size:13px;color:var(--sub);margin-bottom:10px">在下方各門市卡片輸入 iPhone 銷量後，此排行榜會自動更新並依回收率排序。</p>
    <div class="rate-rules-bar">
      <span><strong style="color:var(--green)">回收率</strong> ＝ 成交數 ÷ iPhone 銷量 × 100% ・ 達標 <strong>≥ {{recycle_target}}%</strong></span>
      <span><strong style="color:var(--purple)">接線率</strong> ＝ 接機數 ÷ iPhone 銷量 × 100% ・ 達標 <strong>≥ {{plugin_target}}%</strong></span>
    </div>
    <div id="rate-leaderboard-wrap">
      <div class="empty-hint" id="leaderboard-empty">尚無資料 — 請先在門市卡片輸入 iPhone 銷量 👇</div>
      <table class="data-table" id="leaderboard-table" style="display:none">
        <thead><tr><th>#</th><th>門市</th><th>iPhone 銷量</th><th>接機數</th><th>接線率</th><th>成交數</th><th>回收率</th></tr></thead>
        <tbody id="leaderboard-body"></tbody>
      </table>
    </div>
  </div>
  <div class="section-title">🏪 各門市週次明細</div>
  <div class="filter-bar">
    <input type="text" id="storeSearch" placeholder="🔍 搜尋門市名稱..." oninput="filterStores(this.value)">
    <span style="font-size:13px;color:var(--sub)">共 {{n_stores}} 間門市</span>
  </div>
  <div id="stores-container">{{store_blocks}}</div>
</div>
<footer>{{dealer_kw}} Trade-in 週報 · 由 analyze_and_report.py 自動產生<br>Plugin 判斷：检测信息含 IMEI 碼 · 成交判斷：回收類型＝以旧换新</footer>
<script>
const STORES={{stores_json}};
const BONUS_RECORDS={{bonus_records_json}};
const LS_KEY='ck_iphone_sales';
const LATEST_WEEK='{{latest_week}}';
const RECYCLE_TARGET={{recycle_target}};   // 回收率達標 (%)
const PLUGIN_TARGET ={{plugin_target}};    // 接線率達標 (%)
const HIGH_THRESHOLD=10001;                // ≥ 10,001 算高回收
const BASIC_HIGH_RATE=0.01, ADV_HIGH_RATE=0.02;
const BASIC_LOW_PER=100, ADV_LOW_PER=200;
const ADVANCE_THRESHOLD=200000;            // 月回收 ≥ 20 萬 → 進階
function loadSales(){try{return JSON.parse(localStorage.getItem(LS_KEY))||{};}catch{return{};}}
function saveSales(s){localStorage.setItem(LS_KEY,JSON.stringify(s));}
function rateClass(pct,target){if(pct>=target)return 'high';if(pct>=target/2)return 'mid';return 'low';}
function calcRate(input){
  const idx=parseInt(input.dataset.idx);
  const deal=parseInt(input.dataset.deal);
  const plugin=parseInt(input.dataset.plugin);
  const sales=parseInt(input.value)||0;
  const rb=document.getElementById('rate-badge-'+idx);
  const pb=document.getElementById('plug-badge-'+idx);
  const all=loadSales();
  if(sales>0)all[idx]=sales;else delete all[idx];
  saveSales(all);
  if(sales>0){
    const rPct=deal/sales*100;
    rb.textContent=rPct.toFixed(1)+'%';
    rb.className='rate-badge '+rateClass(rPct,RECYCLE_TARGET);
    const pPct=plugin/sales*100;
    pb.textContent=pPct.toFixed(1)+'%';
    pb.className='rate-badge '+rateClass(pPct,PLUGIN_TARGET);
  }else{
    rb.textContent='—';rb.className='rate-badge';
    pb.textContent='—';pb.className='rate-badge';
  }
  updateLeaderboard();
}
function updateLeaderboard(){
  const all=loadSales(),rows=[];
  Object.entries(all).forEach(([idx,sales])=>{
    const s=STORES[idx];if(!s)return;
    rows.push({idx,name:s.name,deal:s.deal,plugin:s.plugin,sales,
               rPct:s.deal/sales*100, pPct:s.plugin/sales*100});
  });
  const empty=document.getElementById('leaderboard-empty'),table=document.getElementById('leaderboard-table'),tbody=document.getElementById('leaderboard-body');
  if(rows.length===0){empty.style.display='';table.style.display='none';return;}
  empty.style.display='none';table.style.display='';
  rows.sort((a,b)=>b.rPct-a.rPct);  // 依回收率排序
  const colors={high:'#25a244',mid:'#c97200',low:'#d63030'};
  tbody.innerHTML=rows.map((r,i)=>{
    const rc=i===0?'r1':i===1?'r2':i===2?'r3':'';
    const rCl=rateClass(r.rPct,RECYCLE_TARGET), pCl=rateClass(r.pPct,PLUGIN_TARGET);
    const rIcon = r.rPct>=RECYCLE_TARGET ? ' ✅' : '';
    const pIcon = r.pPct>=PLUGIN_TARGET  ? ' ✅' : '';
    return `<tr>
      <td><span class="rate-rank-badge ${rc}">${i+1}</span></td>
      <td><strong>${r.name}</strong></td>
      <td class="num">${r.sales.toLocaleString()}</td>
      <td class="num plugin-cell">${r.plugin}</td>
      <td class="num" style="color:${colors[pCl]};font-weight:800">${r.pPct.toFixed(1)}%${pIcon}</td>
      <td class="num deal-yes">${r.deal}</td>
      <td class="num" style="color:${colors[rCl]};font-size:15px;font-weight:800">${r.rPct.toFixed(1)}%${rIcon}</td>
    </tr>`;
  }).join('');
}
function filterStores(q){q=q.toLowerCase();document.querySelectorAll('.store-card').forEach(c=>{c.style.display=c.querySelector('.store-title').textContent.toLowerCase().includes(q)?'':'none';});}

// ── 獎金計算 ───────────────────────────────────────────
function calcMonthBonus(records){
  // records: [{date,month,week,store,price}]
  const high=records.filter(r=>r.price>=HIGH_THRESHOLD);
  const low =records.filter(r=>r.price<HIGH_THRESHOLD);
  const highSum=high.reduce((s,r)=>s+r.price,0);
  const lowSum =low.reduce((s,r)=>s+r.price,0);
  const total=highSum+lowSum;
  const advanced=total>=ADVANCE_THRESHOLD;
  const highBonus=Math.round(highSum*(advanced?ADV_HIGH_RATE:BASIC_HIGH_RATE));
  const lowBonus=low.length*(advanced?ADV_LOW_PER:BASIC_LOW_PER);
  return {high,low,highSum,lowSum,total,advanced,highBonus,lowBonus,bonus:highBonus+lowBonus};
}
function fmtMoney(n){return '$'+n.toLocaleString();}
function renderBonus(op){
  const target=document.getElementById('bonus-result');
  if(!op){target.innerHTML='<div class="empty-hint">👈 從上方選一位操作人來查詢個人獎金</div>';return;}
  const recs=BONUS_RECORDS[op]||[];
  if(recs.length===0){target.innerHTML='<div class="empty-hint">該操作人無成交記錄</div>';return;}
  // group by month
  const monthMap={};
  recs.forEach(r=>{(monthMap[r.month]=monthMap[r.month]||[]).push(r);});
  const months=Object.keys(monthMap).sort();
  const totalBonus=months.reduce((s,m)=>s+calcMonthBonus(monthMap[m]).bonus,0);
  const totalDeals=recs.length;
  const totalAmt=recs.reduce((s,r)=>s+r.price,0);
  const stores=Array.from(new Set(recs.map(r=>r.store))).join('、');

  let html=`<div class="bonus-summary">
    <div class="bsum-card"><div class="bsum-label">操作人</div><div class="bsum-val">${op}</div></div>
    <div class="bsum-card"><div class="bsum-label">服務門市</div><div class="bsum-val" style="font-size:13px">${stores}</div></div>
    <div class="bsum-card"><div class="bsum-label">總成交件數</div><div class="bsum-val">${totalDeals} 件</div></div>
    <div class="bsum-card"><div class="bsum-label">總回收金額</div><div class="bsum-val">${fmtMoney(totalAmt)}</div></div>
    <div class="bsum-card highlight"><div class="bsum-label">總獎金（全期）</div><div class="bsum-val" style="color:#c97200">${fmtMoney(totalBonus)}</div></div>
  </div>`;

  // 月份卡片
  months.forEach(m=>{
    const b=calcMonthBonus(monthMap[m]);
    const ruleBadge=b.advanced
      ? '<span class="rule-badge adv">🟡 進階獎勵 (≥20萬)</span>'
      : '<span class="rule-badge basic">🟢 基本獎勵</span>';
    // 週次明細
    const weekMap={};
    monthMap[m].forEach(r=>{(weekMap[r.week]=weekMap[r.week]||[]).push(r);});
    const weekKeys=Object.keys(weekMap).sort();
    let weekRows='';
    weekKeys.forEach(wk=>{
      const wRecs=weekMap[wk];
      const wHigh=wRecs.filter(r=>r.price>=HIGH_THRESHOLD);
      const wLow =wRecs.filter(r=>r.price<HIGH_THRESHOLD);
      const wHighSum=wHigh.reduce((s,r)=>s+r.price,0);
      const wLowSum =wLow.reduce((s,r)=>s+r.price,0);
      weekRows+=`<tr>
        <td><strong>${wk}</strong></td>
        <td class="num">${wHigh.length}</td><td class="num">${fmtMoney(wHighSum)}</td>
        <td class="num">${wLow.length}</td><td class="num">${fmtMoney(wLowSum)}</td>
        <td class="num"><strong>${fmtMoney(wHighSum+wLowSum)}</strong></td>
      </tr>`;
    });
    // 全月詳細記錄
    const detailRows=monthMap[m].slice().sort((a,b)=>a.date.localeCompare(b.date)).map(r=>{
      const tier=r.price>=HIGH_THRESHOLD?'<span style="color:#0071e3;font-size:11px">高</span>':'<span style="color:var(--sub);font-size:11px">低</span>';
      return `<tr><td>${r.date}</td><td>${r.week}</td><td>${r.store.replace('燦坤','').replace('TK3C@009','')}</td><td class="num">${fmtMoney(r.price)}</td><td>${tier}</td></tr>`;
    }).join('');

    html+=`<div class="bonus-month-card">
      <div class="bonus-month-head">
        <span style="font-size:18px;font-weight:700">📅 ${m}</span>
        ${ruleBadge}
        <span style="margin-left:auto;font-size:13px;color:var(--sub)">月總回收 <strong style="color:var(--text)">${fmtMoney(b.total)}</strong></span>
      </div>
      <div class="bonus-calc-grid">
        <div class="bcalc">
          <div class="bcalc-title" style="color:var(--accent)">🟦 高回收（≥${fmtMoney(HIGH_THRESHOLD)}）</div>
          <div class="bcalc-row"><span>件數</span><strong>${b.high.length} 件</strong></div>
          <div class="bcalc-row"><span>金額合計</span><strong>${fmtMoney(b.highSum)}</strong></div>
          <div class="bcalc-row"><span>計算公式</span><span style="font-size:11px">${fmtMoney(b.highSum)} × ${(b.advanced?ADV_HIGH_RATE:BASIC_HIGH_RATE)*100}%</span></div>
          <div class="bcalc-row total"><span>小計</span><strong style="color:var(--green)">${fmtMoney(b.highBonus)}</strong></div>
        </div>
        <div class="bcalc">
          <div class="bcalc-title" style="color:var(--sub)">⬜️ 低回收（&lt;${fmtMoney(HIGH_THRESHOLD)}）</div>
          <div class="bcalc-row"><span>件數</span><strong>${b.low.length} 件</strong></div>
          <div class="bcalc-row"><span>金額合計</span><strong>${fmtMoney(b.lowSum)}</strong></div>
          <div class="bcalc-row"><span>計算公式</span><span style="font-size:11px">${b.low.length} × ${fmtMoney(b.advanced?ADV_LOW_PER:BASIC_LOW_PER)}</span></div>
          <div class="bcalc-row total"><span>小計</span><strong style="color:var(--green)">${fmtMoney(b.lowBonus)}</strong></div>
        </div>
        <div class="bcalc highlight">
          <div class="bcalc-title" style="color:#c97200">💰 該月總獎金</div>
          <div class="bcalc-row"><span>高回收獎金</span><strong>${fmtMoney(b.highBonus)}</strong></div>
          <div class="bcalc-row"><span>低回收獎金</span><strong>${fmtMoney(b.lowBonus)}</strong></div>
          <div class="bcalc-row total" style="margin-top:auto"><span>合計</span><strong style="color:#c97200;font-size:22px">${fmtMoney(b.bonus)}</strong></div>
        </div>
      </div>
      <details style="margin-top:14px">
        <summary style="cursor:pointer;font-size:13px;color:var(--accent);font-weight:600">📊 各週明細（${weekKeys.length} 週）</summary>
        <table class="data-table" style="margin-top:8px">
          <thead><tr><th>週次</th><th>高件數</th><th>高金額</th><th>低件數</th><th>低金額</th><th>週小計</th></tr></thead>
          <tbody>${weekRows}</tbody>
        </table>
      </details>
      <details style="margin-top:8px">
        <summary style="cursor:pointer;font-size:13px;color:var(--accent);font-weight:600">📋 全月成交明細（${monthMap[m].length} 筆）</summary>
        <table class="data-table" style="margin-top:8px">
          <thead><tr><th>日期</th><th>週次</th><th>門市</th><th>成交價</th><th>級距</th></tr></thead>
          <tbody>${detailRows}</tbody>
        </table>
      </details>
    </div>`;
  });
  target.innerHTML=html;
}
function initBonusSelect(){
  const sel=document.getElementById('bonusOpSelect');
  if(!sel) return;
  if(sel.options.length>1) return; // 已初始化
  // 排序：依總成交件數降冪
  const ops=Object.keys(BONUS_RECORDS).map(op=>{
    const recs=BONUS_RECORDS[op];
    return {op,count:recs.length,total:recs.reduce((s,r)=>s+r.price,0)};
  }).sort((a,b)=>b.count-a.count);
  ops.forEach(({op,count,total})=>{
    const opt=document.createElement('option');
    opt.value=op;opt.textContent=`${op} (${count} 件 / ${'$'+total.toLocaleString()})`;
    sel.appendChild(opt);
  });
}

// ── 密碼鎖（個人獎金區） ──────────────────────────────
const BONUS_PWD_HASH='46a3bc0b6fdbcda91cb959d786e25cd2db5c38a95aaeac5fb4caad4ba4dab92d';  // SHA-256
const UNLOCK_KEY='ck_bonus_unlocked_v1';
async function sha256(text){
  const buf=new TextEncoder().encode(text);
  const hash=await crypto.subtle.digest('SHA-256',buf);
  return Array.from(new Uint8Array(hash)).map(b=>b.toString(16).padStart(2,'0')).join('');
}
async function tryUnlock(){
  const input=document.getElementById('bonusPwd');
  const err=document.getElementById('pwd-error');
  const h=await sha256(input.value);
  if(h===BONUS_PWD_HASH){
    err.style.display='none';
    sessionStorage.setItem(UNLOCK_KEY,'1');
    showUnlocked();
  }else{
    err.style.display='block';
    input.value='';
    input.focus();
  }
}
function showUnlocked(){
  document.getElementById('bonus-locked-card').style.display='none';
  document.getElementById('bonus-unlocked-card').style.display='';
  initBonusSelect();
}
function lockBonus(){
  sessionStorage.removeItem(UNLOCK_KEY);
  document.getElementById('bonus-locked-card').style.display='';
  document.getElementById('bonus-unlocked-card').style.display='none';
  const pwd=document.getElementById('bonusPwd');
  if(pwd) pwd.value='';
  // 清空已查詢結果，避免被截圖殘留
  const sel=document.getElementById('bonusOpSelect');
  if(sel) sel.value='';
  const res=document.getElementById('bonus-result');
  if(res) res.innerHTML='<div class="empty-hint">👈 從上方選一位操作人來查詢個人獎金</div>';
}
function checkUnlockOnLoad(){
  if(sessionStorage.getItem(UNLOCK_KEY)==='1') showUnlocked();
}
window.addEventListener('load',()=>{const all=loadSales();Object.entries(all).forEach(([idx,sales])=>{const input=document.getElementById('sales-'+idx);if(input){input.value=sales;calcRate(input);}});updateLeaderboard();checkUnlockOnLoad();});
</script>
</body>
</html>"""


def make_op_pills(op_list):
    if not op_list:
        return "<span style='color:#ccc;font-size:12px'>—</span>"
    return "".join(
        f"<span class='op-pill'>{op} <span class='op-cnt'>{cnt}</span></span>"
        for op, cnt in op_list
    )


def render_html(data: dict, dealer_kw: str) -> str:
    results       = data["results"]
    stores        = data["stores"]
    weeks         = data["weeks"]
    week_dates    = data["week_dates"]
    all_old       = data["all_old"]
    all_new       = data["all_new"]
    weekly_totals = data["weekly_totals"]
    global_op     = data["global_op"]
    latest_week   = data["latest_week"] or ""

    grand = {"exec": 0, "plugin": 0, "deal": 0}
    for s in stores:
        for w in weeks:
            d = results[s][w]
            grand["exec"]   += d["exec"]
            grand["plugin"] += d["plugin"]
            grand["deal"]   += d["deal"]

    # 操作人主要門市
    op_store = {}
    for store in stores:
        for w in weeks:
            for op, cnt in results[store][w]["op_deal"]:
                op_store.setdefault(op, {})
                op_store[op][store] = op_store[op].get(store, 0) + cnt

    medals = ["🥇", "🥈", "🥉"]
    global_op_rows = ""
    for rank, (op, total) in enumerate(global_op):
        medal = medals[rank] if rank < 3 else f"{rank+1}."
        main_store = max(op_store.get(op, {"—": 0}),
                         key=lambda k: op_store.get(op, {}).get(k, 0))
        short = main_store.replace("燦坤", "").replace("TK3C@009", "") or main_store
        global_op_rows += (
            f"<tr><td>{medal}</td><td><strong>{op}</strong></td>"
            f"<td>{short}</td><td class='num deal-yes'>{total}</td></tr>"
        )

    # 週次總覽
    weekly_overview_rows = ""
    for w in weeks:
        wt = weekly_totals[w]
        is_new = w == latest_week
        new_cls = " class='new-week'" if is_new else ""
        new_badge = '<span class="new-badge">NEW</span>' if is_new else ""
        weekly_overview_rows += (
            f"<tr{new_cls}><td><strong>{w}</strong>{new_badge}<br><small>{week_dates[w]}</small></td>"
            f"<td class='num'>{wt['exec']}</td>"
            f"<td class='num plugin-cell'>🔌 {wt['plugin']}</td>"
            f"<td class='num deal-yes'>{wt['deal']}</td></tr>"
        )

    old_rows = "".join(
        f"<tr><td>{m}</td><td class='num'>{c}</td></tr>" for m, c in all_old[:10]
    )
    new_rows = "".join(
        f"<tr><td>{lob}</td><td>{m}</td><td class='num'>{c}</td></tr>"
        for lob, m, c in all_new[:10]
    )

    stores_js = {
        stores.index(s): {
            "name":   s,
            "deal":   sum(results[s][w]["deal"]   for w in weeks),
            "plugin": sum(results[s][w]["plugin"] for w in weeks),
        }
        for s in stores
    }

    # 各門市卡片
    store_blocks = ""
    for store in stores:
        idx = stores.index(store)
        te = sum(results[store][w]["exec"]   for w in weeks)
        td = sum(results[store][w]["deal"]   for w in weeks)
        tp = sum(results[store][w]["plugin"] for w in weeks)
        if te == 0:
            continue
        op_agg = {}
        for w in weeks:
            for op, cnt in results[store][w]["op_deal"]:
                op_agg[op] = op_agg.get(op, 0) + cnt
        all_period_ops = make_op_pills(sorted(op_agg.items(), key=lambda x: -x[1]))

        week_rows = ""
        for w in weeks:
            d = results[store][w]
            old_str = "、".join(f"{m}×{c}" for m, c in d["old_top"][:3])
            new_str = "、".join(f"{m}×{c}" for lob, m, c in d["new_items"][:3]) or "—"
            op_str  = make_op_pills(d["op_deal"])
            deal_cls = "deal-yes" if d["deal"] > 0 else ""
            zero_cls = "zero-row" if d["exec"] == 0 else ""
            new_row_cls = " new-week-row" if w == latest_week else ""
            new_badge = '<span class="new-badge">NEW</span>' if w == latest_week else ""
            plugin_icon = "🔌 " if d["plugin"] > 0 else ""
            week_rows += (
                f"<tr class='{zero_cls}{new_row_cls}'>"
                f"<td>{w}{new_badge}<br><small>{week_dates[w]}</small></td>"
                f"<td class='num'>{d['exec']}</td>"
                f"<td class='num plugin-cell'>{plugin_icon}{d['plugin']}</td>"
                f"<td class='num {deal_cls}'>{d['deal']}</td>"
                f"<td class='op-cell'>{op_str}</td>"
                f"<td class='small-text'>{old_str or '—'}</td>"
                f"<td class='small-text'>{new_str}</td>"
                f"</tr>"
            )

        store_blocks += f"""<div class="store-card" id="store-{idx}">
<div class="store-header">
  <div class="store-title">🏪 {store}</div>
  <div class="store-kpis">
    <div class="kpi-chip exec-chip">執行 {te}</div>
    <div class="kpi-chip plugin-chip">Plugin 🔌 {tp}</div>
    <div class="kpi-chip deal-chip">成交 {td}</div>
  </div>
</div>
<div class="sales-row">
  <span class="sales-label">📱 iPhone 銷量</span>
  <input type="number" class="sales-input" id="sales-{idx}" data-idx="{idx}" data-deal="{td}" data-plugin="{tp}" placeholder="輸入銷量…" min="0" oninput="calcRate(this)">
  <span class="sales-divider">|</span>
  <span class="sales-label">回收率</span>
  <span class="rate-badge" id="rate-badge-{idx}" title="目標 ≥ {RECYCLE_TARGET:.0f}%">—</span>
  <span class="sales-divider">|</span>
  <span class="sales-label">接線率</span>
  <span class="rate-badge" id="plug-badge-{idx}" title="目標 ≥ {PLUGIN_TARGET:.0f}%">—</span>
</div>
<div class="op-summary-bar"><span class="op-label">👤 操作人（全期成交）：</span>{all_period_ops}</div>
<table class="data-table">
  <thead><tr><th>週次</th><th>執行數</th><th>Plugin🔌</th><th>成交數</th><th>操作人（成交筆數）</th><th>主要舊機（前3）</th><th>換購新品（前3）</th></tr></thead>
  <tbody>{week_rows}</tbody>
</table>
</div>"""

    html = HTML_TEMPLATE
    replacements = {
        "{{title}}":          f"{dealer_kw} Trade-in 週報",
        "{{dealer_kw}}":      dealer_kw,
        "{{first_week}}":     weeks[0] if weeks else "",
        "{{last_week}}":      weeks[-1] if weeks else "",
        "{{n_stores}}":       str(len([s for s in stores if any(results[s][w]['exec']>0 for w in weeks)])),
        "{{total_exec}}":     str(grand["exec"]),
        "{{total_plugin}}":   str(grand["plugin"]),
        "{{total_deal}}":     str(grand["deal"]),
        "{{n_ops}}":          str(len(global_op)),
        "{{weekly_overview_rows}}": weekly_overview_rows,
        "{{global_op_rows}}": global_op_rows,
        "{{old_rows}}":       old_rows,
        "{{new_rows}}":       new_rows,
        "{{store_blocks}}":   store_blocks,
        "{{stores_json}}":    json.dumps(stores_js, ensure_ascii=False),
        "{{bonus_records_json}}": json.dumps(data.get("bonus_records", {}), ensure_ascii=False),
        "{{latest_week}}":    latest_week,
        "{{recycle_target}}": str(RECYCLE_TARGET),
        "{{plugin_target}}":  str(PLUGIN_TARGET),
    }
    for k, v in replacements.items():
        html = html.replace(k, v)
    return html


# ────────────────────────────────────────────────────────────────────
# Main
# ────────────────────────────────────────────────────────────────────
def main():
    p = argparse.ArgumentParser(description="燦坤 Trade-in 週報分析 + HTML 報告產生器")
    p.add_argument("--input", "-i", nargs="+", required=True,
                   help="一個或多個 Excel 檔路徑（會自動合併）")
    p.add_argument("--output", "-o", default="./燦坤週報.html",
                   help="輸出 HTML 路徑")
    p.add_argument("--dealer", "-d", default=DEALER_DEFAULT,
                   help=f"經銷商名稱關鍵字（預設：{DEALER_DEFAULT}）")
    p.add_argument("--start-date", "-s", default=WEEK_START_DEFAULT,
                   help=f"FY 起始日 YYYY-MM-DD（預設：{WEEK_START_DEFAULT}）")
    p.add_argument("--num-weeks", "-w", type=int, default=0,
                   help="週數（0 = 自動推算到最新資料日）")
    p.add_argument("--export-json", action="store_true",
                   help="同時匯出處理後的 JSON 供下次增量使用")
    args = p.parse_args()

    for f in args.input:
        if not Path(f).exists():
            sys.exit(f"❌ 找不到檔案：{f}")

    print(f"\n🚀 開始分析（經銷商 = '{args.dealer}'）\n")
    data = analyze(args.input, args.dealer, args.start_date, args.num_weeks)

    print(f"\n🎨 產生 HTML 報告 → {args.output}")
    html = render_html(data, args.dealer)
    Path(args.output).write_text(html, encoding="utf-8")
    print(f"✅ 完成！檔案大小：{len(html):,} bytes")

    if args.export_json:
        json_out = Path(args.output).with_suffix(".json")
        json_out.write_text(json.dumps(data, ensure_ascii=False, indent=2),
                            encoding="utf-8")
        print(f"💾 JSON 已匯出 → {json_out}")

    print(f"\n📊 摘要：")
    print(f"   週次：{len(data['weeks'])} 週（{data['weeks'][0]} – {data['weeks'][-1]}）")
    print(f"   門市：{len(data['stores'])} 間")
    print(f"   總執行：{sum(data['weekly_totals'][w]['exec']   for w in data['weeks']):,}")
    print(f"   總接機：{sum(data['weekly_totals'][w]['plugin'] for w in data['weeks']):,}")
    print(f"   總成交：{sum(data['weekly_totals'][w]['deal']   for w in data['weeks']):,}")
    print(f"   操作人：{len(data['global_op'])} 位\n")


if __name__ == "__main__":
    main()
