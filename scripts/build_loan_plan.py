#!/usr/bin/env python3
"""決算書3期分（financials.json）と案件設定（case.json）から、融資用の事業計画パッケージを生成する。

出力（同じフォルダ）:
  control-sheet.json / <slug>-business-plan.md / <slug>-market-pricing.md / <slug>-presentation.pptx / .pdf
  <slug>-bank-appendix.md / .pdf（資金使途・借入余力・返済計画・資金繰り表・3期財務分析）
  <slug>-contract-checklist.md（契約を伴う場合）
数値の計算はこのファイルだけで行い、描画と検証は build-business-plan の plan_gate / render_package を使う。
"""
from __future__ import annotations

import argparse
import copy
import datetime as dt
import hashlib
import json
import math
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any

SKILL_DIR = Path(__file__).resolve().parent.parent
BBP_SCRIPTS = Path(os.environ.get("BUSINESS_PLAN_SCRIPTS", SKILL_DIR.parent / "build-business-plan" / "scripts"))
sys.path.insert(0, str(BBP_SCRIPTS))
import plan_gate  # noqa: E402

JST = dt.timezone(dt.timedelta(hours=9))
FIELDS = (
    "sales", "cost_of_sales", "sga", "depreciation", "interest_expense", "ordinary_income", "net_income",
    "cash", "receivables", "inventories", "payables", "interest_bearing_debt", "officer_loans",
    "annual_principal_repayment", "net_assets", "total_assets",
)
FIELD_LABELS = {
    "sales": "売上高", "cost_of_sales": "売上原価（完成工事原価）", "sga": "販売費及び一般管理費",
    "depreciation": "減価償却費", "interest_expense": "支払利息", "ordinary_income": "経常利益",
    "net_income": "当期純利益（税引後）", "cash": "現預金", "receivables": "売上債権",
    "inventories": "棚卸資産（未成工事支出金等）", "payables": "仕入債務", "interest_bearing_debt": "有利子負債（役員借入金を除く）",
    "officer_loans": "役員借入金", "annual_principal_repayment": "年間元金返済額", "net_assets": "純資産", "total_assets": "総資産",
}


# ---------- 書式 ----------

def man(value: float) -> str:
    """円を「◯万円」表記にする（重要数値の正規値として全成果物で同じ文字列を使う）。"""
    v = value / 10000
    return f"{v:,.0f}万円" if abs(v - round(v)) < 1e-9 else f"{v:,.1f}万円"


def yen(value: float) -> str:
    return f"{value:,.0f}"


def pct(value: float) -> str:
    return f"{value * 100:.1f}%"


def years_text(value: float) -> str:
    if math.isinf(value):
        return "算出不能（簡易CFが0以下）"
    return "0年（現預金が有利子負債を上回る）" if value <= 0 else f"{value:.1f}年"


def floor_unit(value: float, unit: int) -> int:
    return max(0, int(value // unit) * unit)


def ceil_unit(value: float, unit: int) -> int:
    return int(math.ceil(value / unit) * unit)


# ---------- 入力検証 ----------

def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def validate_financials(fin: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    if not str(fin.get("company", "")).strip():
        errors.append("company（会社名）が空です")
    periods = fin.get("periods")
    if not isinstance(periods, list) or len(periods) != 3:
        return errors + ["periodsは決算3期分（3件）が必要です"]
    ends = []
    for i, p in enumerate(periods, 1):
        tag = f"periods[{i}]"
        try:
            ends.append(dt.date.fromisoformat(str(p.get("period_end"))))
        except ValueError:
            errors.append(f"{tag}.period_endはYYYY-MM-DDで必要です")
        if not str(p.get("label", "")).strip():
            errors.append(f"{tag}.labelが空です")
        bad = False
        for key in FIELDS:
            value = p.get(key)
            if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
                errors.append(f"{tag}.{key}（{FIELD_LABELS[key]}）は有限の数値で必要です（不明なら決算書・勘定科目内訳明細書で確認）")
                bad = True
        if bad:
            continue
        if p["sales"] <= 0 or p["total_assets"] <= 0:
            errors.append(f"{tag}: sales と total_assets は正の数が必要です")
        for key in ("cost_of_sales", "sga", "depreciation", "interest_expense", "cash", "receivables", "inventories", "payables", "interest_bearing_debt", "officer_loans", "annual_principal_repayment"):
            if p[key] < 0:
                errors.append(f"{tag}.{key}（{FIELD_LABELS[key]}）は0以上が必要です")
        if p["net_assets"] > p["total_assets"]:
            errors.append(f"{tag}: 純資産が総資産を超えています（転記ミスの疑い）")
    if fin.get("unit") != "円":
        errors.append("unitは「円」にしてください（千円表示の決算書は1,000倍して転記）")
    if fin.get("verified_by_human") is True and not (str(fin.get("verified_by", "")).strip() and str(fin.get("verified_at", "")).strip()):
        errors.append("verified_by_human=true には verified_by（確認者）と verified_at（確認日）が必要です")
    if len(ends) == 3:
        if ends != sorted(ends) or len(set(ends)) != 3:
            errors.append("periodsは古い期から新しい期の順に、異なる決算日で並べてください")
        elif any(not 330 <= (b - a).days <= 400 for a, b in zip(ends, ends[1:])):
            errors.append("決算日の間隔が約1年ではありません（変則決算は手で年換算してから入力）")
    if not errors:
        sales = [p["sales"] for p in periods]
        if any(not 1 / 3 <= b / a <= 3 for a, b in zip(sales, sales[1:])):
            errors.append("売上高が前期比で3倍超または3分の1未満です（千円と円の混在の疑い）")
    for i, src in enumerate(fin.get("source_files", []), 1):
        if not str(src).strip():
            errors.append(f"source_files[{i}]が空です")
    return errors


def validate_case(case: dict[str, Any]) -> list[str]:
    errors = []
    cfg, cons = case["loan"], case["construction"]
    if len(cons["ramp"]) != 12:
        errors.append("construction.rampは12か月分が必要です")
    if not 0 < cons["gross_margin_floor"] <= cons["gross_margin_target"] < 1:
        errors.append("粗利率の下限・目標は 0 < 下限 <= 目標 < 1 が必要です")
    if any(not 0 < s["margin"] < 1 for s in cons["scenarios"]):
        errors.append("シナリオの粗利率は0〜1の範囲が必要です")
    if not 0 <= cfg["grace_months"] < cfg["term_months"] or not 1 <= cfg["disburse_month"] <= 12:
        errors.append("loan: 据置は期間未満、実行月はM1〜M12が必要です")
    if not 6 <= cfg["stress_stop_month"] <= 12 or cfg["min_cash_months"] < 0:
        errors.append("loan: stress_stop_monthは6〜12、min_cash_monthsは0以上が必要です")
    if case["funding_need"]["construction_wc"]["collection_lag_months"] < 1:
        errors.append("funding_need.construction_wc.collection_lag_monthsは1以上が必要です")
    return errors


def warnings_from(metrics: list[dict[str, Any]], capacity: dict[str, Any], case: dict[str, Any]) -> list[str]:
    out = []
    cfg = case["loan"]
    latest = metrics[-1]
    if latest["net_income"] > latest["ordinary_income"]:
        out.append("直近期の当期純利益が経常利益を上回っています（特別利益の可能性）。一時的な利益は返済原資として説明しにくい")
    if case["construction"]["gross_margin_target"] > latest["gross_margin"]:
        out.append(f"計画の粗利率{pct(case['construction']['gross_margin_target'])}が直近実績の売上総利益率{pct(latest['gross_margin'])}を上回ります。標準パッケージ化で達成する根拠の説明が必要です")
    start_year, start_month = map(int, case["year1"]["start"].split("-"))
    disburse = dt.date(start_year + (start_month - 1 + cfg["disburse_month"] - 1) // 12, (start_month - 1 + cfg["disburse_month"] - 1) % 12 + 1, 1)
    if (disburse - dt.date.fromisoformat(latest["period_end"])).days > 183:
        out.append("直近決算から融資実行まで6か月を超えます。直近の試算表（月次決算）も金融機関へ提出してください")
    if capacity.get("status") == "none" and latest["net_assets"] < 0:
        out.append("債務超過のため申込額を0にしました")
    if latest["net_income"] < 0:
        out.append("直近期が最終赤字です。金融機関は改善計画を重視します（資本性ローン・自己資金での開始も比較）")
    if latest["net_assets"] < 0:
        out.append("直近期が債務超過です。通常の新規融資は難しく、役員借入金の資本性認定や増資の検討が先です")
    if latest["debt_years"] > cfg["debt_years_limit"]:
        out.append(f"融資前の債務償還年数が{years_text(latest['debt_years'])}で目安{cfg['debt_years_limit']}年を超えています")
    if metrics[-1]["sales"] < metrics[-2]["sales"] * 0.9:
        out.append("直近期の売上が前期比10%超の減少です。減少理由と回復策の説明が必要です")
    if capacity["status"] != "full":
        out.append(f"資金需要{man(capacity['need_total'])}に対し借入余力{man(capacity['capacity'])}のため、申込額を{man(capacity['request'])}へ縮小しました")
    return out


# ---------- 計算 ----------

def period_metrics(p: dict[str, Any]) -> dict[str, Any]:
    m = {k: p[k] for k in FIELDS}
    m["label"], m["period_end"] = p["label"], p["period_end"]
    m["gross_profit"] = p["sales"] - p["cost_of_sales"]
    m["gross_margin"] = m["gross_profit"] / p["sales"]
    m["operating_income"] = m["gross_profit"] - p["sga"]
    m["operating_margin"] = m["operating_income"] / p["sales"]
    m["ordinary_margin"] = p["ordinary_income"] / p["sales"]
    m["simple_cf"] = min(p["net_income"], p["ordinary_income"]) + p["depreciation"]  # 特別利益で膨らませない
    m["working_capital"] = p["receivables"] + p["inventories"] - p["payables"]
    m["equity_ratio"] = p["net_assets"] / p["total_assets"]
    m["equity_ratio_adj"] = (p["net_assets"] + p["officer_loans"]) / p["total_assets"]
    net_debt = p["interest_bearing_debt"] - p["cash"]
    m["debt_years"] = (net_debt / m["simple_cf"]) if m["simple_cf"] > 0 else (0.0 if net_debt <= 0 else math.inf)
    return m


def loan_capacity(metrics: list[dict[str, Any]], cfg: dict[str, Any], need_total: int) -> dict[str, Any]:
    latest = metrics[-1]
    average_cf = sum(m["simple_cf"] for m in metrics) / 3
    base_cf = min(latest["simple_cf"], average_cf)  # 保守側: 直近と3期平均の小さい方
    amort_years = (cfg["term_months"] - cfg["grace_months"]) / 12
    min_cash = latest["sales"] / 12 * cfg["min_cash_months"]  # 最低限残す手元資金（月商×月数）
    annual_room = base_cf * cfg["repay_cover_ratio"] - latest["annual_principal_repayment"]
    by_years = cfg["debt_years_limit"] * base_cf + max(0.0, latest["cash"] - min_cash) - latest["interest_bearing_debt"]
    by_repay = annual_room / (1 / amort_years + cfg["annual_rate"])  # 新規の元金＋利息が返済余力に収まる額
    ok = base_cf > 0 and latest["net_assets"] >= 0
    capacity = floor_unit(min(by_years, by_repay), cfg["round_unit"]) if ok else 0
    request = min(need_total, capacity)
    status = "none" if request <= 0 else ("full" if request == need_total else "reduced")
    after_years = ((latest["interest_bearing_debt"] + request - max(0.0, latest["cash"] - min_cash)) / base_cf) if base_cf > 0 else math.inf
    return {
        "average_cf": average_cf, "base_cf": base_cf, "annual_room": annual_room, "by_years": by_years,
        "by_repay": by_repay, "capacity": capacity, "need_total": need_total, "request": request, "min_cash": min_cash,
        "status": status, "amort_years": amort_years, "debt_years_after": after_years,
        "new_annual_principal": request / amort_years if amort_years > 0 else 0,
        "new_annual_service": request / amort_years + request * cfg["annual_rate"] if amort_years > 0 else 0,
    }


def build_scenarios(case: dict[str, Any]) -> list[dict[str, Any]]:
    pricing_id = case["pricing"][0]["id"]
    out = []
    for s in case["construction"]["scenarios"]:
        unit_cost = round(s["price"] * (1 - s["margin"]))
        sales, cost = s["jobs"] * s["price"], s["jobs"] * unit_cost
        out.append({
            "scenario_type": s["type"], "name": s["name"], "pricing_basis": case["project"]["pricing_basis"],
            "components": [{"pricing_id": pricing_id, "volume": s["jobs"], "unit_price": s["price"], "unit_direct_cost": unit_cost}],
            "monthly_volume": s["jobs"], "average_unit_price": s["price"], "sales": sales, "direct_cost": cost,
            "gross_profit": sales - cost, "gross_margin": (sales - cost) / sales,
        })
    return out


def funding_need(case: dict[str, Any], base: dict[str, Any]) -> dict[str, Any]:
    need = case["funding_need"]
    items = []
    for item in need["capex"]:
        items.append({"category": "設備資金", "label": item.get("bank_label", item["label"]), "amount": item["monthly"] * len(item["months"]) if "monthly" in item else item["amount"], "note": item.get("note", "")})
    for item in need["opex"]:
        items.append({"category": "運転資金", "label": item.get("bank_label", item["label"]), "amount": item["monthly"] * item["funded_months"], "note": item.get("note", "")})
    comp = base["components"][0]
    lag = need["construction_wc"]["collection_lag_months"]
    wc = comp["volume"] * comp["unit_direct_cost"] * lag
    items.append({"category": "運転資金", "label": "増える工事の立替（原価先払い・入金待ち）", "amount": wc,
                  "note": f"標準{comp['volume']}件×原価{yen(comp['unit_direct_cost'])}円×入金まで{lag}か月"})
    subtotal = sum(i["amount"] for i in items)
    contingency = round(subtotal * need["contingency_rate"])
    items.append({"category": "運転資金", "label": "予備費", "amount": contingency, "note": f"小計の{pct(need['contingency_rate'])}"})
    total = ceil_unit(subtotal + contingency, case["loan"]["round_unit"])
    return {"items": items, "subtotal": subtotal, "total": total,
            "capex": sum(i["amount"] for i in items if i["category"] == "設備資金")}


def loan_schedule(amount: int, cfg: dict[str, Any]) -> list[dict[str, Any]]:
    """融資実行月の翌月から利息、据置後は元金均等。plan_month は M1 起点の月番号。"""
    rows, balance = [], float(amount)
    amort = cfg["term_months"] - cfg["grace_months"]
    principal_each = amount / amort if amort > 0 else 0
    for k in range(1, cfg["term_months"] + 1):
        interest = round(balance * cfg["annual_rate"] / 12)
        principal = round(min(principal_each, balance)) if k > cfg["grace_months"] else 0
        if k == cfg["term_months"]:
            principal = round(balance)
        balance -= principal
        rows.append({"k": k, "plan_month": cfg["disburse_month"] + k, "interest": interest, "principal": principal, "balance": round(balance)})
    return rows


def cash_flow(case: dict[str, Any], scenario: dict[str, Any], schedule: list[dict[str, Any]], request: int, core_monthly: float, stress: bool = False) -> list[dict[str, Any]]:
    """stress=True は『新規事業経由の売上ゼロ、M6でStopし継続費用は解約予告の stress_stop_month まで』の試算。"""
    cfg, need, cons = case["loan"], case["funding_need"], case["construction"]
    comp = scenario["components"][0]
    lag = need["construction_wc"]["collection_lag_months"]
    by_month = {r["plan_month"]: r for r in schedule}
    jobs = [0] * 12 if stress else [round(comp["volume"] * f) for f in cons["ramp"]]
    stop = cfg["stress_stop_month"] if stress else 12
    rows, cumulative, company_cum, op_cum, excl_cum = [], 0.0, 0.0, 0.0, 0.0
    for m in range(1, 13):
        row: dict[str, Any] = {"month": m, "jobs": jobs[m - 1]}
        row["loan_in"] = request if m == cfg["disburse_month"] else 0
        collected_jobs = jobs[m - 1 - lag] if m - 1 - lag >= 0 else 0
        row["collection"] = collected_jobs * comp["unit_price"]
        row["construction_cost"] = jobs[m - 1] * comp["unit_direct_cost"]
        spend = 0
        for item in need["capex"]:
            if "monthly" in item and m in item["months"]:
                spend += item["monthly"]
            elif "amount" in item and item.get("month") == m:
                spend += item["amount"]
        for item in need["opex"]:
            if m in item["cash_months"] and m <= stop:
                spend += item["monthly"]
        row["revo_spend"] = spend
        loan = by_month.get(m, {"interest": 0, "principal": 0})
        row["interest"], row["principal"] = loan["interest"], loan["principal"]
        row["op_net"] = row["collection"] - row["construction_cost"] - spend  # 融資・返済を除く営業収支
        op_cum += row["op_net"]
        row["op_cumulative"] = op_cum
        row["revo_net"] = row["loan_in"] + row["op_net"] - row["interest"] - row["principal"]
        cumulative += row["revo_net"]
        row["revo_cumulative"] = cumulative
        row["core"] = core_monthly
        company_cum += row["revo_net"] + core_monthly
        row["company_cumulative"] = company_cum
        excl_cum += row["revo_net"] - row["loan_in"] + core_monthly
        row["company_cumulative_excl_loan"] = excl_cum
        rows.append(row)
    return rows


# ---------- control-sheet の組み立て ----------

def month_label(start: str, m: int) -> str:
    year, month = map(int, start.split("-"))
    total = year * 12 + (month - 1) + (m - 1)
    return f"{total // 12}年{total % 12 + 1}月"


def fill(value: Any, ctx: dict[str, str]) -> Any:
    if isinstance(value, str):
        return value.format_map(ctx)
    if isinstance(value, list):
        return [fill(v, ctx) for v in value]
    if isinstance(value, dict):
        return {k: fill(v, ctx) for k, v in value.items()}
    return value


def make_ctx(case: dict[str, Any], fin: dict[str, Any], calc: dict[str, Any]) -> dict[str, str]:
    cfg, cap = case["loan"], calc["capacity"]
    return {
        "request": man(cap["request"]), "need_total": man(calc["need"]["total"]), "capacity": man(cap["capacity"]),
        "scale_sales": man(calc["scale_sales"]), "debt_years_after": years_text(cap["debt_years_after"]),
        "dev_fee": man(calc["dev_fee"]), "term_years": f"{cfg['term_months']}か月", "grace": f"{cfg['grace_months']}か月",
        "rate": pct(cfg["annual_rate"]), "company": fin["company"],
    }


def build_control(case: dict[str, Any], fin: dict[str, Any], calc: dict[str, Any], today: dt.date) -> dict[str, Any]:
    data = plan_gate.template()
    basis = case["project"]["pricing_basis"]
    cfg, cap, need = case["loan"], calc["capacity"], calc["need"]
    latest, metrics = calc["metrics"][-1], calc["metrics"]
    verified = fin.get("verified_by_human") is True
    data["status"] = "validated" if verified else "draft"
    data["project"].update(case["project"])
    data["project"]["as_of"] = today.isoformat()
    if not verified:
        data["project"]["objective"] += "（決算書数値は人間未確認のドラフト）"
    data["source_files"] = list(dict.fromkeys(case["source_files"] + fin.get("source_files", [])))

    lock = copy.deepcopy(case["decision_lock"])
    for ref in lock["source_refs"]:
        if "conversation" in ref:
            ref["conversation"]["sha256"] = hashlib.sha256(ref["conversation"]["excerpt"].encode("utf-8")).hexdigest()
    data["decision_lock"] = lock

    periods_label = "・".join(p["label"] for p in fin["periods"])
    client_src = {
        "id": "SFIN", "kind": "client", "title": f"{fin['company']} 決算書（{periods_label}）", "publisher": fin["company"],
        "url": "", "file": (fin.get("source_files") or ["financials.json"])[0], "published_at": latest["period_end"],
        "accessed_at": today.isoformat(), "scope": "決算3期分の貸借対照表・損益計算書・借入金明細", "population": fin["company"],
        "unit": "円", "pricing_basis": "not_applicable",
    }
    data["sources"] = copy.deepcopy(case["sources"]) + [client_src]

    ctx = make_ctx(case, fin, calc)
    fin_claims = [
        {"id": "CF1", "type": "fact", "claim": f"直近期（{latest['label']}）の売上高", "value": yen(latest["sales"]), "unit": "円", "as_of": latest["period_end"], "source_ids": ["SFIN"], "confidence": "high" if verified else "medium"},
        {"id": "CF2", "type": "fact", "claim": "簡易キャッシュフロー（税引後利益＋減価償却費）の3期平均", "value": yen(cap["average_cf"]), "unit": "円/年", "as_of": latest["period_end"], "source_ids": ["SFIN"], "confidence": "high" if verified else "medium"},
        {"id": "CF3", "type": "fact", "claim": f"直近期の有利子負債（役員借入金を除く）と現預金の差（純有利子負債）", "value": yen(latest["interest_bearing_debt"] - latest["cash"]), "unit": "円", "as_of": latest["period_end"], "source_ids": ["SFIN"], "confidence": "high" if verified else "medium"},
        {"id": "CF4", "type": "fact", "claim": "融資前の債務償還年数（純有利子負債÷簡易CF）", "value": years_text(latest["debt_years"]), "unit": "", "as_of": latest["period_end"], "source_ids": ["SFIN"], "confidence": "high" if verified else "medium"},
        {"id": "CA1", "type": "assumption", "claim": f"本業の返済余力（年）＝保守側簡易CF×{pct(cfg['repay_cover_ratio'])}−既存元金返済", "value": yen(cap["annual_room"]), "unit": "円/年", "as_of": today.isoformat(), "source_ids": ["SFIN"], "confidence": "medium"},
        {"id": "CT1", "type": "target", "claim": "新規事業経由の工事で返済開始後も採算が合う月売上（拡大ゲート）", "value": yen(calc["scale_sales"]), "unit": "円/月（税別）", "as_of": today.isoformat(), "source_ids": [], "confidence": "medium"},
    ]
    for claim in fin_claims:
        claim["unit"] = claim["unit"] or "年"
    data["claims"] = fill(copy.deepcopy(case["claims"]), ctx) + fin_claims
    data["offer"].update(case["offer"])
    data["pricing"] = copy.deepcopy(case["pricing"])
    data["contract"] = copy.deepcopy(case["contract"])

    margin = case["construction"]["gross_margin_target"]
    data["financials"] = {
        "pricing_basis": basis,
        "monthly_fixed_cost": {"min": calc["fixed_min"], "max": calc["fixed_max"]},
        "target_gross_margin": margin, "gross_margin_floor": case["construction"]["gross_margin_floor"],
        "break_even_sales": {"min": round(calc["fixed_min"] / margin), "max": round(calc["fixed_max"] / margin)},
        "scale_gate": {"monthly_sales": calc["scale_sales"], "gross_margin": margin, "consecutive_months": 3},
        "loan": {"request": cap["request"], "capacity": cap["capacity"], "need_total": need["total"], "status": cap["status"],
                 "term_months": cfg["term_months"], "grace_months": cfg["grace_months"], "annual_rate": cfg["annual_rate"]},
    }
    data["scenarios"] = calc["scenarios"]

    for spec, month in zip(case["year1"]["months"], data["year1"]["months"]):
        spec = fill(spec, ctx)
        month.update({k: spec[k] for k in ("phase", "deliverables", "kpis", "evidence", "owner")})
        month["phase"] = f"{month_label(case['year1']['start'], month['month'])}｜{spec['phase']}"
        gate = spec["gate"]
        month["gate"].update({"decision": gate["decision"], "conditions": gate["conditions"], "if_failed": gate["if_failed"]})
        if gate.get("expansion"):
            month["gate"].update({"expansion": True, "depends_on": sorted({f"M{month['month'] - 1}", "M6"}), "required_prior_outcome": "Go"})
    data["year1"]["focus"] = fill(case["year1"]["focus"], ctx)
    data["year1"]["success"] = {"拡大ゲート": f"新規事業経由の工事で月売上{man(calc['scale_sales'])}（税別）・粗利率{pct(margin)}以上を3か月連続", "返済": "融資の返済遅延ゼロ（返済は本業の簡易CFで賄う設計。新規事業不振時はM6で停止する前提、金融機関向け別冊§6のストレス試算）", "外部検証": "他社2社以上が3か月以上有料で継続"}
    data["year1"]["minimum_viable"] = {"完工": "累計15件以上（M6）", "粗利率": "平均25%以上", "データ": "写真登録80%以上・原価登録100%"}
    data["years_2_3"] = fill(copy.deepcopy(case["years_2_3"]), ctx)
    data["risks"] = fill(copy.deepcopy(case["risks"]), ctx)
    data["human_gates"] = fill(copy.deepcopy(case["human_gates"]), ctx)
    data["open_questions"] = fill(copy.deepcopy(case.get("open_questions", [])), ctx)
    required = ["plan", "market_pricing", "pptx", "pdf"]
    critical = [
        ("K1", "融資申込額", man(cap["request"])),
        ("K2", "開発増強費（推奨額・顧客未合意）", man(calc["dev_fee"])),
        ("K3", "融資後の債務償還年数", years_text(cap["debt_years_after"])),
        ("K4", "本業の返済余力（年）", man(cap["annual_room"])),
        ("K5", "拡大ゲート月売上（返済開始後の損益分岐を10万円単位で切上げ）", man(calc["scale_sales"])),
    ]
    data["critical_claims"] = [{"id": i, "label": label, "canonical_value": value, "aliases": [value], "required_in": required} for i, label, value in critical]
    data["rendering"]["author"] = case["author"]
    return data


# ---------- 別冊（銀行向け） ----------

def appendix_markdown(case: dict[str, Any], fin: dict[str, Any], calc: dict[str, Any], today: dt.date) -> str:
    cfg, cap, need = case["loan"], calc["capacity"], calc["need"]
    metrics = calc["metrics"]
    base = next(s for s in calc["scenarios"] if s["scenario_type"] == "base")
    verified = fin.get("verified_by_human") is True
    L: list[str] = [f"# {case['project']['name']}｜金融機関向け別冊（資金計画・返済計画）", ""]
    L += [f"- 作成日: {today.isoformat()}", f"- 対象会社: {fin['company']}", f"- 金額基準: 税別・円",
          f"- 決算書数値の確認: {'人間確認済み' if verified else '**未確認（ドラフト。提出前に決算書と照合すること）**'}", ""]
    L += ["## 1. 融資申込の概要", "", "| 項目 | 内容 |", "|---|---|",
          f"| 申込額 | **{man(cap['request'])}**（資金需要 {man(need['total'])}、借入余力 {man(cap['capacity'])}） |",
          f"| 調達の内訳 | 融資 {man(cap['request'])}／自己資金 {man(need['total'] - cap['request'])} |",
          f"| 融資の使途 | 設備資金 {man(min(need['capex'], cap['request']))}／運転資金 {man(max(0, cap['request'] - need['capex']))} |",
          f"| 希望条件（仮定） | 期間{cfg['term_months']}か月・据置{cfg['grace_months']}か月・元金均等・金利{pct(cfg['annual_rate'])}（金利は仮定。実際は金融機関の提示による） |",
          f"| 実行予定 | M{cfg['disburse_month']}（{month_label(case['year1']['start'], cfg['disburse_month'])}） |",
          f"| 返済原資 | 本業の簡易CF（保守側 {man(cap['base_cf'])}/年）。新規の元金＋利息（年{man(round(cap['new_annual_service']))}）が、既存返済を除いた本業CFの{pct(cfg['repay_cover_ratio'])}以内に収まる額を上限にした。新規事業の売上が出ない場合はM6で新規事業を止め、継続費用はM{cfg['stress_stop_month']}までで打ち切る前提（§6のストレス試算） |", ""]
    L += ["### 候補となる制度（最新条件は申込前に窓口で確認）", ""] + [f"- {p}" for p in case["loan"]["products"]] + [""]
    if case.get("dev_cadence"):
        L += ["### 開発体制と進め方（実行力の根拠）", ""] + [f"- {c}" for c in case["dev_cadence"]] + [""]
    L += ["## 2. 資金使途の内訳", "", "| 区分 | 項目 | 金額 | 根拠 |", "|---|---|---:|---|"]
    for item in need["items"]:
        L.append(f"| {item['category']} | {item['label']} | {yen(item['amount'])} | {item['note']} |")
    L += [f"| 合計 | （10万円単位に切上げ） | {yen(need['total'])} | 小計 {yen(need['subtotal'])} |", ""]
    L += ["## 3. 決算3期の財務分析", "", "| 項目 | " + " | ".join(m["label"] for m in metrics) + " |", "|---|" + "---:|" * 3]
    rows = [("売上高", "sales", yen), ("売上総利益", "gross_profit", yen), ("売上総利益率", "gross_margin", pct),
            ("営業利益", "operating_income", yen), ("経常利益", "ordinary_income", yen), ("当期純利益", "net_income", yen),
            ("減価償却費", "depreciation", yen), ("簡易CF（min(純利益,経常利益)＋償却）", "simple_cf", yen), ("現預金", "cash", yen),
            ("有利子負債（役員借入金を除く）", "interest_bearing_debt", yen), ("役員借入金", "officer_loans", yen),
            ("年間元金返済", "annual_principal_repayment", yen), ("正常運転資金", "working_capital", yen),
            ("自己資本比率", "equity_ratio", pct), ("参考：役員借入金を資本とみなした場合", "equity_ratio_adj", pct),
            ("債務償還年数（純有利子負債÷簡易CF）", "debt_years", years_text)]
    for label, key, fmt in rows:
        L.append(f"| {label} | " + " | ".join(fmt(m[key]) for m in metrics) + " |")
    L += ["", "## 4. 借入余力の計算", "", "```text",
          f"保守側の簡易CF = min(直近 {yen(metrics[-1]['simple_cf'])}, 3期平均 {yen(cap['average_cf'])}) = {yen(cap['base_cf'])}",
          f"最低手元資金 = 月商×{cfg['min_cash_months']}か月 = {yen(round(cap['min_cash']))}",
          f"(a) 債務償還年数{cfg['debt_years_limit']}年以内: {cfg['debt_years_limit']}×簡易CF＋(現預金−最低手元資金)−有利子負債 = {yen(cap['by_years'])}",
          f"(b) 返済余力: 年{yen(cap['annual_room'])}（簡易CF×{cfg['repay_cover_ratio']}−既存元金返済 {yen(metrics[-1]['annual_principal_repayment'])}）÷（1/{cap['amort_years']:.1f}年＋金利{pct(cfg['annual_rate'])}） = {yen(cap['by_repay'])}",
          f"借入余力 = min(a, b) を{cfg['round_unit'] // 10000}万円単位で切捨て = {yen(cap['capacity'])}",
          f"申込額 = min(資金需要 {yen(need['total'])}, 借入余力) = {yen(cap['request'])}",
          f"融資後の債務償還年数（現預金は最低手元資金を控除） = {years_text(cap['debt_years_after'])} ／ 新規借入の年間元利 = {yen(round(cap['new_annual_service']))}",
          "```", "", "目安（債務償還年数10年以内・返済は簡易CFの8割以内）は一般的な審査の考え方を置いた仮定で、金融機関ごとに異なる。", ""]
    if calc["warnings"]:
        L += ["### 注意事項", ""] + [f"- {w}" for w in calc["warnings"]] + [""]
    L += ["## 5. 返済計画（年次）", "", "| 年 | 元金 | 利息 | 年末残高 |", "|---:|---:|---:|---:|"]
    for y in range(0, math.ceil(len(calc["schedule"]) / 12)):
        chunk = calc["schedule"][y * 12:(y + 1) * 12]
        L.append(f"| 実行後{y + 1}年目 | {yen(sum(r['principal'] for r in chunk))} | {yen(sum(r['interest'] for r in chunk))} | {yen(chunk[-1]['balance'])} |")
    L += ["", "## 6. 12か月の資金繰り表（標準シナリオ・新規事業分）", "",
          "| 月 | 件数 | 融資 | 工事入金 | 工事原価 | 新規事業費用 | 利息 | 元金 | 営業収支（融資除く） | 営業収支の累計 | 会社累計（融資除く・本業込） | 会社累計（融資込） |",
          "|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|"]
    for r in calc["cash_flows"]["base"]:
        L.append(f"| M{r['month']}（{month_label(case['year1']['start'], r['month'])}） | {r['jobs']} | {yen(r['loan_in'])} | {yen(r['collection'])} | {yen(r['construction_cost'])} | {yen(r['revo_spend'])} | {yen(r['interest'])} | {yen(r['principal'])} | {yen(r['op_net'])} | {yen(r['op_cumulative'])} | {yen(r['company_cumulative_excl_loan'])} | {yen(r['company_cumulative'])} |")
    L += ["", f"営業収支は工事入金−工事原価−新規事業費用で、融資と返済を含まない。会社累計は本業の簡易CF（月割 {yen(round(calc['core_monthly']))}円、既存元金返済を差引済み）を足した増減で、元利返済を含む。工事入金は完工の{case['funding_need']['construction_wc']['collection_lag_months']}か月後、原価は当月払いの保守的な置き方。新規事業経由の工事は既存取引と別の増分として数え、既存工事の置き換えは含めない（実績で検証する）。", "",
          "### シナリオ別の12か月比較（ストレスを含む）", "", "| シナリオ | 定常の月件数 | 12か月の営業収支（融資除く） | M12末の入金待ち工事代金 | 会社累計（融資除く・本業込） | 会社累計（融資込） | 融資込み資金の最小値（新規事業分） |", "|---|---:|---:|---:|---:|---:|---:|"]
    lag = case["funding_need"]["construction_wc"]["collection_lag_months"]
    for key, name, vol, price in [(s["scenario_type"], s["name"], s["monthly_volume"], s["average_unit_price"]) for s in calc["scenarios"]] + [("stress", f"ストレス（売上ゼロ・M6でStop、費用はM{cfg['stress_stop_month']}まで）", 0, 0)]:
        rows_ = calc["cash_flows"][key]
        pending = sum(r["jobs"] for r in rows_[-lag:]) * price
        L.append(f"| {name} | {vol} | {yen(rows_[-1]['op_cumulative'])} | {yen(pending)} | {yen(rows_[-1]['company_cumulative_excl_loan'])} | {yen(rows_[-1]['company_cumulative'])} | {yen(min(r['revo_cumulative'] for r in rows_))} |")
    L += ["", "現金ベースのため、完工済みで入金待ちの工事代金はM12時点の累計に含まれない。工事を増やすシナリオほど入金待ちが増え、1年目の現金は一時的に少なく見える。"]
    L += ["", "## 7. 新規事業の採算", "",
          f"- 新規事業の月間固定キャッシュ負担: 据置中 {man(calc['fixed_min'])}／返済開始後 {man(calc['fixed_max'])}",
          f"- 粗利率{pct(case['construction']['gross_margin_target'])}での損益分岐月売上: {man(round(calc['fixed_min'] / case['construction']['gross_margin_target']))}〜{man(round(calc['fixed_max'] / case['construction']['gross_margin_target']))}",
          f"- 拡大ゲート: 月売上{man(calc['scale_sales'])}（税別）・粗利率{pct(case['construction']['gross_margin_target'])}以上を3か月連続",
          f"- 標準シナリオ（月{base['monthly_volume']}件・月売上{man(base['sales'])}）は、返済開始後の損益分岐に届かない。新規事業単体の黒字化には月{calc['jobs_to_break_even']}件前後が必要で、届くまでの差額は本業の簡易CFで賄う",
          f"- 計画粗利率{pct(case['construction']['gross_margin_target'])}に対し、直近期の売上総利益率は{pct(metrics[-1]['gross_margin'])}。標準パッケージ化と見積の標準化で達成を目指す目標値で、M5の実績で見直す", ""]
    L += ["## 8. 前提と免責", "", "本書の数値は計画値であり、融資の実行・金利・売上・利益を保証しない。決算書の数値は原本と照合し、制度の条件は申込時点の公表内容と金融機関の説明を優先する。税務・法務の最終判断は税理士・弁護士が行う。", ""]
    return "\n".join(L)


def appendix_pdf(markdown: str, path: Path) -> None:
    from render_package import pdf_font  # 共通レンダラーのCJK埋込フォント
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import Paragraph, Preformatted, SimpleDocTemplate, Spacer, Table, TableStyle
    from xml.sax.saxutils import escape

    font, _ = pdf_font()
    body = ParagraphStyle("b", fontName=font, fontSize=8.5, leading=12)
    cell = ParagraphStyle("c", fontName=font, fontSize=7, leading=9)
    h1 = ParagraphStyle("h1", fontName=font, fontSize=15, leading=20, spaceAfter=6)
    h2 = ParagraphStyle("h2", fontName=font, fontSize=11.5, leading=16, spaceBefore=8, spaceAfter=4)
    code = ParagraphStyle("code", fontName=font, fontSize=7.5, leading=10)
    doc = SimpleDocTemplate(str(path), pagesize=landscape(A4), leftMargin=12 * mm, rightMargin=12 * mm, topMargin=10 * mm, bottomMargin=10 * mm,
                            title=path.stem, author="bank-loan-plan")
    width = landscape(A4)[0] - 24 * mm

    def inline(text: str) -> str:
        return escape(text).replace("**", "")

    story: list[Any] = []
    lines, i = markdown.splitlines(), 0
    while i < len(lines):
        line = lines[i]
        if line.startswith("|"):
            block = []
            while i < len(lines) and lines[i].startswith("|"):
                if not set(lines[i].replace("|", "").strip()) <= set("-: "):
                    block.append([Paragraph(inline(c.strip()), cell) for c in lines[i].strip().strip("|").split("|")])
                i += 1
            cols = len(block[0])
            table = Table(block, colWidths=[width / cols] * cols, repeatRows=1)
            table.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.3, colors.grey), ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EFE8DA")), ("VALIGN", (0, 0), (-1, -1), "TOP")]))
            story += [table, Spacer(1, 4)]
            continue
        if line.startswith("```"):
            i += 1
            block = []
            while i < len(lines) and not lines[i].startswith("```"):
                block.append(lines[i])
                i += 1
            story.append(Preformatted("\n".join(block), code))
        elif line.startswith("# "):
            story.append(Paragraph(inline(line[2:]), h1))
        elif line.startswith("## ") or line.startswith("### "):
            story.append(Paragraph(inline(line.lstrip("#").strip()), h2))
        elif line.startswith("- "):
            story.append(Paragraph("・" + inline(line[2:]), body))
        elif line.strip():
            story.append(Paragraph(inline(line), body))
        i += 1
    doc.build(story)


def contract_checklist(case: dict[str, Any], calc: dict[str, Any]) -> str:
    lines = [f"# {case['project']['name']}｜契約確認チェックリスト", "", f"追加開発費の推奨額: {man(calc['dev_fee'])}（顧客未合意）", ""]
    lines += [f"- [ ] {item}" for item in case["contract_checklist"]]
    return "\n".join(lines + ["", "法的な最終判断は実在の弁護士・税理士が行う。", ""])


# ---------- 実行 ----------

def compute(case: dict[str, Any], fin: dict[str, Any]) -> dict[str, Any]:
    metrics = [period_metrics(p) for p in fin["periods"]]
    scenarios = build_scenarios(case)
    base = next(s for s in scenarios if s["scenario_type"] == "base")
    need = funding_need(case, base)
    cfg = case["loan"]
    cap = loan_capacity(metrics, cfg, need["total"])
    schedule = loan_schedule(cap["request"], cfg) if cap["request"] > 0 else []
    core_monthly = (cap["base_cf"] - metrics[-1]["annual_principal_repayment"]) / 12
    flows = {s["scenario_type"]: cash_flow(case, s, schedule, cap["request"], core_monthly) for s in scenarios}
    flows["stress"] = cash_flow(case, base, schedule, cap["request"], core_monthly, stress=True)
    recurring = sum(i["monthly"] for i in case["funding_need"]["opex"] if len(i["cash_months"]) == 12)
    interest_full = round(cap["request"] * cfg["annual_rate"] / 12)
    principal = round(cap["request"] / (cfg["term_months"] - cfg["grace_months"])) if cap["request"] > 0 else 0
    fixed_min, fixed_max = recurring + interest_full, recurring + interest_full + principal
    margin = case["construction"]["gross_margin_target"]
    dev = case["funding_need"]["capex"][0]
    calc = {"metrics": metrics, "scenarios": scenarios, "need": need, "capacity": cap, "schedule": schedule,
            "cash_flows": flows, "core_monthly": core_monthly, "fixed_min": fixed_min, "fixed_max": fixed_max,
            "scale_sales": ceil_unit(fixed_max / margin, 100000),
            "dev_fee": dev["monthly"] * len(dev["months"]) if "monthly" in dev else dev["amount"]}
    comp = base["components"][0]
    calc["jobs_to_break_even"] = math.ceil(fixed_max / (comp["unit_price"] - comp["unit_direct_cost"]))
    calc["warnings"] = warnings_from(metrics, cap, case)
    return calc


def run(cmd: list[str]) -> None:
    result = subprocess.run(cmd, capture_output=True, text=True)
    sys.stdout.write(result.stdout)
    if result.returncode != 0:
        sys.stderr.write(result.stderr)
        raise SystemExit(f"[FAIL] {' '.join(cmd[:3])} ... exit {result.returncode}")


def build(case_path: Path, fin_path: Path, out: Path, force: bool, images: bool, allow_zero: bool = False) -> int:
    case, fin = load_json(case_path), load_json(fin_path)
    errors = validate_financials(fin) + validate_case(case)
    if errors:
        for e in errors:
            print(f"[FAIL] {e}")
        return 1
    control = out / "control-sheet.json"
    if control.exists() and not force:
        print(f"[FAIL] 既存の control-sheet.json を上書きしません（再生成は --force）: {control}")
        return 1
    out.mkdir(parents=True, exist_ok=True)
    today = dt.datetime.now(JST).date()
    calc = compute(case, fin)
    if calc["capacity"]["status"] == "none" and not allow_zero:
        for w in calc["warnings"]:
            print(f"[WARN] {w}")
        print("[FAIL] 借入余力が0のため生成を止めました（赤字・債務超過・既存返済が重い等）。自己資金や資本性ローンで組み直すか、--allow-zero で検討用ドラフトを作ってください")
        return 1
    data = build_control(case, fin, calc, today)
    if calc["capacity"]["status"] != "full":
        data["status"] = "draft"
    control.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    try:
        run([sys.executable, str(BBP_SCRIPTS / "plan_gate.py"), "check", str(control)])
        run([sys.executable, str(BBP_SCRIPTS / "render_package.py"), str(control)] + (["--force"] if force else []))
    except SystemExit:
        if not force:
            control.unlink(missing_ok=True)  # 失敗した生成物を残さず、次の実行を妨げない
        raise
    slug = data["project"]["slug"]
    md = appendix_markdown(case, fin, calc, today)
    if "{" in md.replace("```", ""):
        raise SystemExit("[FAIL] 別冊に未展開のプレースホルダーがあります")
    (out / f"{slug}-bank-appendix.md").write_text(md, encoding="utf-8")
    appendix_pdf(md, out / f"{slug}-bank-appendix.pdf")
    if data["contract"].get("required"):
        (out / f"{slug}-contract-checklist.md").write_text(contract_checklist(case, calc), encoding="utf-8")
    if images:
        make_images(out, slug)
    cap = calc["capacity"]
    print(f"[OK] 申込額 {man(cap['request'])}（需要 {man(calc['need']['total'])}／余力 {man(cap['capacity'])}／{cap['status']}）"
          f"・融資後の債務償還年数 {years_text(cap['debt_years_after'])}・status={data['status']}")
    for w in calc["warnings"]:
        print(f"[WARN] {w}")
    return 0


def make_images(out: Path, slug: str) -> None:
    """レビュー証拠用に PDF と PPTX の全ページを画像化する（pdftoppm / soffice が必要）。"""
    evidence = out / "evidence"
    evidence.mkdir(exist_ok=True)
    run(["pdftoppm", "-png", "-r", "60", str(out / f"{slug}-presentation.pdf"), str(evidence / "pdf")])
    run(["pdftoppm", "-png", "-r", "60", str(out / f"{slug}-bank-appendix.pdf"), str(evidence / "appendix")])
    with tempfile.TemporaryDirectory() as tmp:
        run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, str(out / f"{slug}-presentation.pptx")])
        run(["pdftoppm", "-png", "-r", "60", str(Path(tmp) / f"{slug}-presentation.pdf"), str(evidence / "pptx")])


def self_test() -> int:
    """架空の決算書で計算式と生成パイプラインを検査する。"""
    case = load_json(SKILL_DIR / "cases" / "example" / "case.json")
    fin = load_json(SKILL_DIR / "examples" / "sample-financials.json")
    assert not validate_financials(fin), validate_financials(fin)
    bad = copy.deepcopy(fin)
    bad["periods"] = bad["periods"][:2]
    assert validate_financials(bad), "2期しかない入力を拒否できていない"
    bad = copy.deepcopy(fin)
    bad["periods"][0]["sales"] = None
    assert validate_financials(bad), "欠損値を拒否できていない"
    for key, value in (("sales", float("nan")), ("cost_of_sales", -1), ("net_assets", 10**12)):
        bad = copy.deepcopy(fin)
        bad["periods"][2][key] = value
        assert validate_financials(bad), f"{key}={value} を拒否できていない"
    bad = copy.deepcopy(fin)
    bad["unit"] = "千円"
    assert validate_financials(bad), "千円単位を拒否できていない"
    bad = copy.deepcopy(fin)
    bad["periods"][2]["sales"] = fin["periods"][2]["sales"] // 1000
    assert validate_financials(bad), "桁違いの売上を拒否できていない"
    calc = compute(case, fin)
    cap, cfg = calc["capacity"], case["loan"]
    assert cap["new_annual_service"] <= cap["annual_room"] + 1, "新規元利返済が本業の返済余力を超えている"
    stress = calc["cash_flows"]["stress"]
    assert all(r["jobs"] == 0 for r in stress) and all(r["revo_spend"] == 0 for r in stress if r["month"] > cfg["stress_stop_month"])
    base_rows = calc["cash_flows"]["base"]
    assert all(abs(r["company_cumulative"] - r["company_cumulative_excl_loan"] - (cap["request"] if r["month"] >= cfg["disburse_month"] else 0)) < 1 for r in base_rows)
    m = calc["metrics"]
    assert cap["base_cf"] == min(m[-1]["simple_cf"], sum(x["simple_cf"] for x in m) / 3)
    assert cap["request"] <= cap["capacity"] and cap["request"] <= cap["need_total"]
    assert cap["request"] % cfg["round_unit"] == 0
    assert sum(r["principal"] for r in calc["schedule"]) == cap["request"], "元金の合計が申込額と一致しない"
    assert all(r["principal"] == 0 for r in calc["schedule"][: cfg["grace_months"]]), "据置中に元金がある"
    assert calc["schedule"][-1]["balance"] == 0
    assert cap["new_annual_principal"] <= cap["annual_room"] + 1, "新規返済が本業の返済余力を超えている"
    insolvent = copy.deepcopy(fin)
    insolvent["periods"][2]["net_assets"] = -1
    assert compute(case, insolvent)["capacity"]["request"] == 0, "債務超過で申込額が0にならない"
    # 赤字会社は借入余力0になる
    loss = copy.deepcopy(fin)
    for p in loss["periods"]:
        p["net_income"], p["depreciation"] = -5_000_000, 1_000_000
    assert compute(case, loss)["capacity"]["request"] == 0
    # 余力が需要を下回ると縮小される
    small = copy.deepcopy(fin)
    for p in small["periods"]:
        p["net_income"] = 1_000_000
    small_cap = compute(case, small)["capacity"]
    assert small_cap["status"] in {"reduced", "none"} and small_cap["request"] <= small_cap["capacity"]
    with tempfile.TemporaryDirectory() as tmp:
        code = build(SKILL_DIR / "cases" / "example" / "case.json", SKILL_DIR / "examples" / "sample-financials.json", Path(tmp) / "out", False, False)
        assert code == 0
        out_dir = Path(tmp) / "out"
        files = {p.name for p in out_dir.iterdir()}
        control = load_json(out_dir / "control-sheet.json")
        appendix = (out_dir / "sample-bank-loan-bank-appendix.md").read_text(encoding="utf-8")
        for claim in control["critical_claims"]:
            if claim["id"] in {"K1", "K3"}:
                assert claim["canonical_value"] in appendix, f"別冊に{claim['label']}（{claim['canonical_value']}）がない"
        for suffix in ("-business-plan.md", "-market-pricing.md", "-presentation.pptx", "-presentation.pdf", "-bank-appendix.md", "-bank-appendix.pdf", "-contract-checklist.md"):
            assert any(name.endswith(suffix) for name in files), f"{suffix} が生成されていない"
    print("[PASS] self-test: 計算式・入力検証・赤字/縮小ケース・生成パイプライン")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--case", type=Path, default=SKILL_DIR / "cases" / "example" / "case.json")
    parser.add_argument("--financials", type=Path)
    parser.add_argument("--out", type=Path)
    parser.add_argument("--force", action="store_true", help="人間が差分を確認した後の再生成だけに使う")
    parser.add_argument("--images", action="store_true", help="レビュー証拠用に全ページを画像化する")
    parser.add_argument("--allow-zero", action="store_true", help="借入余力0でも検討用ドラフトを作る")
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()
    if args.self_test:
        return self_test()
    if not args.financials or not args.out:
        parser.error("--financials と --out が必要です")
    return build(args.case, args.financials, args.out, args.force, args.images, args.allow_zero)


if __name__ == "__main__":
    raise SystemExit(main())
