---
name: bank-loan-plan
description: "決算書3期分を入れると、借入余力・融資申込額・資金使途・返済計画・12か月資金繰り表・3シナリオを計算し、金融機関に出せる事業計画パッケージ（事業計画md・市場価格メモ・PPTX・PDF・銀行向け別冊PDF・契約チェックリスト）を生成する。「決算書を入れた」「融資用の事業計画」「銀行に出す計画」「借入余力」「返済計画」「資金繰り表」で使う。事前に build-business-plan の plan_gate.py / render_package.py が必要。"
---

# 融資用事業計画（決算書3期 → 計画パッケージ）

数値の計算は `scripts/build_loan_plan.py` だけが行う。文章・スライドに数字を手入力しない。

## 手順

1. `templates/financials.template.json` をコピーして `financials.json` にし、決算書3期分（古い期→新しい期）を円単位で転記する。原本と照合できたら `verified_by_human`・`verified_by`・`verified_at` を入れる。空欄・千円単位・桁違い・純資産＞総資産はスクリプトが止める。
2. 生成する。
   ```bash
   export BUSINESS_PLAN_SCRIPTS=/path/to/build-business-plan/scripts   # plan_gate.py と render_package.py の場所
   python3 scripts/build_loan_plan.py --case cases/example/case.json --financials financials.json --out ./out --images
   ```
   既存の出力は上書きしない（再生成は人が差分を確認したうえで `--force`）。
3. 別のモデル・担当が `plan_gate.py record-review` → `check --artifacts` で検証し、金融機関へ出す前に人が確認する。

## 計算の中身（`case.json` の `loan` で調整）

- 保守側の簡易CF = min(直近期, 3期平均) の（min(純利益, 経常利益)＋減価償却費）
- 借入余力 = min(債務償還年数の上限［現預金から最低手元資金を控除］, 返済余力÷(1/返済年数＋金利)) を切り捨て。返済余力＝簡易CF×0.8−既存の年間元金返済。
- 申込額 = min(資金需要, 借入余力)。不足は自己資金として表示し `draft` にする。簡易CFが0以下または債務超過なら余力0となり生成を止める（`--allow-zero` で検討用ドラフトのみ）。
- 資金繰り表は、融資を除いた営業収支・融資込みの累計・ストレス（売上ゼロ）を並べ、融資で黒字に見えないようにする。

## ケース設計のヒント

`cases/<id>/case.json` を複製して案件ごとに使う。設計時の参考：

- **開発体制**：週2回以上のMTG（例：月曜は優先順位・木曜はデモと検収）、毎週リリース、AIエージェント並行開発が有効。週1回は「間に合わない」感覚が出やすい。
- **競合優位性**：個別機能はAIで模倣されやすい時代のため、プラットフォーム（複数の利害関係者をつなぐ面）として展開する戦略を事業計画に明記する。
- **代表者略歴**：銀行向け別冊に含める。業歴30年超のアナログ現場知識＋AI習得速度が組み合わさる場合は「希少な強み」として記述すると審査で差が出る。

## 人間ゲート（外さない）

決算書数値の照合、価格・契約、法務・税務の確認、金融機関への提出。金利・制度条件は申込時に窓口で確認する。
