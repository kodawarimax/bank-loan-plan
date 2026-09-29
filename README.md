# bank-loan-plan

決算書3期分から、**借入余力・融資申込額・資金使途・返済計画・12か月の資金繰り表**を計算し、金融機関に相談できる事業計画の一式を作るスキルです（Claude Code / Codex 向け）。

- 数値はすべて `scripts/build_loan_plan.py` が計算します。文章やスライドに数字を手入力しません。
- 融資込みで黒字に見せないよう、営業収支（融資を除く）とストレス試算（売上ゼロ）を併記します。
- 借入余力が0（赤字・債務超過など）のときは生成を止めます。

## 使い方

```bash
export BUSINESS_PLAN_SCRIPTS=/path/to/build-business-plan/scripts
python3 scripts/build_loan_plan.py --self-test          # 架空の決算書で動作確認
python3 scripts/build_loan_plan.py --financials financials.json --out ./out --images
```

- `examples/sample-financials.json` … **架空**の決算書（動作確認用）
- `templates/financials.template.json` … 転記用の空テンプレート
- `cases/example/case.json` … 事業内容・資金使途・12か月計画の**架空のサンプル**。自社の内容に書き換えて使います。

## 依存

`plan_gate.py` と `render_package.py`（`build-business-plan` スキル）が必要です。このリポジトリには含まれていません。Python 3、`reportlab`、`python-pptx`、（画像化に）`pdftoppm` と LibreOffice を使います。

## 免責

サンプルの数値・会社・人物はすべて架空です。計算結果は計画値であり、融資の実行・金利・売上・利益を保証しません。金利や制度の条件は申込時に金融機関で確認してください。決算書の数値の照合、法務・税務の判断、金融機関への提出は、必ず人が行ってください。実在企業の決算書や生成物をこのリポジトリにコミットしないでください（`.gitignore` で一部を除外しています）。
