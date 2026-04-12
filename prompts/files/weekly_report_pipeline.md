# Weekly Intelligence Report 生成パイプライン設計書

## 全体アーキテクチャ

```
[Step 1] Gemini API (Search Grounding)
    → 記事収集・翻訳・個別記事PDF生成（現行システム）
    → 出力: articles_data.json（全記事の構造化データ）

[Step 2] Gemini API (分析・構成)  ← ★今回のプロンプト
    → articles_data.json を入力として
    → フロントレポート用の構造化JSON（front_report.json）を生成

[Step 3] Python スクリプト (reportlab)
    → front_report.json を入力として
    → フロントレポートPDF を生成

[Step 4] Python (pypdf)
    → フロントレポートPDF + 全記事PDF を連結
    → 最終成果物PDF
```

---

## Step 1 → Step 2 の中間データ形式

Step 1（現行システム）の出力を以下のJSON形式に整形してStep 2に渡します。
現行のPDF出力と並行して、このJSONも生成するようにします。

```json
{
  "category": "接着・封止材",
  "collection_date": "2026-04-04",
  "total_articles": 12,
  "countries": ["韓国", "アメリカ", "イギリス", "ドイツ", "ベトナム"],
  "articles": [
    {
      "id": "#01",
      "title": "スマート接着剤が牽引する未来の接着・封止材市場の展望",
      "source": "Insights",
      "country": "韓国",
      "date": "2026-04-01",
      "url": "https://...",
      "summary": "2026年のテープおよび接着剤分野における革新は...",
      "details": "背景と市場の変革\n2026年、接着・封止材の市場は..."
    }
  ]
}
```

---

## Step 2: Gemini プロンプト

以下がGemini APIに渡すプロンプトです。

### System Prompt

```
あなたは、日本の技術系メディアサイト「troy-technical.jp」の
シニアテクニカルアナリストです。

材料工学、特に接着剤・封止材・半導体パッケージング材料・
電池材料に深い専門知識を持ち、日本の製造業の技術者や
新製品企画担当者に向けて、海外技術情報のキュレーションと
分析を行っています。

あなたの役割は、収集された技術記事群を分析し、
Weekly Intelligence Report のフロントセクションを構成する
構造化データを生成することです。

## 出力ルール

1. 必ず指定されたJSON形式で出力すること
2. JSONのみを出力し、前後に説明文やマークダウンの```を付けないこと
3. 日本語で記述すること
4. 各記事を以下の5軸で1〜5の整数で評価すること：
   - tech_novelty（技術新規性）：
     5=学術的ブレークスルー/新規メカニズムの発見
     4=既存技術の大幅な性能向上/新カテゴリの創出
     3=応用範囲の拡大/新しい組み合わせ
     2=既存製品の改良/ラインナップ拡充
     1=既知トレンドの概観/市場レポート
   - proximity（実用化距離）：
     5=既に製品化/今すぐ調達・評価可能
     4=製品発表済み/サンプル入手可能段階
     3=パイロット段階/2-3年以内に製品化見込み
     2=応用研究段階/3-5年先
     1=基礎研究/5年以上先
   - market_impact（市場インパクト）：
     5=業界全体のサプライチェーンや設計思想に影響
     4=主要セグメント（EV、半導体等）に大きな影響
     3=特定の技術分野に中程度の影響
     2=ニッチ市場への影響
     1=学術的関心が主で市場への直接影響は限定的
   - data_reliability（データ信頼性）：
     5=査読付き論文/具体的定量データあり
     4=企業技術報告/具体的スペック開示あり
     3=業界レポート/一部定量データあり
     2=プレスリリース/定性的情報が主
     1=トレンド記事/情報源が不明確
   - japan_relevance（日本関連度）：
     5=日本企業・日本市場に直接影響/日本発の研究
     4=日本のサプライチェーンに波及する可能性が高い
     3=日本企業が関与する技術分野に関連
     2=グローバル動向として間接的に関連
     1=日本との直接的な関連は薄い

5. 「3つの問い」は、読者に判断を迫る問いかけ形式にすること。
   一般的なまとめではなく「あなたの会社は大丈夫か？」という
   緊急性や具体性を持たせること。

6. 深掘り記事は、5軸評価の合計スコアが高い上位3件を選ぶこと。
   ただし、技術新規性が高い記事と実用化距離が近い記事の
   バランスを取ること。

7. 機会/脅威の分析は、日本の以下の立場から評価すること：
   - 材料・素材メーカー
   - セルメーカー/OEM
   - 部品メーカー
   - 調達・購買部門

8. 「技術者の視点」コメントでは、以下を含めること：
   - 論文・製品発表の数値の妥当性評価
   - 実用化に向けた未解決課題の指摘
   - 日本企業にとっての機会と脅威の両面分析
   - 具体的な次のアクション提案

9. アクション提案は、即時/短期/中長期の3段階で、
   具体的な部門名と行動を明記すること。
```

### User Prompt テンプレート

```
以下の記事データを分析し、Weekly Intelligence Report の
フロントセクション用の構造化JSONを生成してください。

技術カテゴリ: {category}
収集日: {collection_date}
記事数: {total_articles}件
対象国: {countries}

--- 記事データ ---
{articles_json}
--- 記事データここまで ---

以下のJSON形式で出力してください：

{
  "report_metadata": {
    "category": "カテゴリ名",
    "date": "YYYY-MM-DD",
    "total_articles": 数値,
    "countries_count": 数値,
    "headline_keyword": "今週のキーワード（10文字以内）",
    "headline_sub": "キーワードの補足（30文字以内）"
  },

  "key_metrics": [
    {"value": "表示値", "unit": "単位", "label": "説明（8文字以内）"}
  ],

  "article_evaluations": [
    {
      "id": "#01",
      "title": "テーブル表示用の短縮タイトル（15文字以内）",
      "type": "種別（市場概観/新製品/学術論文/技術比較/市場危機/製品紹介/解説記事/カタログ等、8文字以内）",
      "tech_novelty": 1-5,
      "proximity": 1-5,
      "market_impact": 1-5,
      "data_reliability": 1-5,
      "japan_relevance": 1-5,
      "one_line_summary": "一行サマリ（60文字以内）",
      "is_highlight": true/false
    }
  ],

  "three_questions": [
    {
      "title": "問いかけ形式のタイトル（40文字以内）",
      "body": "問いの背景と判断材料（150文字以内）"
    }
  ],

  "opportunity_threat": [
    {
      "article_id": "#XX",
      "label": "マトリクス表示用ラベル（15文字以内）",
      "quadrant": "TL/TR/BL/BR",
      "x_position": 0.0-1.0,
      "y_position": 0.0-1.0,
      "opportunity": "機会の説明（30文字以内、なければ空文字）",
      "threat": "脅威の説明（30文字以内、なければ空文字）",
      "color": "red/orange/teal/green/gray"
    }
  ],

  "deep_dives": [
    {
      "article_id": "#XX",
      "section_title": "深掘りセクションのタイトル（30文字以内）",
      "source_line": "記事メタ情報（#XX | 日付 | ソース名）",
      "score_line": "5軸スコアの表示用文字列",
      "body_paragraphs": [
        "本文段落1（200文字以内）",
        "本文段落2（200文字以内、任意）"
      ],
      "expert_comment": "技術者の視点コメント（400文字以内）",
      "expert_label": "コメントボックスのラベル（デフォルト：技術者の視点）"
    }
  ],

  "other_notable": [
    {
      "article_id": "#XX",
      "title": "記事タイトル（ソース名を括弧で付記）",
      "score_line": "主要スコアの表示用文字列",
      "comment": "コメント（100文字以内）"
    }
  ],

  "action_items": [
    {
      "timeframe": "即時（今週中）/ 短期（1ヶ月）/ 中長期（四半期〜）",
      "color": "red/orange/blue",
      "items": [
        "【部門名】具体的なアクション（100文字以内）"
      ]
    }
  ]
}
```

---

## Step 3: Python PDF生成スクリプト

Step 2の出力JSON（front_report.json）を読み込み、
reportlabでPDFを生成するスクリプト。
build_adhesive_v4.py をベースに、JSONからデータを
読み込む形に汎用化したものを使用。

主な変更点：
- ハードコードされた記事データをJSON読み込みに変更
- カテゴリ名、日付、色テーマをJSONから動的に設定
- OTマトリクスの配置をJSONのx/y座標から描画

## Step 4: PDF連結

```python
from pypdf import PdfWriter, PdfReader

writer = PdfWriter()

# フロントレポート
front = PdfReader("front_report.pdf")
for page in front.pages:
    writer.add_page(page)

# 全記事PDF（既存の出力）
articles = PdfReader("articles_full.pdf")
for page in articles.pages:
    writer.add_page(page)

with open("weekly_report_final.pdf", "wb") as f:
    writer.write(f)
```

---

## カテゴリ別のカラーテーマ設定

各技術カテゴリごとにアクセントカラーを変えることで
視覚的に区別できるようにする。

| カテゴリ | アクセント色 | ヘッダーバー色 |
|---------|------------|-------------|
| 全固体電池 | TEAL (#1ABC9C) | TEAL |
| 接着・封止材 | ORANGE (#E8913A) | ORANGE |
| 半導体パッケージング | BLUE (#3498DB) | BLUE |
| EV材料 | GREEN (#27AE60) | GREEN |
| 医療材料 | PURPLE (#8E44AD) | PURPLE |

JSONに `"accent_color": "#E8913A"` を追加して
スクリプト側で動的に適用する。

---

## 運用フロー（まとめ）

1. 毎週金曜：Step 1 で記事収集 → articles_data.json + articles_full.pdf
2. articles_data.json を Step 2 の Gemini プロンプトに投入
3. Gemini の出力（front_report.json）を Step 3 で PDF 化
4. Step 4 で連結 → 最終成果物を WordPress にアップロード

所要時間目安：
- Step 1: 既存パイプライン（自動化済み想定）
- Step 2: Gemini API 1回の呼び出し（30秒〜1分）
- Step 3: Python実行（5秒以内）
- Step 4: Python実行（2秒以内）
