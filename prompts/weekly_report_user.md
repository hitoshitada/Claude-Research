以下の記事データを分析し、Weekly Intelligence Report のフロントセクション用の構造化JSONを生成してください。

技術カテゴリ: {category}
収集日: {collection_date}
記事数: {total_articles}件
対象国: {countries}

--- 記事データ ---
{articles_json}
--- 記事データここまで ---

以下のJSON形式で出力してください。JSONのみを出力し、他の文字は一切含めないでください：

{
  "report_metadata": {
    "category": "カテゴリ名",
    "date": "YYYY-MM-DD",
    "total_articles": 数値,
    "countries_count": 数値,
    "headline_keyword": "今週のキーワード（10文字以内、最もインパクトのある話題を凝縮）",
    "headline_sub": "キーワードの補足説明（30文字以内）",
    "accent_color": "#E8913A"
  },

  "key_metrics": [
    {
      "value": "表示値（数字または短い文字列）",
      "unit": "単位（件、カ国、%、W/mK等）",
      "label": "説明（8文字以内）"
    }
  ],

  "article_evaluations": [
    {
      "id": "#01",
      "title": "テーブル表示用の短縮タイトル（15文字以内）",
      "type": "種別（8文字以内。市場概観/新製品/学術論文/技術比較/市場危機/製品紹介/解説記事/カタログ/企業戦略 等）",
      "tech_novelty": 1,
      "proximity": 1,
      "market_impact": 1,
      "data_reliability": 1,
      "japan_relevance": 1,
      "one_line_summary": "一行サマリ（60文字以内。記事の核心を技術者向けに凝縮）",
      "is_highlight": false
    }
  ],

  "three_questions": [
    {
      "title": "問いかけ形式のタイトル（40文字以内。❶❷❸の番号を先頭に付ける）",
      "body": "問いの背景と判断材料（150文字以内。具体的な数値や企業名を含める）"
    }
  ],

  "opportunity_threat": [
    {
      "article_id": "#XX",
      "label": "表示ラベル（6文字以内。円の下に表示される見出し。内容が伝わる簡潔な名称にすること）",
      "quadrant": "TL",
      "x_position": 0.5,
      "y_position": 0.5,
      "opportunity": "機会（20文字以内。なければ空文字）",
      "threat": "脅威（20文字以内。なければ空文字）",
      "color": "gray"
    }
  ],

  "deep_dives": [
    {
      "article_id": "#XX",
      "section_title": "深掘りセクションのタイトル（30文字以内）",
      "source_line": "#XX | YYYY/MM/DD | ソース名",
      "score_line": "技術新規性●●○○○ 実用化距離●●●●● 市場インパクト●●●●● データ信頼性●●●○○ 日本関連度●●●●○",
      "body_paragraphs": [
        "記事の核心を技術者向けに要約した段落（200文字以内）",
        "補足情報や技術的詳細の段落（200文字以内、省略可）"
      ],
      "expert_comment": "【機会】と【脅威】の両面を含む技術者の視点コメント（400文字以内）",
      "expert_label": "技術者の視点"
    }
  ],

  "other_notable": [
    {
      "article_id": "#XX",
      "title": "記事タイトル（ソース名を括弧で付記）",
      "score_line": "主要な3軸のスコアを●○で表記",
      "comment": "技術者向けの一行コメント（100文字以内）"
    }
  ],

  "action_items": [
    {
      "timeframe": "即時（今週中）",
      "color": "red",
      "items": [
        "【部門名】具体的なアクション内容（100文字以内。何を、誰が、どうすべきかを明記）"
      ]
    },
    {
      "timeframe": "短期（1ヶ月）",
      "color": "orange",
      "items": []
    },
    {
      "timeframe": "中長期（四半期〜）",
      "color": "blue",
      "items": []
    }
  ]
}

注意事項：
- key_metrics は4件にすること（記事数、対象国数、および記事中の注目数値2件）
- article_evaluations は全記事分を含めること（記事IDの昇順）
- is_highlight は、5軸合計スコアが上位の記事、または特定の軸で突出している記事にtrueを設定
- three_questions は3件
- deep_dives は3件（記事が3件以下の場合は全件）
- other_notable は deep_dives に含まれない記事から注目度の高い順に3〜5件を選択
- score_line の●○表記は、スコア値の数だけ●を、残りを○で埋める（例：スコア3なら●●●○○）
- opportunity_threat は厳密に6〜8件に絞ること（多すぎるとマトリクスが読めなくなる）。5軸合計スコア上位の記事を優先し、類似テーマの記事はまとめて1項目にしてよい。BL象限（影響小）は最大2件まで
- opportunity_threat の配置ルール：
    - 同じ象限内の項目同士は x_position と y_position の差がそれぞれ0.25以上になるように分散させること
    - label は6文字以内に短くすること（円の直径に相当するスペースしかないため）
    - opportunity/threat のテキストも20文字以内に収めること
- accent_color はカテゴリに応じて設定：
    全固体電池="#1ABC9C", 接着・封止材="#E8913A",
    半導体パッケージング="#3498DB", 半導体PLP="#3498DB",
    EV材料="#27AE60", 高分子・樹脂="#8E44AD",
    機能性材料="#E67E22", ナノテクノロジー="#2ECC71",
    光通信・フォトニクス="#3498DB", 量子コンピュータ="#9B59B6",
    AI・機械学習="#1ABC9C", iPS細胞・再生医療="#E74C3C",
    細胞培養技術="#E74C3C"
