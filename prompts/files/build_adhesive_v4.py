#!/usr/bin/env python3
"""
接着・封止材 ウィークリーレポート v4
2026年04月04日 | 12件
- 冒頭に記事評価マトリクス（技術新規性/実用化距離/市場インパクト/データ信頼性/日本関連度）
- 機会 vs 脅威マトリクス
- 深掘り3件
- アクション提案
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus.flowables import Flowable

pdfmetrics.registerFont(TTFont("JP", "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"))
pdfmetrics.registerFont(TTFont("JPP", "/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf"))

# Colors
NAVY = HexColor("#1B2A4A")
DBLUE = HexColor("#2C3E6B")
ACC = HexColor("#E8913A")
ACC_LT = HexColor("#FDF0E2")
LG = HexColor("#F5F6F8")
MG = HexColor("#8A8FA0")
TXT = HexColor("#3A3A3A")
W = HexColor("#FFFFFF")
RED = HexColor("#C0392B")
GRN = HexColor("#27AE60")
TEAL = HexColor("#1ABC9C")

PW, PH = A4
ML=18*mm; MR=18*mm; MT=20*mm; MB=18*mm
CW = PW - ML - MR

# Styles
S = {}
S['body'] = ParagraphStyle('body', fontName='JP', fontSize=9, textColor=TXT,
                            leading=16, alignment=TA_JUSTIFY, spaceAfter=2*mm)
S['bs'] = ParagraphStyle('bs', fontName='JP', fontSize=8, textColor=MG, leading=13, spaceAfter=1*mm)
S['exp'] = ParagraphStyle('exp', fontName='JP', fontSize=9, textColor=DBLUE,
                           leading=15, leftIndent=8*mm, rightIndent=4*mm, spaceBefore=2*mm, spaceAfter=3*mm)
S['qt'] = ParagraphStyle('qt', fontName='JP', fontSize=10, textColor=NAVY,
                          leading=16, spaceBefore=3*mm, spaceAfter=1*mm, leftIndent=6*mm)
S['qb'] = ParagraphStyle('qb', fontName='JP', fontSize=8.5, textColor=TXT,
                          leading=14, spaceAfter=3*mm, leftIndent=6*mm)
S['sub'] = ParagraphStyle('sub', fontName='JP', fontSize=11, textColor=DBLUE,
                           leading=16, spaceBefore=5*mm, spaceAfter=2*mm)
S['act'] = ParagraphStyle('act', fontName='JP', fontSize=9, textColor=TXT,
                           leading=15, spaceAfter=2*mm, leftIndent=8*mm)
S['foot'] = ParagraphStyle('foot', fontName='JPP', fontSize=7, textColor=MG, alignment=TA_CENTER)
S['tc'] = ParagraphStyle('tc', fontName='JP', fontSize=6.5, textColor=TXT, leading=10)
S['tcb'] = ParagraphStyle('tcb', fontName='JP', fontSize=6.5, textColor=NAVY, leading=10)


# ============================================================
# Flowables
# ============================================================

class Header(Flowable):
    def __init__(self, w, h=38*mm):
        Flowable.__init__(self); self.width=w; self.height=h
    def draw(self):
        c=self.canv
        c.setFillColor(NAVY); c.rect(-ML,-4*mm,PW,self.height+4*mm,fill=1,stroke=0)
        c.setFillColor(ACC); c.rect(-ML,-4*mm,4*mm,self.height+4*mm,fill=1,stroke=0)
        c.setFont("JP",22); c.setFillColor(W)
        c.drawString(4*mm, self.height-11*mm, "接着・封止材")
        c.setFont("JPP",13)
        c.drawString(4*mm, self.height-19*mm, "Weekly Intelligence Report")
        c.setFont("JP",9); c.setFillColor(HexColor("#B0BEC5"))
        c.drawString(4*mm, self.height-27*mm, "2026年04月04日 | 12件 | 5カ国")
        c.setFont("JPP",7)
        c.drawString(4*mm, self.height-32*mm, "troy-technical.jp")
        rx=self.width-58*mm
        c.setFont("JP",10); c.setFillColor(ACC)
        c.drawString(rx, self.height-11*mm, "今週のキーワード")
        c.setFont("JP",14); c.setFillColor(W)
        c.drawString(rx, self.height-20*mm, "ボンドショック")
        c.setFont("JP",14); c.setFillColor(ACC)
        c.drawString(rx+52*mm, self.height-20*mm, "& TIM革新")
        c.setFont("JPP",7); c.setFillColor(HexColor("#B0BEC5"))
        c.drawString(rx, self.height-26*mm, "韓国で接着剤供給危機 / 液体金属TIM 57.4 W/mK")

class Div(Flowable):
    def __init__(self, w, text, color=NAVY):
        Flowable.__init__(self); self.width=w; self.height=8*mm
        self.text=text; self.color=color
    def draw(self):
        c=self.canv
        c.setFillColor(self.color); c.rect(0,0,3*mm,self.height,fill=1,stroke=0)
        c.setFillColor(LG); c.rect(4*mm,0,self.width-4*mm,self.height,fill=1,stroke=0)
        c.setFont("JP",11); c.setFillColor(self.color)
        c.drawString(7*mm, 2.2*mm, self.text)

class ExpBox(Flowable):
    def __init__(self, w, text, label="技術者の視点"):
        Flowable.__init__(self); self.width=w; self.text=text; self.label=label
        self._p = Paragraph(text, S['exp'])
        _, self._ph = self._p.wrap(w-14*mm, 500)
        self.height = self._ph + 12*mm
    def draw(self):
        c=self.canv
        c.setFillColor(ACC_LT); c.roundRect(0,0,self.width,self.height,2*mm,fill=1,stroke=0)
        c.setFillColor(ACC); c.rect(0,0,3*mm,self.height,fill=1,stroke=0)
        c.setFont("JP",7.5); c.setFillColor(ACC)
        c.drawString(6*mm, self.height-5*mm, f"▶ {self.label}")
        self._p.drawOn(c, 5*mm, 2*mm)

class Metrics(Flowable):
    def __init__(self, w, ms):
        Flowable.__init__(self); self.width=w; self.ms=ms; self.height=18*mm
    def draw(self):
        c=self.canv; n=len(self.ms); cw=(self.width-(n-1)*2.5*mm)/n
        for i,(v,u,l) in enumerate(self.ms):
            x=i*(cw+2.5*mm)
            c.setFillColor(LG); c.roundRect(x,0,cw,self.height,2*mm,fill=1,stroke=0)
            c.setFont("JP",14); c.setFillColor(ACC)
            c.drawCentredString(x+cw/2, 9*mm, v)
            c.setFont("JP",6); c.setFillColor(MG)
            c.drawCentredString(x+cw/2, 5.5*mm, u)
            c.setFont("JP",6); c.setFillColor(TXT)
            c.drawCentredString(x+cw/2, 1.5*mm, l)


# ============================================================
# Article Evaluation Matrix (content-based axes)
# ============================================================

def score_dots(n, max_n=5):
    """Render score as filled/empty circles using text"""
    filled = "●" * n
    empty = "○" * (max_n - n)
    if n >= 4:
        color = "#C0392B"   # red = high
    elif n >= 3:
        color = "#E8913A"   # orange
    elif n >= 2:
        color = "#8A8FA0"   # gray
    else:
        color = "#C0C4CE"   # light
    return Paragraph(
        f'<font color="{color}" size="7">{filled}</font>'
        f'<font color="#D0D4DE" size="7">{empty}</font>',
        ParagraphStyle('dots', fontName='JP', fontSize=7, alignment=TA_CENTER, leading=10))


def build_triage():
    hs = ParagraphStyle('th', fontName='JP', fontSize=5.5, textColor=W, leading=8, alignment=TA_CENTER)
    hl = ParagraphStyle('thl', fontName='JP', fontSize=5.5, textColor=W, leading=8)

    header = [
        Paragraph('#', hs),
        Paragraph('記事タイトル', hl),
        Paragraph('種別', hs),
        Paragraph('技術<br/>新規性', hs),
        Paragraph('実用化<br/>距離', hs),
        Paragraph('市場<br/>インパクト', hs),
        Paragraph('データ<br/>信頼性', hs),
        Paragraph('日本<br/>関連度', hs),
        Paragraph('一行サマリ', hl),
    ]

    # (num, title, type, tech_novelty, proximity, market_impact, data_reliability, japan_relevance, summary)
    # tech_novelty: 5=breakthrough, 1=既知トレンド
    # proximity: 5=今すぐ使える製品, 1=基礎研究10年先
    # market_impact: 5=業界全体, 1=ニッチ
    # data_reliability: 5=査読論文+定量, 1=定性的プレスリリース
    # japan_relevance: 5=日本市場直結, 1=関連薄い
    articles = [
        ("#01", "スマート接着剤市場展望", "市場概観",
         1, 2, 3, 1, 2,
         "自己修復・ナノ材料統合のトレンド概観。新規性低く定性的。"),
        ("#02", "ボンドショック（韓国）", "市場危機",
         1, 5, 5, 3, 4,
         "接着剤の価格高騰・供給不足。自動車・半導体で生産停止。日本波及リスクあり。"),
        ("#03", "環境配慮型接着剤", "製品紹介",
         2, 4, 2, 2, 2,
         "Master Bond社。溶剤フリー・RoHS準拠。規制対応の参考に。"),
        ("#04", "構造用接着剤市場2035年", "市場予測",
         1, 3, 4, 3, 3,
         "EV・航空宇宙の異種材接合需要で2035年まで成長。debonding技術にも言及。"),
        ("#05", "パワーモジュールTIM 5選", "技術比較",
         3, 3, 3, 3, 3,
         "相変化/CNT/マイクロチャネル比較。熱抵抗15〜25%削減。設計指針として有用。"),
        ("#06", "UV17Med 医療用UV接着剤", "新製品",
         3, 4, 2, 3, 2,
         "TPU向けUV硬化型。ISO 10993-5合格。医療用ニッチ。"),
        ("#07", "EV Protect 4006", "新製品",
         4, 4, 4, 3, 4,
         "EV電池用難燃PUフォーム。難燃+軽量+NVH統合。封止材の新カテゴリ。"),
        ("#08", "ActiveCopper 銅焼結", "製品紹介",
         3, 4, 3, 2, 3,
         "MacDermid Alpha。ダイアタッチ向け銅焼結。ウェハーレベルPKG対応。"),
        ("#09", "HALA Contec 2026カタログ", "カタログ",
         2, 4, 2, 2, 2,
         "ギャップフィラーTGF-BXS-SI（1.2W/mK）。製品カタログ更新。"),
        ("#10", "液体金属TIM 57.4 W/mK", "学術論文",
         5, 1, 4, 5, 2,
         "SAM界面制御でLM/AlN複合TIM。チップ温度50.8%低減。査読論文。"),
        ("#11", "UV硬化システム産業応用", "解説記事",
         1, 3, 2, 1, 1,
         "医療・光学・自動車でのUV硬化の概観。新規性低い。"),
        ("#12", "可視光ポリオレフィングラフト", "学術論文",
         5, 1, 3, 5, 2,
         "触媒フリーPE/PP機能化。ホットメルト超えせん断強度。査読論文。"),
    ]

    rows = [header]
    for a in articles:
        num, title, atype, tn, px, mi, dr, jr, summary = a
        # Calculate composite score for highlight
        total = tn + px + mi + dr + jr
        rows.append([
            Paragraph(f'<b>{num}</b>', S['tcb']),
            Paragraph(title, S['tcb']),
            Paragraph(atype, ParagraphStyle('at', fontName='JP', fontSize=6, textColor=MG, leading=9, alignment=TA_CENTER)),
            score_dots(tn, 5),
            score_dots(px, 5),
            score_dots(mi, 5),
            score_dots(dr, 5),
            score_dots(jr, 5),
            Paragraph(summary, S['tc']),
        ])

    col_w = [8*mm, 26*mm, 12*mm, 14*mm, 14*mm, 14*mm, 14*mm, 14*mm, CW-116*mm]
    t = Table(rows, colWidths=col_w, repeatRows=1)

    style_cmds = [
        ('BACKGROUND', (0,0), (-1,0), NAVY),
        ('TEXTCOLOR', (0,0), (-1,0), W),
        ('BACKGROUND', (0,1), (-1,-1), W),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [W, LG]),
        ('GRID', (0,0), (-1,-1), 0.3, HexColor("#D0D4DE")),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 1.2*mm),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1.2*mm),
        ('LEFTPADDING', (0,0), (-1,-1), 1*mm),
        ('RIGHTPADDING', (0,0), (-1,-1), 1*mm),
        ('ALIGN', (2,0), (7,-1), 'CENTER'),
    ]
    # Highlight top articles: #02(row2), #07(row7), #10(row10), #12(row12)
    for r in [2, 7, 10, 12]:
        style_cmds.append(('BACKGROUND', (0, r), (-1, r), HexColor("#FFF8E1")))

    t.setStyle(TableStyle(style_cmds))
    return t


def build_legend():
    return Paragraph(
        '<font color="#C0392B" size="7">●●●●</font><font color="#D0D4DE" size="7">○</font> 高　'
        '<font color="#E8913A" size="7">●●●</font><font color="#D0D4DE" size="7">○○</font> 中高　'
        '<font color="#8A8FA0" size="7">●●</font><font color="#D0D4DE" size="7">○○○</font> 中　'
        '<font color="#C0C4CE" size="7">●</font><font color="#D0D4DE" size="7">○○○○</font> 低　'
        '| 背景黄色＝注目記事',
        ParagraphStyle('lg', fontName='JP', fontSize=7, textColor=MG, leading=11))


def build_reading_guide():
    """Short reading guide below the table"""
    return Paragraph(
        "各列の見方 ― "
        "<b>技術新規性</b>：ブレークスルー度合い（学術論文＞製品改良＞トレンド概観）　"
        "<b>実用化距離</b>：製品として使える近さ（新製品発表＞応用開発＞基礎研究）　"
        "<b>市場インパクト</b>：業界全体への影響規模　"
        "<b>データ信頼性</b>：定量データ・査読の有無　"
        "<b>日本関連度</b>：日本の企業・サプライチェーンとの直接的関連性",
        ParagraphStyle('rg', fontName='JP', fontSize=7, textColor=TXT, leading=11, spaceAfter=2*mm))


# ============================================================
# OT Matrix
# ============================================================

class OTMatrix(Flowable):
    def __init__(self, w):
        Flowable.__init__(self); self.width=w; self.height=95*mm
    def draw(self):
        c=self.canv
        c.setFillColor(LG); c.roundRect(0,0,self.width,self.height,2*mm,fill=1,stroke=0)
        c.setFont("JP",9); c.setFillColor(NAVY)
        c.drawString(4*mm, self.height-6*mm, "日本企業にとっての「機会 vs 脅威」マトリクス")

        mx=10*mm; my=8*mm; mw=self.width-20*mm; mh=self.height-22*mm
        hw=mw/2; hh=mh/2

        c.setFillColor(HexColor("#E8F5E9")); c.rect(mx,my+hh,hw,hh,fill=1,stroke=0)
        c.setFillColor(HexColor("#FFF8E1")); c.rect(mx+hw,my+hh,hw,hh,fill=1,stroke=0)
        c.setFillColor(W); c.rect(mx,my,hw,hh,fill=1,stroke=0)
        c.setFillColor(HexColor("#FFEBEE")); c.rect(mx+hw,my,hw,hh,fill=1,stroke=0)

        c.setStrokeColor(HexColor("#B0BEC5")); c.setLineWidth(1)
        c.line(mx,my+hh,mx+mw,my+hh); c.line(mx+hw,my,mx+hw,my+mh)
        c.setLineWidth(0.5); c.rect(mx,my,mw,mh,fill=0,stroke=1)

        c.setFont("JP",7); c.setFillColor(GRN)
        c.drawString(mx+2*mm, my+hh+hh-4*mm, "機会 大・脅威 小")
        c.setFillColor(ACC)
        c.drawString(mx+hw+2*mm, my+hh+hh-4*mm, "機会 大・脅威 大")
        c.setFillColor(MG)
        c.drawString(mx+2*mm, my+hh-4*mm, "影響小（参考）")
        c.setFillColor(RED)
        c.drawString(mx+hw+2*mm, my+hh-4*mm, "脅威 大・機会 小")

        c.setFont("JP",7); c.setFillColor(MG)
        c.saveState(); c.translate(mx-5*mm, my+mh/2)
        c.rotate(90); c.drawCentredString(0,0,"← 機会 →"); c.restoreState()
        c.drawCentredString(mx+mw/2, my-4*mm, "← 脅威 →")

        items = [
            (0.3,0.7,"#02 ボンドショック","韓国市場供給機会","原材料共通で日本波及リスク",RED,"TR"),
            (0.55,0.35,"#04 構造材市場2035","EV/航空宇宙の長期需要","海外勢との競合激化",ACC,"TR"),
            (0.4,0.7,"#07 EV Protect","封止材の新カテゴリ創出","",TEAL,"TL"),
            (0.6,0.35,"#10 液体金属TIM","次世代TIM市場","",TEAL,"TL"),
            (0.3,0.2,"#06 UV17Med","医療用ニッチ","",GRN,"TL"),
            (0.3,0.7,"#01 スマート接着剤","","",MG,"BL"),
            (0.6,0.5,"#03 環境配慮型","","",MG,"BL"),
            (0.3,0.3,"#12 可視光グラフト","","",MG,"BL"),
            (0.6,0.2,"#09 HALA Contec","","",MG,"BL"),
            (0.4,0.5,"#08 ActiveCopper","","海外勢の先端技術が脅威",HexColor("#E57373"),"BR"),
        ]
        for xf,yf,label,opp,thr,color,quad in items:
            bx = mx + (hw if quad[1]=='R' else 0)
            by = my + (hh if quad[0]=='T' else 0)
            px=bx+xf*hw; py=by+yf*hh
            r = 2.5*mm if (opp or thr) else 1.8*mm
            c.setFillColor(color); c.circle(px,py,r,fill=1,stroke=0)
            c.setFont("JP",5.5 if (opp or thr) else 5)
            c.setFillColor(color if (opp or thr) else MG)
            c.drawString(px+r+1*mm, py-1*mm, label)
            if opp:
                c.setFont("JP",5); c.setFillColor(GRN)
                c.drawString(px+r+1*mm, py-5*mm, f"↑ {opp}")
            if thr:
                c.setFont("JP",5); c.setFillColor(RED)
                c.drawString(px+r+1*mm, py-(9*mm if opp else 5*mm), f"↓ {thr}")


# ============================================================
# Story
# ============================================================

def build_story():
    story = []

    # === P1: Header + Evaluation Matrix ===
    story.append(Header(CW))
    story.append(Spacer(1,4*mm))
    story.append(Metrics(CW, [
        ("12","件","収集記事"), ("5","カ国","情報ソース"),
        ("57.4","W/mK","液体金属TIM"), ("15-25%","削減","TIM置換効果"),
    ]))
    story.append(Spacer(1,4*mm))

    story.append(Div(CW, "今週の全12記事 ― 5軸評価で読むべき記事を選ぶ"))
    story.append(Spacer(1,1*mm))
    story.append(build_reading_guide())
    story.append(build_triage())
    story.append(Spacer(1,1*mm))
    story.append(build_legend())

    story.append(PageBreak())

    # === P2: 3 Questions + OT Matrix ===
    story.append(Div(CW, "今週、判断に影響しうる3つの問い"))
    story.append(Spacer(1,3*mm))

    for qt, qb in [
        ("❶ 韓国「ボンドショック」は日本の調達にも波及するか？",
         "韓国で接着剤の価格高騰・供給不足が深刻化。自動車・半導体で生産停止も。"
         "エポキシ・イソシアネート系原材料は日韓で供給元が重なり、波及リスクは無視できません。"),
        ("❷ 液体金属TIM 57.4 W/mK ― 熱設計の前提が変わるか？",
         "SAM界面制御でLM/AlN複合TIMが従来品の5〜10倍の熱伝導率を達成。"
         "ただし10wt%充填のラボスケール。ポンプアウト耐性・量産コストは未検証。"),
        ("❸ EV Protect 4006 ― バッテリー封止材に新カテゴリか？",
         "難燃＋軽量＋NVHを1材料で実現。従来の二重構成を置き換えうる。"
         "CTP設計の潮流と相性がよく、組立工程の簡素化に直結。"),
    ]:
        story.append(Paragraph(qt, S['qt']))
        story.append(Paragraph(qb, S['qb']))

    story.append(Spacer(1,2*mm))
    story.append(Div(CW, "日本企業にとっての「機会 vs 脅威」", ACC))
    story.append(Spacer(1,3*mm))
    story.append(OTMatrix(CW))
    story.append(Spacer(1,1*mm))
    story.append(Paragraph(
        '<font color="#8A8FA0">大丸＝機会/脅威の詳細あり、小丸＝参考。</font>', S['bs']))

    story.append(PageBreak())

    # === P3: Deep Dives 1 & 2 ===
    story.append(Div(CW, "深掘り ① ― 韓国「ボンドショック」と日本への示唆"))
    story.append(Spacer(1,2*mm))
    story.append(Paragraph('<font color="#8A8FA0">#02 | 2026/04/02 | 文化日報 | '
        '技術新規性●○○○○ 実用化距離●●●●●  市場インパクト●●●●● データ信頼性●●●○○ 日本関連度●●●●○</font>', S['bs']))
    story.append(Paragraph(
        "韓国産業界が「ボンドショック」と呼ばれる深刻な接着剤供給危機に直面。"
        "国際的な原材料価格高騰と供給不足が原因で、自動車・半導体で生産停止事例が発生。"
        "特に構造用接着剤の不足は、EV化に伴う異種材接合需要の急拡大と重なり深刻です。"
        "韓国政府（中小ベンチャー企業部）も不公平な取引慣行の調査に着手しました。", S['body']))

    story.append(ExpBox(CW,
        "【機会】韓国市場への日本製接着剤の短期的供給機会。構造用エポキシ・PU系で日本メーカーは技術優位。"
        "【脅威】原材料（BASF、ヘンケル、ダウ等）は日韓共通のため、同様の価格圧力が時間差で波及するリスク。"
        "エポキシ樹脂、MDI/TDI系イソシアネートの調達が重複している場合は要注意。"
        "短期：在庫水準確認。中長期：バイオベース代替・調達先多角化を検討。"
    ))

    story.append(Spacer(1,4*mm))

    story.append(Div(CW, "深掘り ② ― 液体金属TIM 57.4 W/mK の衝撃と冷静な評価"))
    story.append(Spacer(1,2*mm))
    story.append(Paragraph('<font color="#8A8FA0">#10 | 2026/04/01 | ACS Chem. Mater. | '
        '技術新規性●●●●● 実用化距離●○○○○ 市場インパクト●●●●○ データ信頼性●●●●● 日本関連度●●○○○</font>', S['bs']))
    story.append(Paragraph(
        "SAMによる界面工学でLM/AlN複合TIMが57.4 W/m·Kの熱伝導率と0.122 K·cm2/Wの有効熱抵抗を達成。"
        "40nm AlN、10wt%充填、シランカップリング剤修飾が最適条件。"
        "熱電発電機で出力電圧1.45倍増、チップ冷却で過剰温度50.8%低減を実証。", S['body']))

    story.append(ExpBox(CW,
        "57.4 W/mKは現行シリコーンTIM（5〜15 W/mK）の5〜10倍。熱設計パラダイムを変えうる数値。"
        "ただし：①ガリウム系LMはAl基材への腐食性が課題 ②ラボスケールで量産コスト未知 "
        "③ポンプアウト耐性データが限定的。"
        "【注目点】SAM界面制御アプローチ自体は汎用性が高く、他のフィラー系にも応用可能な方法論。"
        "自社TIM開発の性能向上手法として参照する価値あり。"
    ))

    story.append(PageBreak())

    # === P4: Deep Dive 3 + Notable Others ===
    story.append(Div(CW, "深掘り ③ ― EV Protect 4006：封止材の新カテゴリ"))
    story.append(Spacer(1,2*mm))
    story.append(Paragraph('<font color="#8A8FA0">#07 | 2026/04/04 | H.B. Fuller | '
        '技術新規性●●●●○ 実用化距離●●●●○ 市場インパクト●●●●○ データ信頼性●●●○○ 日本関連度●●●●○</font>', S['bs']))
    story.append(Paragraph(
        "H.B.フラーのEV Protect 4006 SFRは、EVバッテリーモジュール向けの"
        "液状塗布型2液性難燃性低密度PUフォーム材。セル間の熱伝播防止、低密度による軽量化、"
        "半構造的特性によるNVH吸収を1つの材料で実現する点が特徴的です。", S['body']))

    story.append(ExpBox(CW,
        "【機会】「難燃＋軽量＋NVH」の三位一体は、従来のシリコーンフォーム＋ギャップフィラーの"
        "二重構成を置き換える新カテゴリ。CTP設計と相性良く、パック組立の簡素化に貢献。"
        "日本のバッテリーメーカー/OEMにとって評価対象に入れるべき製品。"
        "【脅威】H.B.フラーはグローバル販路を持ち、日本の封止材メーカーにとっては"
        "EV封止材市場での直接的な競合製品。耐熱温度・長期信頼性での技術的差別化が必要。"
    ))

    story.append(Spacer(1,4*mm))

    story.append(Div(CW, "その他の注目記事"))
    story.append(Spacer(1,2*mm))

    for title, scores, comment in [
        ("#05 パワーモジュールTIM技術5選（Patsnap）",
         "技術新規性●●●○○ 実用化距離●●●○○ 市場インパクト●●●○○",
         "相変化/CNT/マイクロチャネルの3アーキテクチャを比較。熱グリース置換で熱抵抗15〜25%削減。"
         "パワエレ/半導体の熱設計者向け。具体的な材料選定の判断材料として有用。"),
        ("#08 ActiveCopper 銅焼結ペースト（MacDermid Alpha）",
         "技術新規性●●●○○ 実用化距離●●●●○ 市場インパクト●●●○○",
         "銅焼結はダイアタッチの主流技術化が進行中。ウェハーレベルPKG/ハイブリッドボンディング対応。"
         "日本の半導体PKG材料メーカーにとっては海外勢への対抗策が必要な領域。"),
        ("#12 可視光ポリオレフィングラフト重合（ACS JACS）",
         "技術新規性●●●●● 実用化距離●○○○○ データ信頼性●●●●●",
         "触媒フリーでPE/PP機能化。せん断強度が市販ホットメルトを桁違いに超過。"
         "リサイクル材にも適用可能。実用化は5年以上先だが、接着の基本概念を変える研究。"),
    ]:
        story.append(Paragraph(f'<b>{title}</b>', S['tcb']))
        story.append(Paragraph(f'<font color="#8A8FA0">{scores}</font>', S['bs']))
        story.append(Paragraph(comment, S['body']))

    story.append(PageBreak())

    # === P5: Actions ===
    story.append(Div(CW, "今週のアクション提案", ACC))
    story.append(Spacer(1,4*mm))

    story.append(Paragraph(
        "記事評価マトリクスと機会/脅威分析を踏まえたアクション提案です。", S['body']))

    for period, color, items in [
        ("即時（今週中）", RED, [
            "【調達確認】ボンドショック波及リスク：韓国系サプライヤーからのエポキシ・PU系"
            "接着剤調達がある場合、在庫水準と代替ソースを確認。",
            "【技術共有】液体金属TIM論文（DOI: 10.1021/acsaenm.6c00043）を"
            "熱設計チームへ転送。次世代TIMベンチマークとして評価検討。",
            "【競合分析】H.B.フラー EV Protect 4006 SFRのスペックシート入手。"
            "自社バッテリー封止材との比較着手。",
        ]),
        ("短期（1ヶ月）", ACC, [
            "【R&D検討】可視光グラフト重合技術（DOI: 10.1021/jacs.5c21265）について"
            "自社のポリオレフィン接着課題との適合性を評価。",
            "【半導体PKG】ActiveCopper銅焼結の技術仕様入手。"
            "自社ダイアタッチ材料ポートフォリオとの比較。",
            "【調達モニタリング】エポキシ樹脂・MDI/TDI系イソシアネートの国際相場追跡を開始。",
        ]),
        ("中長期（四半期〜）", DBLUE, [
            "【ポートフォリオ】構造用接着剤市場2035年予測を踏まえ、EV・航空宇宙向けの"
            "中長期ロードマップ見直し。debonding技術への投資妥当性を議論。",
            "【原材料戦略】バイオベース代替・調達先多角化・戦略的在庫水準見直しを含む"
            "原材料戦略の再構築を提案。",
        ]),
    ]:
        ps = ParagraphStyle('ap', parent=S['sub'], textColor=color)
        story.append(Paragraph(f"▍{period}", ps))
        for item in items:
            story.append(Paragraph(f"• {item}", S['act']))

    story.append(Spacer(1,8*mm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=MG))
    story.append(Spacer(1,2*mm))
    story.append(Paragraph(
        "troy-technical.jp 独自キュレーション。記事著作権は各原著作者に帰属。"
        " | Gemini API + Claude | 2026年04月04日", S['foot']))
    return story

def pf(canvas, doc):
    canvas.saveState()
    canvas.setFont("JPP",7); canvas.setFillColor(MG)
    canvas.drawCentredString(PW/2, 10*mm,
        f"接着・封止材 Weekly Intelligence Report | 2026年04月04日 | Page {doc.page}")
    canvas.setStrokeColor(ACC); canvas.setLineWidth(2)
    canvas.line(ML, PH-8*mm, PW-MR, PH-8*mm)
    canvas.restoreState()

if __name__ == "__main__":
    out = "/home/claude/adhesive_v4.pdf"
    doc = SimpleDocTemplate(out, pagesize=A4,
        leftMargin=ML, rightMargin=MR, topMargin=MT, bottomMargin=MB)
    doc.build(build_story(), onFirstPage=pf, onLaterPages=pf)
    print(f"Done: {out}")
