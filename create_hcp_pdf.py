from weasyprint import HTML, CSS
from weasyprint.text.fonts import FontConfiguration

html_content = """
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<style>
  @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@300;400;500;700;900&display=swap');

  * { box-sizing: border-box; margin: 0; padding: 0; }

  body {
    font-family: 'Noto Sans JP', 'Hiragino Sans', 'Yu Gothic', sans-serif;
    font-size: 10pt;
    color: #1A1A2E;
    background: #FFFFFF;
  }

  /* ===== PAGE SETTINGS ===== */
  @page {
    size: A4;
    margin: 18mm 18mm 18mm 18mm;
  }
  @page cover {
    size: A4;
    margin: 0;
  }
  .page-break { page-break-after: always; }

  /* ===== COVER ===== */
  .cover {
    page: cover;
    width: 100%;
    min-height: 297mm;
    background: #FFFFFF;
    padding: 0;
    display: flex;
    flex-direction: column;
  }
  .cover-top-bar {
    background: #00A8E0;
    height: 8px;
    width: 100%;
  }
  .cover-body {
    padding: 20mm 18mm 10mm 18mm;
    flex: 1;
  }
  .cover h1 {
    font-size: 28pt;
    font-weight: 700;
    color: #00A8E0;
    line-height: 1.3;
    margin-bottom: 10px;
    margin-top: 20px;
  }
  .cover h2 {
    font-size: 13pt;
    font-weight: 400;
    color: #003087;
    margin-bottom: 8px;
  }
  .cover .citation {
    font-size: 9pt;
    color: #007AB3;
    margin-bottom: 30px;
  }

  /* KPI cards */
  .kpi-row {
    display: flex;
    gap: 12px;
    margin: 20px 0;
  }
  .kpi-card {
    flex: 1;
    background: #003087;
    border-radius: 6px;
    padding: 16px 12px;
    text-align: center;
  }
  .kpi-card .kpi-num {
    font-size: 26pt;
    font-weight: 700;
    color: #00A8E0;
    display: block;
    line-height: 1.1;
  }
  .kpi-card .kpi-label {
    font-size: 8.5pt;
    color: #FFFFFF;
    display: block;
    margin-top: 6px;
    line-height: 1.4;
  }

  .cover-tagline {
    text-align: right;
    font-size: 9pt;
    color: #003087;
    margin-top: 20px;
    font-style: italic;
  }
  .cover-bottom-bar {
    background: #003087;
    height: 12px;
    width: 100%;
    margin-top: auto;
  }

  /* ===== SECTION HEADER ===== */
  .section-header {
    display: flex;
    align-items: stretch;
    margin-bottom: 12px;
    margin-top: 20px;
  }
  .section-bar {
    width: 5px;
    background: #00A8E0;
    border-radius: 2px;
    margin-right: 10px;
    flex-shrink: 0;
  }
  .section-title {
    font-size: 12pt;
    font-weight: 700;
    color: #003087;
    padding: 4px 0;
    line-height: 1.3;
  }

  /* ===== KEY STATEMENT BOX ===== */
  .key-box {
    background: #E8F5FF;
    border-left: 4px solid #00A8E0;
    padding: 12px 16px;
    margin: 12px 0;
    border-radius: 0 4px 4px 0;
  }
  .key-box p {
    font-size: 11pt;
    font-weight: 700;
    color: #003087;
    text-align: center;
  }

  /* ===== BODY TEXT ===== */
  .body-text {
    font-size: 9.5pt;
    color: #1A1A2E;
    line-height: 1.7;
    margin-bottom: 10px;
  }
  .body-text strong {
    color: #003087;
  }

  /* ===== TABLES ===== */
  table {
    width: 100%;
    border-collapse: collapse;
    margin: 10px 0;
    font-size: 9pt;
  }
  .table-header td, .table-header th {
    background: #003087;
    color: #FFFFFF;
    font-weight: 700;
    padding: 8px 10px;
    text-align: center;
  }
  .table-row-even td {
    background: #F5F7FA;
    padding: 7px 10px;
    text-align: center;
  }
  .table-row-odd td {
    background: #FFFFFF;
    padding: 7px 10px;
    text-align: center;
  }
  .table-row-highlight td {
    background: #E8F5FF;
    padding: 8px 10px;
    text-align: center;
    font-weight: 700;
  }
  .dbs-val {
    color: #00A8E0;
    font-weight: 700;
  }
  .dbs-val-big {
    color: #00A8E0;
    font-weight: 700;
    font-size: 12pt;
  }
  .label-cell {
    background: #E8F5FF !important;
    color: #003087;
    font-weight: 700;
    text-align: left !important;
  }

  /* ===== TWO-COL COMPARISON ===== */
  .compare-table td {
    width: 50%;
    vertical-align: top;
  }
  .compare-header-navy {
    background: #003087;
    color: #FFFFFF;
    font-weight: 700;
    padding: 10px;
    text-align: center;
    font-size: 10pt;
  }
  .compare-header-teal {
    background: #007AB3;
    color: #FFFFFF;
    font-weight: 700;
    padding: 10px;
    text-align: center;
    font-size: 10pt;
  }
  .compare-body-blue {
    background: #E8F5FF;
    padding: 10px;
    font-size: 9pt;
    text-align: center;
    line-height: 1.6;
  }
  .compare-body-gray {
    background: #F5F7FA;
    padding: 10px;
    font-size: 9pt;
    text-align: center;
    line-height: 1.6;
  }

  /* ===== PATIENT CARDS ===== */
  .patient-table td {
    width: 50%;
    vertical-align: top;
  }
  .patient-card-navy {
    background: #003087;
    color: #FFFFFF;
    padding: 12px;
    text-align: center;
  }
  .patient-card-teal {
    background: #007AB3;
    color: #FFFFFF;
    padding: 12px;
    text-align: center;
  }
  .patient-card-title {
    font-size: 11pt;
    font-weight: 700;
    display: block;
    margin-bottom: 6px;
  }
  .patient-card-detail {
    font-size: 8.5pt;
    line-height: 1.6;
  }

  /* ===== SELECTION CARDS ===== */
  .selection-table td { width: 50%; vertical-align: top; }
  .sel-header-navy {
    background: #003087;
    color: #FFFFFF;
    font-weight: 700;
    padding: 10px;
    text-align: center;
    font-size: 10pt;
  }
  .sel-header-teal {
    background: #007AB3;
    color: #FFFFFF;
    font-weight: 700;
    padding: 10px;
    text-align: center;
    font-size: 10pt;
  }
  .sel-body-blue {
    background: #E8F5FF;
    padding: 12px;
    font-size: 9pt;
    line-height: 1.8;
  }
  .sel-body-gray {
    background: #F5F7FA;
    padding: 12px;
    font-size: 9pt;
    line-height: 1.8;
  }

  /* ===== SAFETY BOX ===== */
  .safety-box {
    background: #E8F5FF;
    border: 2px solid #00A8E0;
    padding: 14px;
    text-align: center;
    border-radius: 4px;
    margin: 10px 0;
  }
  .safety-box p {
    font-size: 11pt;
    font-weight: 700;
    color: #003087;
  }

  /* ===== CONCLUSION CARDS ===== */
  .conc-row { display: flex; gap: 10px; margin: 12px 0; }
  .conc-card {
    flex: 1;
    background: #003087;
    padding: 14px 10px;
    text-align: center;
    border-radius: 4px;
  }
  .conc-num {
    font-size: 20pt;
    font-weight: 700;
    color: #00A8E0;
    display: block;
    line-height: 1;
    margin-bottom: 8px;
  }
  .conc-text {
    font-size: 8.5pt;
    color: #FFFFFF;
    line-height: 1.5;
  }

  /* ===== CTA ===== */
  .cta-box {
    background: #00A8E0;
    padding: 18px;
    text-align: center;
    border-radius: 4px;
    margin: 16px 0;
  }
  .cta-box .cta-title {
    font-size: 14pt;
    font-weight: 700;
    color: #FFFFFF;
    display: block;
    margin-bottom: 8px;
  }
  .cta-box .cta-sub {
    font-size: 10pt;
    color: #FFFFFF;
  }

  /* ===== FOOTER ===== */
  .footer {
    margin-top: 20px;
    padding-top: 10px;
    border-top: 1px solid #E0E0E0;
  }
  .footer .device-name {
    font-size: 9pt;
    font-weight: 700;
    color: #003087;
    margin-bottom: 4px;
  }
  .footer .disclaimer {
    font-size: 7.5pt;
    color: #666666;
    line-height: 1.6;
  }
  .footer .tagline {
    text-align: right;
    font-size: 9pt;
    color: #003087;
    margin-top: 8px;
    font-style: italic;
  }

  /* ===== SUB HEADLINE ===== */
  .sub-headline {
    font-size: 11.5pt;
    font-weight: 700;
    color: #003087;
    margin: 14px 0 8px 0;
    line-height: 1.4;
  }
  .data-label {
    font-size: 10pt;
    font-weight: 700;
    color: #003087;
    margin: 12px 0 6px 0;
  }
</style>
</head>
<body>

<!-- ===== COVER ===== -->
<div class="cover">
  <div class="cover-top-bar"></div>
  <div class="cover-body">
    <h1>薬剤抵抗性てんかんに、<br>より確かな選択を。</h1>
    <h2>視床前核DBS（ANT-DBS）と迷走神経刺激療法（VNS）の有効性比較</h2>
    <p class="citation">Zhu J, et al. Journal of Clinical Neuroscience 90 (2021) 112–117</p>

    <div class="kpi-row">
      <div class="kpi-card">
        <span class="kpi-num">65.28%</span>
        <span class="kpi-label">ANT-DBS 平均発作減少率<br>（術後12ヵ月）</span>
      </div>
      <div class="kpi-card">
        <span class="kpi-num">72.22%</span>
        <span class="kpi-label">ANT-DBS レスポンダー率<br>（術後12ヵ月）</span>
      </div>
      <div class="kpi-card">
        <span class="kpi-num">22.22%</span>
        <span class="kpi-label">ANT-DBS 発作消失率<br>（術後12ヵ月）</span>
      </div>
    </div>

    <p class="cover-tagline">Engineering the extraordinary &nbsp;|&nbsp; Medtronic</p>
  </div>
  <div class="cover-bottom-bar"></div>
</div>

<div class="page-break"></div>

<!-- ===== SECTION 1 ===== -->
<div class="section-header">
  <div class="section-bar"></div>
  <div class="section-title">SECTION 1&nbsp;&nbsp;疾患背景：今もコントロールできていない患者がいる</div>
</div>

<div class="key-box">
  <p>てんかん患者の約30%は、薬物療法では発作をコントロールできていません。</p>
</div>

<p class="body-text">
世界に約7,000万人のてんかん患者が存在し、そのうち約3割が<strong>薬剤抵抗性てんかん（DRE）</strong>に分類されます。2剤以上の抗てんかん薬（AED）で適切に治療を行っても発作が持続するこれらの患者は、切除術の適応外であるか、手術後も発作が続くケースも少なくありません。
</p>
<p class="body-text"><strong>神経刺激療法は、薬物療法・切除術に次ぐ有力な選択肢です。</strong></p>

<table class="compare-table">
  <tr>
    <td class="compare-header-teal">VNS（迷走神経刺激療法）</td>
    <td class="compare-header-navy">ANT-DBS（視床前核深部脳刺激療法）</td>
  </tr>
  <tr>
    <td class="compare-body-blue">FDA承認：1997年（4歳以上）<br>薬剤抵抗性てんかんの標準的神経刺激療法</td>
    <td class="compare-body-gray">FDA承認：2018年（18歳以上）<br>メドトロニック Activa™ システム使用</td>
  </tr>
</table>

<!-- ===== SECTION 2 ===== -->
<div class="section-header">
  <div class="section-bar"></div>
  <div class="section-title">SECTION 2&nbsp;&nbsp;試験概要：同一チームによる厳格な単施設比較</div>
</div>

<p class="body-text">
これまで異なる施設・異なる評価基準で行われてきたVNSとANT-DBSの比較を、<strong>同一施設・同一専門チームにより初めて実施した後ろ向き研究</strong>です。
</p>

<table>
  <tr class="table-row-even">
    <td class="label-cell" style="width:35%">試験デザイン</td>
    <td>後ろ向き観察研究、単施設</td>
  </tr>
  <tr class="table-row-odd">
    <td class="label-cell">施設</td>
    <td>宣武医院 神経外科・機能神経外科センター（北京首都医科大学）</td>
  </tr>
  <tr class="table-row-even">
    <td class="label-cell">登録期間</td>
    <td>2013年6月〜2018年7月</td>
  </tr>
  <tr class="table-row-odd">
    <td class="label-cell">観察期間</td>
    <td>術前ベースライン〜術後12ヵ月</td>
  </tr>
  <tr class="table-row-even">
    <td class="label-cell">評価間隔</td>
    <td>術後3・6・9・12ヵ月</td>
  </tr>
  <tr class="table-row-odd">
    <td class="label-cell">評価方法</td>
    <td>外来での問診により発作頻度を記録</td>
  </tr>
</table>

<table class="patient-table" style="margin-top:12px;">
  <tr>
    <td class="patient-card-teal">
      <span class="patient-card-title">VNS群  N=17</span>
      <span class="patient-card-detail">平均年齢 20.24 ± 11.40歳（範囲：5〜41歳）<br>焦点性発作：8例、全般発作：9例</span>
    </td>
    <td class="patient-card-navy">
      <span class="patient-card-title">ANT-DBS群  N=18</span>
      <span class="patient-card-detail">平均年齢 28.94 ± 12.00歳（範囲：12〜52歳）<br>焦点性発作：13例、全般発作：5例</span>
    </td>
  </tr>
</table>

<div class="page-break"></div>

<!-- ===== SECTION 3 ===== -->
<div class="section-header">
  <div class="section-bar"></div>
  <div class="section-title">SECTION 3&nbsp;&nbsp;主要結果：ANT-DBSは全時点でVNSを上回る発作減少を達成</div>
</div>

<p class="sub-headline">術後12ヵ月で、ANT-DBSはVNSを約17ポイント上回る発作減少率を示しました。</p>

<p class="data-label">発作減少率の推移（術後3・6・9・12ヵ月）</p>
<table>
  <tr class="table-header">
    <td>経過期間</td>
    <td>ANT-DBS（N=18）</td>
    <td>VNS（N=17）</td>
  </tr>
  <tr class="table-row-even">
    <td>術後3ヵ月</td>
    <td class="dbs-val">57.22%</td>
    <td>36.06%</td>
  </tr>
  <tr class="table-row-odd">
    <td>術後6ヵ月</td>
    <td class="dbs-val">61.61%</td>
    <td>39.94%</td>
  </tr>
  <tr class="table-row-even">
    <td>術後9ヵ月</td>
    <td class="dbs-val">63.94%</td>
    <td>45.24%</td>
  </tr>
  <tr class="table-row-highlight">
    <td>術後12ヵ月 ★</td>
    <td class="dbs-val-big">65.28%</td>
    <td>48.35%</td>
  </tr>
</table>

<p class="data-label" style="margin-top:16px;">レスポンダー率・発作消失率（術後12ヵ月）</p>
<table>
  <tr class="table-header">
    <td>指標</td>
    <td>ANT-DBS（N=18）</td>
    <td>VNS（N=17）</td>
  </tr>
  <tr class="table-row-even">
    <td>レスポンダー率（発作50%以上減少）</td>
    <td class="dbs-val">72.22%（N=13）</td>
    <td>58.82%（N=10）</td>
  </tr>
  <tr class="table-row-odd">
    <td>発作消失率</td>
    <td class="dbs-val">22.22%（N=4）</td>
    <td>17.65%（N=3）</td>
  </tr>
  <tr class="table-row-even">
    <td>非レスポンダー</td>
    <td>27.78%（N=5）</td>
    <td>41.18%（N=7）</td>
  </tr>
</table>

<!-- ===== SECTION 4 ===== -->
<div class="section-header">
  <div class="section-bar"></div>
  <div class="section-title">SECTION 4&nbsp;&nbsp;患者選択の手引き：術前データから治療選択の根拠を</div>
</div>

<p class="body-text">本試験の結果は、術前の臨床的特徴に基づいた治療選択の可能性を示しています。</p>

<table class="selection-table">
  <tr>
    <td class="sel-header-navy">ANT-DBSが適している可能性が高い患者</td>
    <td class="sel-header-teal">VNSが適している可能性が高い患者</td>
  </tr>
  <tr>
    <td class="sel-body-blue">
      ・手術時年齢が高い<br>
      ・<strong>焦点性発作</strong>（発作減少率 71.15%）<br>
      ・<strong>開頭術の既往がある</strong>（発作減少率 81.20%）<br>
      ・<strong>てんかん罹病期間が長い</strong>（P&lt;0.05）
    </td>
    <td class="sel-body-gray">
      ・<strong>発症年齢が高い</strong>（P&lt;0.05）<br>
      ・手術時年齢が高い<br>
      ・焦点性発作（発作減少率 59.00%）
    </td>
  </tr>
</table>

<div class="page-break"></div>

<!-- ===== SECTION 5 ===== -->
<div class="section-header">
  <div class="section-bar"></div>
  <div class="section-title">SECTION 5&nbsp;&nbsp;安全性：継続刺激による良好な忍容性</div>
</div>

<div class="safety-box">
  <p>観察期間中、両治療とも重篤な合併症・副作用は報告されませんでした。</p>
</div>

<table>
  <tr class="table-header">
    <td style="width:30%">治療</td>
    <td>主な有害事象（参考）</td>
  </tr>
  <tr class="table-row-even">
    <td style="font-weight:bold;">VNS</td>
    <td>嗄声、咳嗽、疼痛、呼吸困難</td>
  </tr>
  <tr class="table-row-odd">
    <td style="font-weight:bold;">ANT-DBS</td>
    <td>植込み部疼痛、感染、リード偏位</td>
  </tr>
</table>

<!-- ===== CONCLUSION ===== -->
<div class="section-header" style="margin-top:24px;">
  <div class="section-bar"></div>
  <div class="section-title">まとめ</div>
</div>

<div class="conc-row">
  <div class="conc-card">
    <span class="conc-num">01</span>
    <span class="conc-text">ANT-DBSは全時点でVNSを<br>上回る発作減少率を達成</span>
  </div>
  <div class="conc-card">
    <span class="conc-num">02</span>
    <span class="conc-text">術前の臨床データで<br>治療選択の根拠を得られる可能性</span>
  </div>
  <div class="conc-card">
    <span class="conc-num">03</span>
    <span class="conc-text">両治療とも時間経過で効果が蓄積し、<br>安全に継続可能</span>
  </div>
</div>

<!-- ===== CTA ===== -->
<div class="cta-box">
  <span class="cta-title">薬剤抵抗性てんかんの患者さんに、次のステップを。</span>
  <span class="cta-sub">ANT-DBSの治療選択・デバイス情報は、メドトロニック担当者へお問い合わせください。</span>
</div>

<!-- ===== FOOTER ===== -->
<div class="footer">
  <p class="device-name">Medtronic Activa™ PC（型番：37601）/ Activa™ RC（型番：37612）</p>
  <p class="disclaimer">
    本資材は医療従事者向けの情報提供を目的として作成されています。記載の臨床データは Zhu J, et al. J Clin Neurosci. 2021;90:112–117 を出典とします。<br>
    本文中に記載の試験デザイン・対象患者・結果は原著論文に基づきます。本資材の内容は実際の製品使用に際して医師の判断に代わるものではありません。
  </p>
  <p class="tagline">Engineering the extraordinary &nbsp;|&nbsp; Medtronic</p>
</div>

</body>
</html>
"""

font_config = FontConfiguration()
html = HTML(string=html_content)
css = CSS(string='', font_config=font_config)
html.write_pdf(
    '/home/user/Shota/HCP_ANT-DBS_EvidenceSummary.pdf',
    font_config=font_config
)
print('PDF saved: /home/user/Shota/HCP_ANT-DBS_EvidenceSummary.pdf')
