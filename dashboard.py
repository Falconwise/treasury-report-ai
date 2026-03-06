"""
Treasury Report AI - Interactive HTML Dashboard
"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import base64
import io
import os
from datetime import datetime


def _generate_utilization_chart(facilities, output_dir):
    fig, ax = plt.subplots(figsize=(10, 4.5))
    banks = []
    for f in facilities:
        name = (f.bank.replace('Saudi National Bank (SNB)', 'SNB')
                      .replace('Saudi British Bank (SABB)', 'SABB')
                      .replace('Al Rajhi Bank', 'Al Rajhi'))
        banks.append(name)

    limits = [f.limit for f in facilities]
    outstandings = [f.outstanding for f in facilities]
    availables = [f.available for f in facilities]
    y_pos = np.arange(len(banks))
    bar_height = 0.5

    navy = '#1B3A5C'
    steel = '#2E75B6'
    light_blue = '#D6E4F0'
    gold = '#C4A35A'

    bars_outstanding = ax.barh(y_pos, outstandings, bar_height,
                               label='Outstanding', color=steel, edgecolor='white', linewidth=0.5)
    bars_available = ax.barh(y_pos, availables, bar_height, left=outstandings,
                             label='Available', color=light_blue, edgecolor='white', linewidth=0.5)

    for i, limit in enumerate(limits):
        ax.plot(limit, i, 'D', color=gold, markersize=8, zorder=5)

    for i, (out, avail, limit) in enumerate(zip(outstandings, availables, limits)):
        util_pct = out / limit * 100
        ax.text(out / 2, i, 'SAR ' + str(int(out)) + 'M\n(' + f'{util_pct:.0f}' + '%)',
                ha='center', va='center', fontsize=8, fontweight='bold', color='white')
        if avail > 15:
            ax.text(out + avail / 2, i, 'SAR ' + str(int(avail)) + 'M',
                    ha='center', va='center', fontsize=7.5, color=navy)

    ax.set_yticks(y_pos)
    ax.set_yticklabels(banks, fontsize=10, fontweight='bold', color=navy)
    ax.set_xlabel('SAR Millions', fontsize=10, color=navy)
    ax.set_title('Credit Facility Utilization', fontsize=14, fontweight='bold', color=navy, pad=15)
    limit_marker = mpatches.Patch(color=gold, label='Facility Limit')
    ax.legend(handles=[bars_outstanding, bars_available, limit_marker],
              loc='lower right', fontsize=9, framealpha=0.9)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_color('#CCCCCC')
    ax.spines['left'].set_color('#CCCCCC')
    ax.tick_params(colors='#666666')
    ax.set_xlim(0, max(limits) * 1.12)
    plt.tight_layout()

    chart_path = os.path.join(output_dir, 'utilization_chart.png')
    fig.savefig(chart_path, dpi=150, bbox_inches='tight', facecolor='white')
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    buf.seek(0)
    b64 = base64.b64encode(buf.read()).decode('utf-8')
    plt.close(fig)
    return b64, chart_path


def _generate_covenant_gauge_chart(summary, output_dir):
    fig, ax = plt.subplots(figsize=(4, 4))
    sizes = [summary['compliant'], summary['warning'], summary['breach']]
    colors_list = ['#2E7D32', '#F57F17', '#C62828']
    labels = ['Compliant', 'Warning', 'Breach']
    filtered = [(s, c, l) for s, c, l in zip(sizes, colors_list, labels) if s > 0]
    if not filtered:
        filtered = [(1, '#2E7D32', 'Compliant')]
    sizes_f = [x[0] for x in filtered]
    colors_f = [x[1] for x in filtered]
    labels_f = [x[2] for x in filtered]
    wedges, texts, autotexts = ax.pie(
        sizes_f, colors=colors_f, labels=labels_f,
        autopct='%1.0f%%', startangle=90,
        pctdistance=0.75, labeldistance=1.15,
        wedgeprops=dict(width=0.4, edgecolor='white', linewidth=2)
    )
    for text in texts:
        text.set_fontsize(9)
        text.set_fontweight('bold')
    for text in autotexts:
        text.set_fontsize(9)
        text.set_color('white')
        text.set_fontweight('bold')
    total = summary['total_covenants']
    ax.text(0, 0, str(total) + '\nCovenants',
            ha='center', va='center', fontsize=14, fontweight='bold', color='#1B3A5C')
    ax.set_title('Covenant Status', fontsize=12, fontweight='bold', color='#1B3A5C', pad=10)
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    buf.seek(0)
    b64 = base64.b64encode(buf.read()).decode('utf-8')
    plt.close(fig)
    return b64


def generate_html_dashboard(analysis, output_dir):
    summary = analysis['summary']
    results = analysis['results']
    facilities = analysis['facilities']
    financials = analysis['financials']

    util_b64, chart_path = _generate_utilization_chart(facilities, output_dir)
    gauge_b64 = _generate_covenant_gauge_chart(summary, output_dir)

    overall = summary['overall_status']
    if overall == "COMPLIANT":
        status_bg = "#2E7D32"
        status_icon = "&#10003;"
    elif overall == "WARNING":
        status_bg = "#F57F17"
        status_icon = "&#9888;"
    else:
        status_bg = "#C62828"
        status_icon = "&#10007;"

    covenant_rows = ""
    current_bank = ""
    for r in results:
        if r.bank != current_bank:
            current_bank = r.bank
            covenant_rows += '<tr class="bank-header"><td colspan="6">' + r.bank + '</td></tr>\n'
        if r.status == "COMPLIANT":
            sc, ic = "status-green", "&#10003;"
        elif r.status == "WARNING":
            sc, ic = "status-amber", "&#9888;"
        else:
            sc, ic = "status-red", "&#10007;"
        hc = "positive" if r.headroom_pct >= 0 else "negative"
        covenant_rows += '<tr>'
        covenant_rows += '<td class="covenant-name">' + r.covenant_name + '</td>'
        covenant_rows += '<td class="center">' + r.threshold_str + '</td>'
        covenant_rows += '<td class="center bold">' + r.actual_display + '</td>'
        covenant_rows += '<td class="center ' + hc + '">' + f'{r.headroom_pct:+.1f}' + '%</td>'
        covenant_rows += '<td class="center"><span class="' + sc + '">' + ic + ' ' + r.status + '</span></td>'
        covenant_rows += '</tr>\n'

    facility_rows = ""
    total_limit = sum(f.limit for f in facilities)
    total_outstanding = sum(f.outstanding for f in facilities)
    total_available = sum(f.available for f in facilities)
    for f in facilities:
        uc = "#2E7D32" if f.utilization_pct < 75 else ("#F57F17" if f.utilization_pct < 90 else "#C62828")
        sb = (f.bank.replace('Saudi National Bank (SNB)', 'SNB')
                    .replace('Saudi British Bank (SABB)', 'SABB')
                    .replace('Al Rajhi Bank', 'Al Rajhi'))
        facility_rows += '<tr>'
        facility_rows += '<td class="bold">' + sb + '</td>'
        facility_rows += '<td>' + f.facility_type + '</td>'
        facility_rows += '<td class="right">' + f'{f.limit:.0f}' + '</td>'
        facility_rows += '<td class="right bold">' + f'{f.outstanding:.0f}' + '</td>'
        facility_rows += '<td class="right">' + f'{f.available:.0f}' + '</td>'
        facility_rows += '<td class="center" style="color:' + uc + '; font-weight:bold;">' + f'{f.utilization_pct:.1f}' + '%</td>'
        facility_rows += '<td class="center">' + f.rate + '</td>'
        facility_rows += '</tr>\n'
    total_util = total_outstanding / total_limit * 100 if total_limit > 0 else 0
    facility_rows += '<tr class="total-row">'
    facility_rows += '<td class="bold">TOTAL</td><td></td>'
    facility_rows += '<td class="right bold">' + f'{total_limit:.0f}' + '</td>'
    facility_rows += '<td class="right bold">' + f'{total_outstanding:.0f}' + '</td>'
    facility_rows += '<td class="right bold">' + f'{total_available:.0f}' + '</td>'
    facility_rows += '<td class="center bold">' + f'{total_util:.1f}' + '%</td><td></td>'
    facility_rows += '</tr>'

    maturity_rows = ""
    sorted_fac = sorted(facilities, key=lambda f: f.maturity_date)
    for f in sorted_fac:
        sb = (f.bank.replace('Saudi National Bank (SNB)', 'SNB')
                    .replace('Saudi British Bank (SABB)', 'SABB')
                    .replace('Al Rajhi Bank', 'Al Rajhi'))
        remaining = f.remaining_years
        urgency = "urgent" if remaining < 1.5 else ("soon" if remaining < 2.5 else "safe")
        maturity_rows += '<tr>'
        maturity_rows += '<td class="bold">' + sb + '</td>'
        maturity_rows += '<td>' + f.facility_type + '</td>'
        maturity_rows += '<td class="right bold">SAR ' + f'{f.outstanding:.0f}' + 'M</td>'
        maturity_rows += '<td class="center">' + f.maturity_date + '</td>'
        maturity_rows += '<td class="center maturity-' + urgency + '">' + f'{remaining:.1f}' + ' years</td>'
        maturity_rows += '</tr>\n'

    debt_equity = financials.total_liabilities / financials.total_equity if financials.total_equity else 0
    current_ratio = financials.total_current_assets / financials.total_current_liabilities if financials.total_current_liabilities else 0
    ebitda_margin = financials.ebitda / financials.revenue * 100 if financials.revenue else 0
    now_str = datetime.now().strftime('%d %B %Y, %H:%M')
    now_date = datetime.now().strftime('%d %B %Y')

    html = '<!DOCTYPE html>\n<html lang="en">\n<head>\n'
    html += '<meta charset="UTF-8">\n<meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
    html += '<title>Treasury Report AI - Covenant Compliance Dashboard</title>\n'
    html += '<style>\n'
    html += CSS_TEMPLATE
    html += '</style>\n</head>\n<body>\n'

    # Header
    html += '<div class="header"><div>'
    html += '<h1>' + summary['company_name'] + '</h1>'
    html += '<div class="subtitle">Covenant Compliance Dashboard | ' + summary['reporting_date'] + '</div>'
    html += '</div><div class="brand">'
    html += '<div class="logo">Treasury Report AI</div>'
    html += '<div class="date">Generated: ' + now_str + '</div>'
    html += '</div></div>\n'

    # Container start
    html += '<div class="container">\n'

    # Status banner
    html += '<div class="status-banner" style="background:' + status_bg + '">'
    html += '<span class="icon">' + status_icon + '</span>'
    html += '<span>Overall Compliance Status: ' + overall + '</span></div>\n'

    # KPI cards
    html += '<div class="kpi-row">'
    html += '<div class="kpi-card green"><div class="label">Compliant</div>'
    html += '<div class="value">' + str(summary['compliant']) + '</div>'
    html += '<div class="detail">covenants within safe range</div></div>'
    html += '<div class="kpi-card amber"><div class="label">Warning</div>'
    html += '<div class="value">' + str(summary['warning']) + '</div>'
    html += '<div class="detail">within 15% of breach threshold</div></div>'
    html += '<div class="kpi-card red"><div class="label">Breach</div>'
    html += '<div class="value">' + str(summary['breach']) + '</div>'
    html += '<div class="detail">covenants in breach</div></div>'
    html += '<div class="kpi-card gold"><div class="label">Total Facilities</div>'
    html += '<div class="value">SAR ' + f'{total_limit:.0f}' + 'M</div>'
    html += '<div class="detail">' + f'{total_util:.1f}' + '% utilized (SAR ' + f'{total_outstanding:.0f}' + 'M drawn)</div></div>'
    html += '</div>\n'

    # Covenant detail table
    html += '<div class="section"><div class="section-header">Financial Covenant Compliance Detail</div>'
    html += '<div class="section-body"><table><thead><tr>'
    html += '<th>Covenant</th><th class="center">Required</th><th class="center">Actual</th>'
    html += '<th class="center">Headroom</th><th class="center">Status</th>'
    html += '</tr></thead><tbody>' + covenant_rows + '</tbody></table></div></div>\n'

    # Two column: chart + gauge
    html += '<div class="two-col">'
    html += '<div class="section"><div class="section-header">Facility Utilization</div>'
    html += '<div class="chart-container"><img src="data:image/png;base64,' + util_b64 + '" alt="Facility Utilization Chart"></div></div>'
    html += '<div class="section"><div class="section-header">Covenant Overview</div>'
    html += '<div class="chart-container" style="padding-top:30px;">'
    html += '<img src="data:image/png;base64,' + gauge_b64 + '" alt="Covenant Status" style="max-width:300px;">'
    html += '<div style="margin-top:15px;"><table style="width:80%;margin:0 auto;">'
    html += '<tr><td class="bold">EBITDA Margin</td><td class="right bold" style="color:var(--green)">' + f'{ebitda_margin:.1f}' + '%</td></tr>'
    html += '<tr><td class="bold">Debt / Equity</td><td class="right bold">' + f'{debt_equity:.2f}' + 'x</td></tr>'
    html += '<tr><td class="bold">Current Ratio</td><td class="right bold">' + f'{current_ratio:.2f}' + 'x</td></tr>'
    html += '<tr><td class="bold">Net Income</td><td class="right bold">SAR ' + f'{financials.net_income:.0f}' + 'M</td></tr>'
    html += '</table></div></div></div></div>\n'

    # Facility details
    html += '<div class="section"><div class="section-header">Credit Facility Summary (SAR Millions)</div>'
    html += '<div class="section-body"><table><thead><tr>'
    html += '<th>Bank</th><th>Type</th><th class="right">Limit</th><th class="right">Outstanding</th>'
    html += '<th class="right">Available</th><th class="center">Utilization</th><th class="center">Rate</th>'
    html += '</tr></thead><tbody>' + facility_rows + '</tbody></table></div></div>\n'

    # Maturity timeline
    html += '<div class="section"><div class="section-header">Maturity Timeline</div>'
    html += '<div class="section-body"><table><thead><tr>'
    html += '<th>Bank</th><th>Facility Type</th><th class="right">Outstanding</th>'
    html += '<th class="center">Maturity Date</th><th class="center">Time Remaining</th>'
    html += '</tr></thead><tbody>' + maturity_rows + '</tbody></table></div></div>\n'

    # Footer
    html += '</div>\n'  # container
    html += '<div class="footer"><span class="brand-footer">Treasury Report AI</span> | '
    html += 'Automated Covenant Compliance Monitoring for Saudi Corporates | ' + now_date
    html += '<br><span style="font-size:11px; opacity:0.7;">This report is auto-generated from management accounts. '
    html += 'For official covenant compliance, refer to audited financial statements.</span></div>\n'
    html += '</body>\n</html>'

    output_path = os.path.join(output_dir, 'dashboard.html')
    with open(output_path, 'w', encoding='utf-8') as fh:
        fh.write(html)
    return output_path


CSS_TEMPLATE = """
:root {
    --navy: #0D1B2A; --dark-blue: #1B3A5C; --steel-blue: #2E75B6;
    --light-blue: #D6E4F0; --gold: #C4A35A; --green: #2E7D32;
    --amber: #F57F17; --red: #C62828; --bg: #F8F9FA;
    --card-bg: #FFFFFF; --text: #333333; --text-light: #666666; --border: #E0E0E0;
}
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif; background: var(--bg); color: var(--text); line-height: 1.6; }
.header { background: linear-gradient(135deg, var(--navy) 0%, var(--dark-blue) 100%); color: white; padding: 30px 40px; display: flex; justify-content: space-between; align-items: center; }
.header h1 { font-size: 24px; font-weight: 700; letter-spacing: 0.5px; }
.header .subtitle { opacity: 0.8; font-size: 14px; margin-top: 4px; }
.header .brand { text-align: right; }
.header .brand .logo { font-size: 18px; font-weight: 700; color: var(--gold); }
.header .brand .date { font-size: 12px; opacity: 0.7; margin-top: 4px; }
.container { max-width: 1280px; margin: 0 auto; padding: 25px; }
.kpi-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; margin-bottom: 25px; }
.kpi-card { background: var(--card-bg); border-radius: 10px; padding: 22px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); border-left: 4px solid var(--steel-blue); transition: transform 0.2s; }
.kpi-card:hover { transform: translateY(-2px); box-shadow: 0 4px 15px rgba(0,0,0,0.1); }
.kpi-card .label { font-size: 12px; text-transform: uppercase; color: var(--text-light); letter-spacing: 1px; }
.kpi-card .value { font-size: 28px; font-weight: 700; color: var(--dark-blue); margin: 6px 0 2px; }
.kpi-card .detail { font-size: 12px; color: var(--text-light); }
.kpi-card.green { border-left-color: var(--green); }
.kpi-card.amber { border-left-color: var(--amber); }
.kpi-card.red { border-left-color: var(--red); }
.kpi-card.gold { border-left-color: var(--gold); }
.status-banner { color: white; padding: 18px 30px; border-radius: 10px; margin-bottom: 25px; display: flex; align-items: center; gap: 15px; font-size: 18px; font-weight: 600; box-shadow: 0 3px 12px rgba(0,0,0,0.15); }
.status-banner .icon { font-size: 32px; }
.section { background: var(--card-bg); border-radius: 10px; margin-bottom: 25px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); overflow: hidden; }
.section-header { background: linear-gradient(135deg, var(--dark-blue), var(--steel-blue)); color: white; padding: 15px 25px; font-size: 16px; font-weight: 600; letter-spacing: 0.3px; }
.section-body { padding: 0; }
.two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 25px; }
table { width: 100%; border-collapse: collapse; }
thead th { background: var(--dark-blue); color: white; padding: 12px 15px; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; text-align: left; }
tbody td { padding: 10px 15px; font-size: 13px; border-bottom: 1px solid var(--border); }
tbody tr:hover { background: #F5F8FC; }
tbody tr:nth-child(even) { background: #FAFBFC; }
.bank-header { background: var(--light-blue) !important; }
.bank-header td { font-weight: 700; color: var(--dark-blue); font-size: 14px; padding: 10px 15px; }
.total-row { background: var(--light-blue) !important; }
.total-row td { font-weight: 700; border-top: 2px solid var(--steel-blue); }
.center { text-align: center; } .right { text-align: right; } .bold { font-weight: 600; }
.covenant-name { padding-left: 30px !important; }
.positive { color: var(--green); font-weight: 600; }
.negative { color: var(--red); font-weight: 600; }
.status-green { background: #E8F5E9; color: var(--green); padding: 4px 12px; border-radius: 20px; font-weight: 600; font-size: 12px; white-space: nowrap; }
.status-amber { background: #FFF8E1; color: var(--amber); padding: 4px 12px; border-radius: 20px; font-weight: 600; font-size: 12px; white-space: nowrap; }
.status-red { background: #FFEBEE; color: var(--red); padding: 4px 12px; border-radius: 20px; font-weight: 600; font-size: 12px; white-space: nowrap; }
.maturity-urgent { color: var(--red); font-weight: 700; }
.maturity-soon { color: var(--amber); font-weight: 600; }
.maturity-safe { color: var(--green); }
.chart-container { padding: 20px; text-align: center; }
.chart-container img { max-width: 100%; height: auto; border-radius: 8px; }
.footer { text-align: center; padding: 25px; color: var(--text-light); font-size: 12px; border-top: 1px solid var(--border); margin-top: 20px; }
.footer .brand-footer { color: var(--gold); font-weight: 700; }
@media print { body { background: white; } .container { max-width: 100%; padding: 10px; } .section { page-break-inside: avoid; } }
"""


if __name__ == "__main__":
    print("Dashboard module loaded. Use generate_html_dashboard(analysis, output_dir).")