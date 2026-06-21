import pandas as pd
import json
import os
from datetime import datetime

GOLD_PRICE = 930
SILVER_PRICE = 17

def analyze_sales_data(file_path):
    df = pd.read_excel(file_path)
    df = df[df['销售时间'].notna()].copy()
    df['销售时间'] = pd.to_datetime(df['销售时间'])

    max_date = df['销售时间'].max()
    cutoff = max_date - pd.Timedelta(days=30)
    df_month = df[df['销售时间'] >= cutoff].copy()

    total_revenue = float(df_month['业绩金额'].sum())
    total_profit = float(df_month['利润金额'].sum())

    df_sales = df_month[df_month['业绩金额'] > 0].copy()
    sales_count = len(df_sales)
    sales_revenue = float(df_sales['业绩金额'].sum())

    # 价格区间
    bins = [0, 100, 300, 500, 1000, 2000, 5000, float('inf')]
    labels = ['0-100元', '100-300元', '300-500元', '500-1000元', '1000-2000元', '2000-5000元', '5000元以上']
    df_sales['价格区间'] = pd.cut(df_sales['业绩金额'], bins=bins, labels=labels, right=True)

    range_stats = []
    range_details = {}
    for label in labels:
        subset = df_sales[df_sales['价格区间'] == label]
        if len(subset) == 0:
            continue
        range_stats.append({
            'range': label,
            'count': len(subset),
            'total': round(float(subset['业绩金额'].sum()), 2),
            'avg': round(float(subset['业绩金额'].mean()), 2),
        })
        transactions = []
        for _, row in subset.iterrows():
            transactions.append({
                'date': row['销售时间'].strftime('%m-%d %H:%M'),
                'name': str(row['商品名称']),
                'cat': str(row['小类']) if pd.notna(row['小类']) else '',
                'amount': round(float(row['业绩金额']), 2),
            })
        range_details[label] = {
            'transactions': transactions,
            'total_amount': round(float(subset['业绩金额'].sum()), 2),
            'count': len(subset)
        }

    # 品类
    cat_stats = []
    for cat, group in df_sales.groupby('小类'):
        if pd.notna(cat):
            cat_stats.append({
                'name': str(cat),
                'count': len(group),
                'total': round(float(group['业绩金额'].sum()), 2),
            })
    cat_stats.sort(key=lambda x: x['total'], reverse=True)

    # 商品TOP8
    prod_stats = []
    for name, group in df_sales.groupby('商品名称'):
        if pd.notna(name):
            prod_stats.append({
                'name': str(name),
                'count': len(group),
                'total': round(float(group['业绩金额'].sum()), 2),
            })
    prod_stats.sort(key=lambda x: x['total'], reverse=True)
    prod_stats = prod_stats[:8]

    # 回收分析
    rec = df_month[df_month['销售类型'] == '回收'].copy()
    recycle_items = []
    total_paid = 0.0
    total_gold_g = 0.0
    total_silver_g = 0.0
    gold_paid = 0.0
    silver_paid = 0.0

    for _, row in rec.iterrows():
        name = str(row['商品名称'])
        weight = abs(float(row['金重'])) if pd.notna(row['金重']) else 0.0
        paid = abs(float(row['业绩金额'])) if pd.notna(row['业绩金额']) else 0.0
        date_str = row['销售时间'].strftime('%m-%d %H:%M')
        is_gold = any(kw in name for kw in ['金', 'AU', 'au', '素圈'])
        metal = 'gold' if is_gold else 'silver'
        price = GOLD_PRICE if is_gold else SILVER_PRICE
        value = round(weight * price, 2)
        pnl = round(value - paid, 2)

        total_paid += paid
        if is_gold:
            total_gold_g += weight
            gold_paid += paid
        else:
            total_silver_g += weight
            silver_paid += paid

        recycle_items.append({
            'date': date_str,
            'name': name,
            'metal': metal,
            'weight': round(weight, 3),
            'paid': round(paid, 2),
            'value': value,
            'pnl': pnl,
            'price': price
        })

    total_value = round(total_gold_g * GOLD_PRICE + total_silver_g * SILVER_PRICE, 2)
    total_pnl = round(total_value - total_paid, 2)
    gold_value = round(total_gold_g * GOLD_PRICE, 2)
    silver_value = round(total_silver_g * SILVER_PRICE, 2)

    # 组装数据
    data = {
        'generate_time': datetime.now().strftime('%Y-%m-%d %H:%M'),
        'total_revenue': round(total_revenue, 2),
        'total_profit': round(total_profit, 2),
        'sales_count': sales_count,
        'sales_revenue': round(sales_revenue, 2),
        'range_stats': range_stats,
        'range_details': range_details,
        'cat_stats': cat_stats,
        'prod_stats': prod_stats,
        'recycle_items': recycle_items,
        'recycle_total_paid': round(total_paid, 2),
        'recycle_total_value': total_value,
        'recycle_total_pnl': total_pnl,
        'recycle_gold_g': round(total_gold_g, 3),
        'recycle_silver_g': round(total_silver_g, 3),
        'recycle_gold_paid': round(gold_paid, 2),
        'recycle_silver_paid': round(silver_paid, 2),
        'recycle_gold_value': gold_value,
        'recycle_silver_value': silver_value,
        'recycle_gold_count': len([x for x in recycle_items if x['metal'] == 'gold']),
        'recycle_silver_count': len([x for x in recycle_items if x['metal'] == 'silver']),
        'gold_price': GOLD_PRICE,
        'silver_price': SILVER_PRICE,
    }

    # 生成HTML
    report_html = generate_html_report(data)
    data['report_html'] = report_html
    return data


def generate_html_report(data):
    template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'report_template.html')
    with open(template_path, 'r', encoding='utf-8') as f:
        html = f.read()
    # 注入数据
    data_json = json.dumps(data, ensure_ascii=False)
    html = html.replace('__DATA__PLACEHOLDER__', data_json)
    return html


if __name__ == '__main__':
    result = analyze_sales_data(
        '/Users/hanyang/Library/Containers/com.tencent.xinWeChat/Data/Documents/xwechat_files/qq410466366_6561/msg/file/2026-06/销售明细查询_2026-06-18.xls'
    )
    print(json.dumps({k: v for k, v in result.items() if k != 'report_html'}, ensure_ascii=False, indent=2))
