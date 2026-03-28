"""
从掌上高考 API 批量采集省控线数据并写入 gaokao.db
"""
import sys, os, io, json, time, sqlite3, urllib.request

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DB_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data', 'gaokao.db')

# 掌上高考省份ID映射
PROV_IDS = {
    '北京':11,'天津':12,'河北':13,'山西':14,'内蒙古':15,
    '辽宁':21,'吉林':22,'黑龙江':23,
    '上海':31,'江苏':32,'浙江':33,'安徽':34,'福建':35,'江西':36,'山东':37,
    '河南':41,'湖北':42,'湖南':43,'广东':44,'广西':45,'海南':46,
    '重庆':50,'四川':51,'贵州':52,'云南':53,'西藏':54,
    '陕西':61,'甘肃':62,'青海':63,'宁夏':64,'新疆':65,
}

HEADERS = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)', 'Referer': 'https://gaokao.eol.cn/'}
YEARS = [2022, 2023, 2024, 2025]

def fetch_control_lines():
    """从掌上高考录取数据中提取各省各年各批次的省控线"""
    results = []

    for prov, pid in PROV_IDS.items():
        for year in YEARS:
            url = (f'https://api.zjzw.cn/web/api/?uri=apidata/api/gk/score/province'
                   f'&e_sort=zslx_rank,bindx_rank&local_province_id={pid}'
                   f'&page=1&size=200&year={year}')
            req = urllib.request.Request(url, headers=HEADERS)
            try:
                with urllib.request.urlopen(req, timeout=15) as resp:
                    data = json.loads(resp.read().decode('utf-8'))
                    items = data.get('data', {}).get('item', [])

                    # 提取唯一的 (批次, 科类, 省控线) 组合
                    seen = set()
                    for item in items:
                        batch = item.get('local_batch_name', '')
                        ke_lei = item.get('local_type_name', '')
                        proscore = item.get('proscore')
                        if proscore is None:
                            continue
                        key = (prov, year, batch, ke_lei)
                        if key not in seen:
                            seen.add(key)
                            results.append({
                                'province': prov,
                                'year': year,
                                'batch': batch,
                                'ke_lei': ke_lei,
                                'control_score': int(proscore),
                            })

                    print(f'  {prov} {year}: {len(seen)} batch/type combos from {len(items)} records')
            except Exception as e:
                print(f'  {prov} {year}: ERROR {e}')

            time.sleep(0.5)  # 限速

    return results


def save_to_db(results):
    """创建 control_line 表并写入数据"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    c.execute('''CREATE TABLE IF NOT EXISTS control_line (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        province TEXT NOT NULL,
        year INTEGER NOT NULL,
        batch TEXT NOT NULL,
        ke_lei TEXT NOT NULL,
        control_score INTEGER NOT NULL,
        UNIQUE(province, year, batch, ke_lei)
    )''')

    inserted = 0
    for r in results:
        try:
            c.execute('''INSERT OR REPLACE INTO control_line (province, year, batch, ke_lei, control_score)
                         VALUES (?, ?, ?, ?, ?)''',
                      (r['province'], r['year'], r['batch'], r['ke_lei'], r['control_score']))
            inserted += 1
        except Exception as e:
            print(f'  Insert error: {e}')

    conn.commit()

    # 验证
    total = c.execute('SELECT COUNT(*) FROM control_line').fetchone()[0]
    provinces = c.execute('SELECT COUNT(DISTINCT province) FROM control_line').fetchone()[0]
    years = c.execute('SELECT COUNT(DISTINCT year) FROM control_line').fetchone()[0]

    print(f'\n=== DONE ===')
    print(f'Inserted/updated: {inserted}')
    print(f'Total records: {total}')
    print(f'Provinces: {provinces}')
    print(f'Years: {years}')

    # 样本数据
    print('\nSample data (吉林):')
    for row in c.execute('SELECT * FROM control_line WHERE province="吉林" ORDER BY year, batch'):
        print(f'  {row}')

    conn.close()


if __name__ == '__main__':
    print('=== Fetching provincial control lines from 掌上高考 API ===\n')
    results = fetch_control_lines()
    print(f'\nTotal results: {len(results)}')

    if results:
        save_to_db(results)
