"""
构建一分一段表（score_rank_table）
从浏览器抓取的JSON数据导入SQLite
"""
import sys, os, io, json, sqlite3

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
DB_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data', 'gaokao.db')


def create_table(conn):
    c = conn.cursor()
    c.execute('DROP TABLE IF EXISTS score_rank_table')
    c.execute('''CREATE TABLE score_rank_table (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        province TEXT NOT NULL,
        year INTEGER NOT NULL,
        ke_lei TEXT NOT NULL,
        score INTEGER NOT NULL,
        cum_rank INTEGER NOT NULL,
        UNIQUE(province, year, ke_lei, score)
    )''')
    c.execute('CREATE INDEX idx_srt_prov_year_ke ON score_rank_table(province, year, ke_lei)')
    c.execute('CREATE INDEX idx_srt_score ON score_rank_table(province, year, ke_lei, score)')
    conn.commit()


def insert_data(conn, province, year, ke_lei, data):
    """data: list of [score, cum_rank]"""
    c = conn.cursor()
    inserted = 0
    for score, rank in data:
        try:
            c.execute('INSERT OR REPLACE INTO score_rank_table (province, year, ke_lei, score, cum_rank) VALUES (?,?,?,?,?)',
                      (province, year, ke_lei, score, rank))
            inserted += 1
        except Exception as e:
            print(f'  Error: {e}')
    conn.commit()
    print(f'  {province} {year} {ke_lei}: {inserted} rows inserted')
    return inserted


def load_json_file(filepath):
    """Load score-rank data from a JSON file"""
    with open(filepath, 'r', encoding='utf-8') as f:
        return json.load(f)


def stats(conn):
    c = conn.cursor()
    total = c.execute('SELECT COUNT(*) FROM score_rank_table').fetchone()[0]
    print(f'\n=== score_rank_table Stats ===')
    print(f'Total rows: {total}')

    for row in c.execute('''SELECT province, year, ke_lei, COUNT(*), MIN(score), MAX(score)
                           FROM score_rank_table
                           GROUP BY province, year, ke_lei
                           ORDER BY province, year, ke_lei'''):
        print(f'  {row[0]} {row[1]} {row[2]}: {row[3]} scores ({row[4]}~{row[5]})')

    # Sample lookup
    print('\nSample: 吉林 2024 物理 640分 →')
    row = c.execute('SELECT cum_rank FROM score_rank_table WHERE province="吉林" AND year=2024 AND ke_lei="物理" AND score=640').fetchone()
    if row:
        print(f'  位次: {row[0]}')


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('--init', action='store_true', help='Create/reset table')
    parser.add_argument('--import-file', type=str, help='Import JSON file: {"province":"吉林","year":2024,"ke_lei":"物理","data":[[score,rank],...]}')
    parser.add_argument('--import-inline', type=str, help='Import inline JSON')
    parser.add_argument('--stats', action='store_true', help='Show stats')
    args = parser.parse_args()

    conn = sqlite3.connect(DB_PATH)

    if args.init:
        create_table(conn)
        print('Table created.')

    if args.import_file:
        obj = load_json_file(args.import_file)
        insert_data(conn, obj['province'], obj['year'], obj['ke_lei'], obj['data'])

    if args.import_inline:
        obj = json.loads(args.import_inline)
        insert_data(conn, obj['province'], obj['year'], obj['ke_lei'], obj['data'])

    if args.stats:
        stats(conn)

    conn.close()
