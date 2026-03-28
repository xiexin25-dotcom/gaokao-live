"""
从大学生必备网批量抓取一分一段表，解析HTML表格并写入SQLite
"""
import sys, os, io, re, json, time, sqlite3
from urllib.request import Request, urlopen
from html.parser import HTMLParser

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
DB_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data', 'gaokao.db')

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36',
}

# 页面URL和对应的省份/年份/科类
# 格式: (url, province, year, ke_lei)
# 数据源: dxsbb.com (大学生必备网) + gaokao.eol.cn (掌上高考)
PAGES = [
    # ═══ 吉林 2024（新高考首年）═══
    ('https://www.dxsbb.com/news/146534.html', '吉林', 2024, '物理'),
    ('https://www.dxsbb.com/news/146535.html', '吉林', 2024, '历史'),
    # ═══ 吉林 2023（旧高考）═══
    ('https://www.dxsbb.com/news/136900.html', '吉林', 2023, '理科'),
    ('https://www.dxsbb.com/news/136901.html', '吉林', 2023, '文科'),
    # ═══ 吉林 2022 ═══
    ('https://www.dxsbb.com/news/117702.html', '吉林', 2022, '理科'),
    ('https://www.dxsbb.com/news/117703.html', '吉林', 2022, '文科'),

    # ═══ 辽宁 ═══ (eol.cn, dxsbb为图片格式)
    ('https://gaokao.eol.cn/liao_ning/dongtai/202306/t20230625_2446855.shtml', '辽宁', 2023, '物理'),
    ('https://gaokao.eol.cn/liao_ning/dongtai/202306/t20230625_2446857.shtml', '辽宁', 2023, '历史'),

    # ═══ 河北 ═══
    ('https://www.dxsbb.com/news/146488.html', '河北', 2024, '物理'),
    ('https://www.dxsbb.com/news/146489.html', '河北', 2024, '历史'),
    # 河北 2023 (eol.cn, dxsbb为图片格式)
    ('https://gaokao.eol.cn/he_bei/dongtai/202306/t20230625_2446858.shtml', '河北', 2023, '物理'),
    ('https://gaokao.eol.cn/he_bei/dongtai/202306/t20230625_2446944.shtml', '河北', 2023, '历史'),

    # ═══ 山东 ═══
    ('https://www.dxsbb.com/news/146529.html', '山东', 2024, '综合'),
    ('https://www.dxsbb.com/news/137046.html', '山东', 2023, '综合'),

    # ═══ 浙江 ═══ (eol.cn, dxsbb为图片格式)
    ('https://gaokao.eol.cn/zhe_jiang/dongtai/202406/t20240626_2619511.shtml', '浙江', 2024, '综合'),
    ('https://gaokao.eol.cn/zhe_jiang/dongtai/202306/t20230626_2447372.shtml', '浙江', 2023, '综合'),

    # ═══ 重庆 ═══
    ('https://www.dxsbb.com/news/146467.html', '重庆', 2024, '物理'),
    ('https://www.dxsbb.com/news/146468.html', '重庆', 2024, '历史'),
    ('https://www.dxsbb.com/news/136970.html', '重庆', 2023, '物理'),
    ('https://www.dxsbb.com/news/136971.html', '重庆', 2023, '历史'),

    # ═══ 黑龙江 2024 ═══
    ('https://www.dxsbb.com/news/146587.html', '黑龙江', 2024, '物理'),
    ('https://www.dxsbb.com/news/146588.html', '黑龙江', 2024, '历史'),
]


class TableParser(HTMLParser):
    """解析 HTML 中的 <table> 提取一分一段表数据
    注意：dxsbb.com 会在数字单元格中注入 <a class="keyWord"> 链接，
    如 <td>10<a>985</a>&nbsp;</td> 实际值为 10985。
    因此需要在 </td> 时才拼接完整文本。
    """

    def __init__(self):
        super().__init__()
        self.in_table = False
        self.in_tr = False
        self.in_td = False
        self.cell_parts = []      # 累积单个 <td> 内的文本片段
        self.current_row = []
        self.rows = []

    def handle_starttag(self, tag, attrs):
        if tag == 'table':
            self.in_table = True
        elif tag == 'tr' and self.in_table:
            self.in_tr = True
            self.current_row = []
        elif tag in ('td', 'th') and self.in_tr:
            self.in_td = True
            self.cell_parts = []

    def handle_endtag(self, tag):
        if tag == 'table':
            self.in_table = False
        elif tag == 'tr' and self.in_tr:
            self.in_tr = False
            if self.current_row:
                self.rows.append(self.current_row)
        elif tag in ('td', 'th') and self.in_td:
            # 拼接单元格内所有文本片段，去除 &nbsp; (\xa0)
            cell_text = ''.join(self.cell_parts).replace('\xa0', '').strip()
            self.current_row.append(cell_text)
            self.in_td = False

    def handle_data(self, data):
        if self.in_td:
            self.cell_parts.append(data)

    def handle_entityref(self, name):
        """处理 &nbsp; 等HTML实体"""
        if self.in_td:
            if name == 'nbsp':
                self.cell_parts.append(' ')
            else:
                self.cell_parts.append(f'&{name};')


def parse_segment_table(html):
    """解析分段表（10分一行，+9到+0列，共11列）→ [(score, cum_rank), ...]
    只匹配列数>=8的表格，避免误匹配3-4列的单分表。
    """
    parser = TableParser()
    parser.feed(html)

    # 先检查是否真的是11列段表：数据行应有>=8列
    data_rows = [r for r in parser.rows if len(r) >= 8]
    if len(data_rows) < 5:
        return []  # 不是段表格式

    data = []
    for row in data_rows:
        # Try to parse base score from first cell
        try:
            base_score = int(row[0].strip())
        except (ValueError, IndexError):
            continue

        for j in range(1, min(len(row), 11)):
            val = row[j].strip().replace(',', '').replace('，', '')
            if not val or val == '-' or val == '--':
                continue
            try:
                rank = int(val)
            except ValueError:
                continue
            score = base_score + (10 - j)
            if score >= 100:  # 过滤太低的分数
                data.append((score, rank))

    return data


def parse_single_score_table(html):
    """解析单分表（每行一个分数）→ [(score, cum_rank), ...]
    支持3列(分数/人数/累计)和15列(山东多科目)格式。
    对于15列格式，第3列是全体累计人数。
    """
    parser = TableParser()
    parser.feed(html)

    data = []
    for row in parser.rows:
        if len(row) < 2:
            continue
        # 提取所有数字列
        nums = []
        for cell in row:
            val = cell.strip().replace(',', '').replace('，', '')
            # 处理"619分及以上"这类文本
            m = re.match(r'^(\d+)', val)
            if m:
                nums.append(int(m.group(1)))
            else:
                try:
                    nums.append(int(val))
                except ValueError:
                    pass
        if len(nums) >= 3:
            # 第1个=分数, 第3个=全体累计人数（3列和15列通用）
            score, rank = nums[0], nums[2]
            if 100 <= score <= 800 and rank > 0:
                data.append((score, rank))
        elif len(nums) == 2:
            score, rank = nums[0], nums[1]
            if 100 <= score <= 800 and rank > 0:
                data.append((score, rank))

    return data


def fetch_and_parse(url):
    """Fetch URL and parse table data"""
    req = Request(url, headers=HEADERS)
    with urlopen(req, timeout=15) as resp:
        html = resp.read().decode('utf-8', errors='ignore')

    # Try segment table first (10分一行)
    data = parse_segment_table(html)
    if len(data) > 50:
        return data, 'segment'

    # Fallback: single score table
    data = parse_single_score_table(html)
    if len(data) > 20:
        return data, 'single'

    return data, 'unknown'


def main():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    # Ensure table exists
    c.execute('''CREATE TABLE IF NOT EXISTS score_rank_table (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        province TEXT NOT NULL,
        year INTEGER NOT NULL,
        ke_lei TEXT NOT NULL,
        score INTEGER NOT NULL,
        cum_rank INTEGER NOT NULL,
        UNIQUE(province, year, ke_lei, score)
    )''')
    c.execute('CREATE INDEX IF NOT EXISTS idx_srt_prov_year_ke ON score_rank_table(province, year, ke_lei)')
    conn.commit()

    total_inserted = 0
    errors = []

    for url, province, year, ke_lei in PAGES:
        print(f'Fetching {province} {year} {ke_lei}...')
        try:
            data, fmt = fetch_and_parse(url)
            if not data:
                print(f'  WARNING: No data extracted from {url}')
                errors.append((province, year, ke_lei, 'no data'))
                continue

            # Sort by score descending
            data.sort(key=lambda x: -x[0])

            # Deduplicate
            seen = set()
            unique_data = []
            for s, r in data:
                if s not in seen:
                    seen.add(s)
                    unique_data.append((s, r))

            # Insert
            for score, rank in unique_data:
                c.execute('INSERT OR REPLACE INTO score_rank_table (province, year, ke_lei, score, cum_rank) VALUES (?,?,?,?,?)',
                          (province, year, ke_lei, score, rank))

            conn.commit()
            total_inserted += len(unique_data)
            print(f'  OK: {len(unique_data)} scores ({unique_data[0][0]}~{unique_data[-1][0]}), format={fmt}')

        except Exception as e:
            print(f'  ERROR: {e}')
            errors.append((province, year, ke_lei, str(e)))

        time.sleep(1.5)  # 限速

    # Stats
    print(f'\n=== DONE ===')
    print(f'Total inserted: {total_inserted}')

    if errors:
        print(f'\nErrors ({len(errors)}):')
        for e in errors:
            print(f'  {e}')

    print('\nCoverage:')
    for row in c.execute('''SELECT province, year, ke_lei, COUNT(*), MIN(score), MAX(score)
                           FROM score_rank_table
                           GROUP BY province, year, ke_lei
                           ORDER BY province, year, ke_lei'''):
        print(f'  {row[0]} {row[1]} {row[2]}: {row[3]} scores ({row[4]}~{row[5]})')

    conn.close()


if __name__ == '__main__':
    main()
