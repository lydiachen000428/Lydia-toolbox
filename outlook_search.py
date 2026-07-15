"""
outlook_search.py — 解析郵件內文，抽取料號、pcs、CTN、鐳雕號
"""
import re
from html.parser import HTMLParser

ITEM_RE    = re.compile(r'(8SH[A-Z]\d{6}|S6SH\d{6}|\d{9,})')
# 格式1: 料號*2000pcs  (order mail 格式)；涵蓋 8SH/S6SH 及純數字料號（9碼以上）
PCS_STAR_RE = re.compile(r'(8SH[A-Z]\d{6}|S6SH\d{6}|\d{9,})\*(\d+)\s*pcs', re.I)
# 格式2: 料號 ... N CTN
CTN_RE     = re.compile(r'(\d+)\s*(?:CTN|CTNS|Carton|箱)', re.I)
PO_RE      = re.compile(r'PO\s*#?\s*(\d{5,7})', re.I)
# 鐳雕號：日期貼/鐳雕號:XXXX 或 鐳雕號:XXXX
LASER_RE   = re.compile(r'鐳雕號[:：]\s*(\d+)', re.I)
# 交期段落：遇到「交期：YYYY/M/D」開始新的一批
DATE_RE    = re.compile(r'交期[：:]\s*(\d{4}[/\-]\d{1,2}[/\-]\d{1,2})')


class _HTMLStripper(HTMLParser):
    def __init__(self):
        super().__init__()
        self._parts = []
    def handle_data(self, data): self._parts.append(data)
    def get_text(self): return '\n'.join(self._parts)

def _strip_html(html): p = _HTMLStripper(); p.feed(html); return p.get_text()


def parse_items_from_body(body: str, db: dict = None):
    """
    從郵件內文解析料號、總pcs（或CTN）、鐳雕號。
    支援格式：
      - 8SHB130231*2000pcs/SMS2DV-SP_越南  （order mail）
      - 8SHB130231 ... 50 CTN              （傳統格式）
    db: {item_no: {pcs_per_ctn: N}} 用於換算 CTN
    """
    lines = body.splitlines()
    results = []        # [{item_no, total_pcs, ctns, date_code}]
    current_laser = ''  # 當前段落的鐳雕號

    for line in lines:
        # 更新鐳雕號（段落標題行）
        laser_m = LASER_RE.search(line)
        if laser_m:
            current_laser = laser_m.group(1)

        # 格式1：料號*Npcs
        for m in PCS_STAR_RE.finditer(line):
            item_no = m.group(1)
            total_pcs = int(m.group(2))
            # 換算 CTN
            ctns = None
            if db:
                rec = db.get(item_no, {})
                ppc = rec.get('pcs_per_ctn')
                if ppc:
                    ctns = round(total_pcs / ppc)
            results.append({
                'item_no':    item_no,
                'total_pcs':  total_pcs,
                'ctns':       ctns,
                'date_code':  current_laser,
                'from_format': 'star',
            })

    # 若格式1找不到任何料號，退回格式2
    if not results:
        for i, line in enumerate(lines):
            laser_m = LASER_RE.search(line)
            if laser_m:
                current_laser = laser_m.group(1)
            for m in ITEM_RE.finditer(line):
                item_no = m.group(1)
                # 在前後5行找 CTN
                search = ' '.join(lines[max(0,i-1):min(len(lines),i+6)])
                ctn_m = CTN_RE.search(search)
                ctns = int(ctn_m.group(1)) if ctn_m else None
                results.append({
                    'item_no':   item_no,
                    'total_pcs': None,
                    'ctns':      ctns,
                    'date_code': current_laser,
                    'from_format': 'ctn',
                })

    # 去重（同料號保留第一筆）
    seen = {}
    deduped = []
    for r in results:
        if r['item_no'] not in seen:
            seen[r['item_no']] = True
            deduped.append(r)
    return deduped


def extract_po(body: str, subject: str = '') -> str:
    for text in [body, subject]:
        m = PO_RE.search(text)
        if m:
            return m.group(1)
    return ''
