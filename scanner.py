#!/usr/bin/env python3
"""
SCOSCHE Historical Packing List Scanner
掃描歷史 PACKING xlsx，抓料號規格 + 嘜頭，存入 data/items.json。

使用：
  python scanner.py            # 增量（只掃新檔）
  python scanner.py --full     # 完整重掃
"""

import json, os, re, sys, time, zipfile
from datetime import datetime
from pathlib import Path
import xml.etree.ElementTree as ET

# ── 路徑設定 ──
BASE_DIR    = Path(__file__).parent
CONFIG_PATH = BASE_DIR / 'config.json'
DATA_DIR    = BASE_DIR / 'data'
JSON_PATH   = DATA_DIR / 'items.json'
SCAN_META   = DATA_DIR / 'scan_meta.json'

cfg = json.loads(CONFIG_PATH.read_text(encoding='utf-8')) if CONFIG_PATH.exists() else {}
HIST_FOLDER = cfg.get('hist_folder', '')
SCAN_MONTHS = int(cfg.get('scan_months', 12))
CUSTOMER    = 'SCOSCHE'
ITEM_RE     = re.compile(r'^[A-Z0-9]{6,}$')   # 6碼以上英數字，涵蓋所有 SCOSCHE 料號格式
NS          = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
DATA_DIR.mkdir(parents=True, exist_ok=True)


def get_files(full_scan=False):
    last_mtime = 0
    if not full_scan and SCAN_META.exists():
        last_mtime = json.loads(SCAN_META.read_text()).get('last_scan_mtime', 0)

    cutoff = time.time() - SCAN_MONTHS * 30 * 86400
    files = []
    root = HIST_FOLDER
    if not root or not os.path.isdir(root):
        print(f"[ERROR] hist_folder 不存在：{root!r}", file=sys.stderr)
        print("請修改 config.json 中的 hist_folder 路徑", file=sys.stderr)
        return []

    def add(fp):
        mt = os.path.getmtime(fp)
        if mt >= cutoff and (not last_mtime or mt > last_mtime):
            files.append((fp, mt))

    for f in os.listdir(root):
        if 'PACKING' in f and f.endswith('.xlsx') and not f.startswith('~$'):
            add(os.path.join(root, f))

    for d in os.listdir(root):
        dp = os.path.join(root, d)
        if not os.path.isdir(dp):
            continue
        try:
            for f in os.listdir(dp):
                if 'PACKING' in f and f.endswith('.xlsx') and not f.startswith('~$'):
                    add(os.path.join(dp, f))
        except Exception:
            pass

    files.sort(key=lambda x: x[1], reverse=True)
    return files


def col_idx(ref):
    m = re.match(r'([A-Z]+)', ref)
    if not m:
        return -1
    c = 0
    for ch in m.group(1):
        c = c * 26 + (ord(ch) - 64)
    return c - 1


def parse_xlsx(path):
    try:
        with zipfile.ZipFile(path, 'r') as z:
            names = z.namelist()
            ss = []
            if 'xl/sharedStrings.xml' in names:
                root = ET.fromstring(z.read('xl/sharedStrings.xml'))
                for si in root.findall('s:si', NS):
                    ss.append(''.join(t.text or '' for t in si.findall('.//s:t', NS)))

            sheet_file = None
            if 'xl/workbook.xml' in names and 'xl/_rels/workbook.xml.rels' in names:
                rels = {}
                for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels')):
                    rels[r.get('Id')] = r.get('Target')
                for sh in ET.fromstring(z.read('xl/workbook.xml')).findall(
                        './/{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet'):
                    rid = sh.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    if sh.get('name') == 'PAC' and rid in rels:
                        sheet_file = 'xl/' + rels[rid]
                        break
            if not sheet_file:
                sheet_file = next(
                    (n for n in names if n.startswith('xl/worksheets/sheet') and n.endswith('.xml')), None)
            if not sheet_file:
                return {}
            ws_root = ET.fromstring(z.read(sheet_file))
    except Exception:
        return {}

    def cell_val(c):
        v = c.find('s:v', NS)
        if v is None or v.text is None:
            return ''
        return ss[int(v.text)] if c.get('t') == 's' else v.text

    rows_data = []
    for row_el in ws_root.findall('.//s:row', NS):
        rd = {}
        for c in row_el.findall('s:c', NS):
            ci = col_idx(c.get('r', ''))
            if ci >= 0:
                rd[ci] = cell_val(c)
        rows_data.append(rd)

    items = {}
    for rd in rows_data:
        item_no = rd.get(5, '').strip()
        if not ITEM_RE.match(item_no):
            continue
        dim = rd.get(8, '').strip().lstrip('=')
        if not dim or dim == 'None':
            continue
        desc = rd.get(6, '').strip()
        try: pcs = int(float(rd.get(7, 0))) if rd.get(7) else None
        except: pcs = None
        try: nw = float(rd.get(10, 0)) if rd.get(10) else None
        except: nw = None
        try: gw = float(rd.get(11, 0)) if rd.get(11) else None
        except: gw = None
        if item_no not in items:
            items[item_no] = dict(desc=desc, pcs_per_ctn=pcs, dim=dim, nw=nw, gw=gw, shipping_mark='')

    if not items:
        return {}

    PART_RE = re.compile(r'^Part Number:\s*(.+)', re.I)
    ROHS_RE = re.compile(r'GOODS ROHS COMPLIANT', re.I)

    # 建立兩層比對 map：
    #   prefix_map: desc 分號前的前綴 → item_no  (精確比對，優先)
    #   full_map:   完整 desc → item_no           (模糊比對，備用)
    prefix_map = {}
    full_map   = {}
    for item_no, d in items.items():
        desc = d.get('desc', '')
        if desc:
            full_map[desc] = item_no
            prefix = desc.split(';')[0].strip()
            if prefix:
                prefix_map[prefix] = item_no

    def match_part(part_desc):
        """把嘜頭的 Part Number 值比對回 item_no"""
        # 1. 完全符合 prefix
        if part_desc in prefix_map:
            return prefix_map[part_desc]
        # 2. part_desc 包含在某個 desc 裡，或某個 desc 包含在 part_desc 裡
        for desc, k in full_map.items():
            if desc and (part_desc in desc or desc in part_desc):
                return k
        # 3. part_desc 包含在某個 prefix 裡，或某個 prefix 包含在 part_desc 裡
        for pfx, k in prefix_map.items():
            if pfx and (part_desc in pfx or pfx in part_desc):
                return k
        return None

    mark_start = None
    for i, rd in enumerate(rows_data):
        if any('SHIPPING MARK' in str(v).upper() for v in rd.values()):
            mark_start = i + 1
            break

    if mark_start is not None:
        # 找出所有含 "Part Number:" 的欄
        mark_cols = set()
        for rd in rows_data[mark_start:]:
            for ci, val in rd.items():
                if PART_RE.match(str(val).strip()):
                    mark_cols.add(ci)

        for ci in sorted(mark_cols):
            lines = [rd[ci].strip() for rd in rows_data[mark_start:]
                     if ci in rd and rd[ci].strip()]
            if not lines:
                continue

            # 同一欄可能有多個料號的嘜頭堆疊，以 "Part Number:" 為邊界切塊
            blocks = []
            cur_block = []
            for ln in lines:
                if PART_RE.match(ln) and cur_block:
                    blocks.append(cur_block)
                    cur_block = [ln]
                else:
                    cur_block.append(ln)
            if cur_block:
                blocks.append(cur_block)

            for block in blocks:
                # 找 Part Number 行
                part_line = next((l for l in block if PART_RE.match(l)), None)
                if not part_line:
                    continue
                part_desc = PART_RE.match(part_line).group(1).strip()
                matched = match_part(part_desc)
                if not matched:
                    continue

                # 從 Part Number 開始截到 ROHS/RFID
                start_idx = block.index(part_line)
                trimmed = []
                for line in block[start_idx:]:
                    trimmed.append(line)
                    if 'RFID' in line.upper():
                        break
                    if ROHS_RE.search(line):
                        nxt_idx = block.index(line) + 1
                        if nxt_idx < len(block) and 'RFID' in block[nxt_idx].upper():
                            trimmed.append(block[nxt_idx])
                        break
                items[matched]['shipping_mark'] = '\n'.join(trimmed)

    return items


def scan(full_scan=False):
    t0 = time.time()
    files = get_files(full_scan)
    if not files and not full_scan:
        print("無新增檔案，增量掃描完成")
        return

    print(f"{'完整' if full_scan else '增量'}掃描：{len(files)} 個檔案")

    # 載入現有 DB
    existing = {}
    if JSON_PATH.exists():
        for r in json.loads(JSON_PATH.read_text(encoding='utf-8')):
            existing[r['item_no']] = r

    new_c = upd_c = 0
    now = datetime.now().isoformat()
    for fp, mtime in files:
        data = parse_xlsx(fp)
        if not data:
            continue
        fname = os.path.basename(fp)
        for item_no, d in data.items():
            ex = existing.get(item_no)
            if ex is None:
                existing[item_no] = dict(
                    customer=CUSTOMER, item_no=item_no,
                    desc=d['desc'], pcs_per_ctn=d['pcs_per_ctn'],
                    dim=d['dim'], nw=d['nw'], gw=d['gw'],
                    shipping_mark=d['shipping_mark'],
                    source_file=fname, source_mtime=mtime, updated_at=now)
                new_c += 1
            elif ex.get('manual_edited'):
                # 使用者在介面手動修改過的料號，掃描不覆蓋
                continue
            elif mtime > (ex.get('source_mtime') or 0) or (full_scan and not ex.get('shipping_mark') and d['shipping_mark']):
                # 更新條件：檔案有更新，或完整掃描時嘜頭是空白且本次抓到了
                # 新檔沒抓到嘜頭時保留原有嘜頭，避免被清空
                ex.update(desc=d['desc'], pcs_per_ctn=d['pcs_per_ctn'],
                           dim=d['dim'], nw=d['nw'], gw=d['gw'],
                           shipping_mark=d['shipping_mark'] or ex.get('shipping_mark', ''),
                           source_file=fname, source_mtime=mtime, updated_at=now)
                upd_c += 1

    # 原子寫入：先寫 temp，再 rename，避免中途失敗導致 JSON 截斷
    import tempfile, shutil
    tmp_fd, tmp_path = tempfile.mkstemp(dir=JSON_PATH.parent, suffix='.tmp')
    try:
        with os.fdopen(tmp_fd, 'w', encoding='utf-8') as f:
            json.dump(list(existing.values()), f, ensure_ascii=False, indent=2)
        shutil.move(tmp_path, str(JSON_PATH))
    except Exception:
        try: os.unlink(tmp_path)
        except: pass
        raise
    SCAN_META.write_text(json.dumps({'last_scan_mtime': time.time(), 'last_scan_at': now}), encoding='utf-8')
    print(f"新增 {new_c}，更新 {upd_c}，共 {len(existing)} 筆（{time.time()-t0:.1f}s）")


if __name__ == '__main__':
    full = '--full' in sys.argv
    scan(full_scan=full or not SCAN_META.exists())
