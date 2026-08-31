import os
import sqlite3
import json
import io
import traceback
from flask import Flask, request, jsonify, send_file
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
DB_FILE = 'seating.db'
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

PARTY_ROLE_LABELS = {
    'host': '主方（上级）',
    'guest': '宾方（下级）',
    'unassigned': '未分类',
}

def normalize_party_role(value):
    normalized = str(value or '').strip().lower()
    if normalized in {'host', '主方', '主', '上级', '主方（上级）', '主方(上级)'}:
        return 'host'
    if normalized in {'guest', '宾方', '宾', '下级', '宾方（下级）', '宾方(下级)', '客方'}:
        return 'guest'
    return 'unassigned'

# ==========================================
# 数据库初始化
# ==========================================
def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS meetings
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, layout JSON, seats JSON, records JSON, sms_template TEXT, updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
    c.execute('''CREATE TABLE IF NOT EXISTS templates
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, config JSON)''')
    c.execute('''CREATE TABLE IF NOT EXISTS attendees
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, meeting_id INTEGER, sort_order INTEGER,
                  name TEXT, unit TEXT, category TEXT, phone TEXT, extra JSON,
                  party_role TEXT DEFAULT 'unassigned')''')
    attendee_columns = {row[1] for row in c.execute("PRAGMA table_info(attendees)").fetchall()}
    if 'party_role' not in attendee_columns:
        c.execute("ALTER TABLE attendees ADD COLUMN party_role TEXT DEFAULT 'unassigned'")
    conn.commit()
    conn.close()

# ==========================================
# 会议管理与状态同步 API
# ==========================================
@app.route('/api/meetings', methods=['GET', 'POST'])
def handle_meetings():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    if request.method == 'GET':
        c.execute("SELECT id, name, updated_at FROM meetings ORDER BY updated_at DESC")
        return jsonify([dict(r) for r in c.fetchall()])
    else:
        name = request.json.get('name', '新建会议')
        c.execute("INSERT INTO meetings (name, layout, seats, records, sms_template) VALUES (?, '{}', '{}', '{}', '(参会通知){姓名}您好，您的座位在 {排号}排 {座号}，请准时入场。')", (name,))
        meeting_id = c.lastrowid
        conn.commit()
        conn.close()
        return jsonify({"status": "success", "id": meeting_id})

@app.route('/api/meetings/<int:m_id>', methods=['GET'])
def get_meeting(m_id):
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute("SELECT * FROM meetings WHERE id=?", (m_id,))
    row = c.fetchone()
    conn.close()
    if row: return jsonify(dict(row))
    return jsonify({"error": "Meeting not found"}), 404

@app.route('/api/meetings/<int:m_id>/save', methods=['POST'])
def save_meeting_state(m_id):
    data = request.json
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("UPDATE meetings SET layout=?, seats=?, records=?, sms_template=?, updated_at=CURRENT_TIMESTAMP WHERE id=?",
              (json.dumps(data.get('layout', {})), json.dumps(data.get('seats', {})),
               json.dumps(data.get('records', {})), data.get('sms_template', ''), m_id))
    conn.commit()
    conn.close()
    return jsonify({"status": "success"})

@app.route('/api/meetings/<int:m_id>/rename', methods=['POST'])
def rename_meeting(m_id):
    name = request.json.get('name', '未命名会议')
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("UPDATE meetings SET name=? WHERE id=?", (name, m_id))
    conn.commit()
    conn.close()
    return jsonify({"status": "success"})

@app.route('/api/meetings/<int:m_id>/delete', methods=['DELETE'])
def delete_meeting(m_id):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("DELETE FROM meetings WHERE id=?", (m_id,))
    c.execute("DELETE FROM attendees WHERE meeting_id=?", (m_id,))
    conn.commit()
    conn.close()
    return jsonify({"status": "success"})

# ==========================================
# 人员管理 API
# ==========================================
@app.route('/api/meetings/<int:m_id>/attendees', methods=['GET', 'POST', 'PUT'])
def handle_attendees(m_id):
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    if request.method == 'GET':
        c.execute("SELECT * FROM attendees WHERE meeting_id=? ORDER BY sort_order ASC, id ASC", (m_id,))
        rows = []
        for r in c.fetchall():
            d = dict(r)
            d['extra'] = json.loads(d['extra']) if d['extra'] else {}
            d['party_role'] = normalize_party_role(d.get('party_role'))
            rows.append(d)
        conn.close()
        return jsonify(rows)
    elif request.method == 'POST':
        data = request.json
        c.execute("SELECT MAX(sort_order) FROM attendees WHERE meeting_id=?", (m_id,))
        max_sort = c.fetchone()[0] or 0
        c.execute("INSERT INTO attendees (meeting_id, sort_order, name, unit, category, phone, extra, party_role) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                  (m_id, max_sort + 1, data.get('name',''), data.get('unit',''), data.get('category',''), data.get('phone',''), json.dumps(data.get('extra', {})), normalize_party_role(data.get('party_role'))))
        conn.commit()
        conn.close()
        return jsonify({"status": "success"})
    elif request.method == 'PUT':
        data = request.json
        c.execute("UPDATE attendees SET name=?, unit=?, category=?, phone=?, extra=?, party_role=? WHERE id=? AND meeting_id=?",
                  (data.get('name',''), data.get('unit',''), data.get('category',''), data.get('phone',''), json.dumps(data.get('extra', {})), normalize_party_role(data.get('party_role')), data.get('id'), m_id))
        conn.commit()
        conn.close()
        return jsonify({"status": "success"})

@app.route('/api/meetings/<int:m_id>/attendees/reorder', methods=['POST'])
def reorder_attendees(m_id):
    ids = request.json.get('ids', [])
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    for idx, att_id in enumerate(ids):
        c.execute("UPDATE attendees SET sort_order=? WHERE id=? AND meeting_id=?", (idx + 1, att_id, m_id))
    conn.commit()
    conn.close()
    return jsonify({"status": "success"})

@app.route('/api/meetings/<int:m_id>/attendees/bulk_role', methods=['POST'])
def bulk_set_attendee_party_role(m_id):
    ids = request.json.get('ids', [])
    party_role = normalize_party_role(request.json.get('party_role'))
    if not ids:
        return jsonify({"status": "success", "updated": 0})
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    placeholders = ','.join('?' * len(ids))
    c.execute(
        f"UPDATE attendees SET party_role=? WHERE id IN ({placeholders}) AND meeting_id=?",
        (party_role, *ids, m_id),
    )
    updated = c.rowcount
    conn.commit()
    conn.close()
    return jsonify({"status": "success", "updated": updated, "party_role": party_role})

@app.route('/api/meetings/<int:m_id>/attendees/bulk_delete', methods=['POST'])
def bulk_delete_attendees(m_id):
    ids = request.json.get('ids', [])
    if not ids: return jsonify({"status": "success"})
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    placeholders = ','.join('?' * len(ids))
    c.execute(f"DELETE FROM attendees WHERE id IN ({placeholders}) AND meeting_id=?", (*ids, m_id))
    conn.commit()
    conn.close()
    return jsonify({"status": "success"})

@app.route('/api/meetings/<int:m_id>/upload', methods=['POST'])
def upload_attendees(m_id):
    if 'file' not in request.files: return jsonify({"error": "没有文件"}), 400
    file = request.files['file']
    try:
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        df = pd.read_csv(file) if file.filename.endswith('.csv') else pd.read_excel(file)
        df = df.fillna('')
        cols = df.columns.tolist()

        idx_col = '序号' if '序号' in cols else None
        name_col = '姓名' if '姓名' in cols else (cols[0] if len(cols)>0 else None)
        unit_col = '单位' if '单位' in cols else (cols[1] if len(cols)>1 else None)
        phone_col = '电话' if '电话' in cols else None
        cat_col = '类别' if '类别' in cols else None
        party_col = next((col for col in ['主宾', '主宾身份', '所属方', '主客', '身份'] if col in cols), None)

        if idx_col:
            df[idx_col] = pd.to_numeric(df[idx_col], errors='coerce').fillna(9999)
            df = df.sort_values(by=idx_col, kind='stable')

        c.execute("SELECT MAX(sort_order) FROM attendees WHERE meeting_id=?", (m_id,))
        current_sort = c.fetchone()[0] or 0

        c.execute("SELECT name, unit, phone FROM attendees WHERE meeting_id=?", (m_id,))
        existing_attendees = {(r[0], r[1], r[2]): True for r in c.fetchall()}

        duplicates = []
        added_count = 0

        for _, row in df.iterrows():
            if not name_col: break
            name = str(row[name_col]).strip()
            if not name: continue

            unit = str(row[unit_col]).strip() if unit_col else ''
            category = str(row[cat_col]).strip() if cat_col else ''
            phone = str(row[phone_col]).strip() if phone_col else ''
            phone = ''.join(filter(str.isdigit, phone))
            party_role = normalize_party_role(row[party_col] if party_col else '')

            duplicate_key = (name, unit, phone)
            if duplicate_key in existing_attendees:
                duplicates.append(f"{name}({unit})")
                continue

            core_columns = {'序号', '姓名', '单位', '电话', '类别'}
            if party_col:
                core_columns.add(party_col)
            extra = {col: str(row[col]) for col in cols if col not in core_columns}
            current_sort += 1

            c.execute("INSERT INTO attendees (meeting_id, sort_order, name, unit, category, phone, extra, party_role) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                      (m_id, current_sort, name, unit, category, phone, json.dumps(extra), party_role))
            existing_attendees[duplicate_key] = True
            added_count += 1

        conn.commit()
        conn.close()

        message = f"名单导入成功！已按文件顺序排序。新增 {added_count} 人。"
        if duplicates:
            message += f"\n以下人员因重复（同姓名同单位）未导入：{', '.join(duplicates[:10])}"
            if len(duplicates) > 10:
                message += f" 等{len(duplicates)}人"

        return jsonify({"status": "success", "message": message})
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

@app.route('/api/templates', methods=['GET', 'POST'])
def handle_templates():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    if request.method == 'POST':
        data = request.json
        c.execute("INSERT INTO templates (name, config) VALUES (?, ?)", (data.get('name'), json.dumps(data.get('config'))))
        conn.commit()
        conn.close()
        return jsonify({"status": "success"})
    else:
        c.execute("SELECT id, name, config FROM templates ORDER BY id DESC")
        res = [{"id": r[0], "name": r[1], "config": json.loads(r[2])} for r in c.fetchall()]
        conn.close()
        return jsonify(res)

@app.route('/api/templates/<int:t_id>', methods=['DELETE'])
def delete_template(t_id):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("DELETE FROM templates WHERE id=?", (t_id,))
    conn.commit()
    conn.close()
    return jsonify({"status": "success"})

# ==========================================
# 综合导出 API
# ==========================================
def format_custom_text(att, template):
    if not att: return ""
    txt = template.replace('{姓名}', att.get('name','')).replace('{单位}', att.get('unit',''))\
                  .replace('{类别}', att.get('category','')).replace('{电话}', att.get('phone',''))
    extra = att.get('extra', {})
    if isinstance(extra, str):
        try: extra = json.loads(extra)
        except: extra = {}
    for k, v in extra.items():
        txt = txt.replace(f'{{{k}}}', str(v))
    return txt

def format_col_label(label):
    label_str = str(label)
    if not label_str: return ''
    if label_str.endswith('座') or label_str.endswith('号'): return label_str.rstrip('座').rstrip('号')
    try:
        int(label_str)
        return label_str
    except ValueError:
        return label_str

def get_seat_col_label(s_data, conf, idx):
    custom = s_data.get('customLabel')
    if custom is not None and custom != '': return str(custom)
    labels = conf.get('colLabels', [])
    return format_col_label(labels[idx] if idx < len(labels) else str(idx + 1))

def get_contrast_color(hexcolor):
    if not hexcolor: return 'FFFFFFFF'
    r = int(hexcolor[0:2], 16)
    g = int(hexcolor[2:4], 16)
    b = int(hexcolor[4:6], 16)
    yiq = ((r*299)+(g*587)+(b*114))/1000
    return 'FF000000' if yiq >= 128 else 'FFFFFFFF'

@app.route('/api/export', methods=['POST'])
def export_excel():
    try:
        data = request.json
        conf_name = data.get('conf_name', '未命名会议')
        layout_type = data.get('layoutType', 'classroom')
        podium = data.get('podium', {})
        main = data.get('main', {})
        longtable_config = data.get('longtableConfig', {})
        roundtable_config = data.get('roundtableConfig', {})
        utable_config = data.get('utableConfig', {})
        seats_data = data.get('seats_data', {})
        sms_template = data.get('sms_template', '')
        export_template = data.get('export_template', '{姓名}\n{单位}')
        region_types = data.get('region_types', [])

        layout_type_name = {
            'classroom': '课桌式会议',
            'longtable': '长会议桌式会议',
            'roundtable': '圆桌式会议',
            'utable': 'U型桌式会议'
        }.get(layout_type, '会议排座')

        def parse_seat_key(key):
            """解析座位键，返回适合导出和短信使用的 (区域, 座位) 文案。"""
            if key.startswith('lt-'):
                parts = key.split('-')
                if len(parts) == 4 and parts[1] == 'head':
                    side_name = '左侧' if parts[2] == 'left' else '右侧'
                    return f'桌外主位-{side_name}', f'第{int(parts[3]) + 1}座'
                if len(parts) >= 4:
                    side = parts[1]
                    row = int(parts[-2]) + 1
                    pos = int(parts[-1]) + 1
                    side_names = {'top': '上侧', 'bottom': '下侧', 'left': '左侧', 'right': '右侧'}
                    side_config = longtable_config.get('sides', {}).get(side, {})
                    zone = side_config.get('label') or side_names.get(side, side)
                    return zone, f'第{row}排 第{pos}座'
            elif key.startswith('rt-'):
                parts = key.split('-')
                if len(parts) >= 4:
                    table_no = int(parts[1]) + 1
                    ring_id = '-'.join(parts[2:-1])
                    raw_pos = int(parts[-1])
                    pos = raw_pos + 1
                    ring_label = ring_id
                    for ring in roundtable_config.get('rings', []):
                        if str(ring.get('id')) == ring_id:
                            ring_label = ring.get('label') or ring_id
                            break
                    rings = roundtable_config.get('rings', [])
                    is_main_ring = bool(rings) and str(rings[0].get('id')) == ring_id
                    if roundtable_config.get('dualPrincipal') and table_no == 1 and is_main_ring and raw_pos in (0, 1):
                        return f'第{table_no}桌-{ring_label}', '主人主位' if raw_pos == 0 else '宾客主位'
                    return f'第{table_no}桌-{ring_label}', f'第{pos}座'
            elif key.startswith('u-'):
                parts = key.split('-')
                if len(parts) in (3, 4):
                    zone = parts[1]
                    if zone == 'leader':
                        raw_pos = int(parts[2])
                        count = max(int(utable_config.get('leaderSeatCount', 1) or 1), 1)
                        center_left = (count - 1) // 2
                        order = [center_left]
                        offset = 1
                        while len(order) < count:
                            right = center_left + offset
                            left = center_left - offset
                            if right < count: order.append(right)
                            if left >= 0: order.append(left)
                            offset += 1
                        order_no = order.index(raw_pos) + 1 if raw_pos in order else raw_pos + 1
                        return '领导席', f'第{order_no}位'
                    column = int(parts[2]) if len(parts) == 4 else 0
                    raw_pos = int(parts[3]) if len(parts) == 4 else int(parts[2])
                    side_name = '左侧参会席' if zone == 'left' else '右侧参会席'
                    if len(parts) == 3:
                        return side_name, f'距领导第{raw_pos + 1}位'
                    column_label = f'第{column + 1}列' + ('（内侧）' if column == 0 else '')
                    return f'{side_name}-{column_label}', f'距领导第{raw_pos + 1}位'
            else:
                parts = key.split('-')
                if len(parts) >= 3:
                    return parts[0], f"{parts[1]}-{parts[2]}"
            return key, ''

        def seat_sort_key(key):
            return tuple((0, int(part)) if part.isdigit() else (1, part) for part in key.split('-'))

        def get_label(conf, axis, idx):
            labels = conf.get(f'{axis}Labels', [])
            return str(labels[idx]) if idx < len(labels) else str(idx + 1)

        wb = Workbook()
        ws1 = wb.active
        ws1.title = "会场布局图"
        fill_aisle = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
        fill_disabled = PatternFill(start_color="9E9E9E", end_color="9E9E9E", fill_type="solid")

        podium_color = podium.get('color', '#f97316').replace('#', 'FF')
        main_color = main.get('color', '#22c55e').replace('#', 'FF')
        podium_text_color = podium.get('textColor', '#ffffff').replace('#', 'FF')
        main_text_color = main.get('textColor', '#ffffff').replace('#', 'FF')

        fill_podium_seat = PatternFill(start_color=podium_color, end_color=podium_color, fill_type="solid")
        fill_seat = PatternFill(start_color=main_color, end_color=main_color, fill_type="solid")

        align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
        border_thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        max_cols = max(podium.get('cols', 0), main.get('cols', 0)) if layout_type == 'classroom' else 4
        offset_p = max(0, (max_cols - podium.get('cols', 0)) // 2) + 2 if podium.get('cols', 0) > 0 else 2
        offset_m = max(0, (max_cols - main.get('cols', 0)) // 2) + 2 if main.get('cols', 0) > 0 else 2

        for c in range(max_cols + 2): ws1.column_dimensions[get_column_letter(c+1)].width = 15

        if layout_type == 'classroom':
            ws1.merge_cells(start_row=1, start_column=1, end_row=2, end_column=max_cols+2)
            cell_title = ws1.cell(row=1, column=1, value=conf_name)
            cell_title.font = Font(size=20, bold=True)
            cell_title.alignment = align_center
        current_row = 4

        if layout_type == 'classroom' and podium.get('rows', 0) > 0:
            ws1.merge_cells(start_row=current_row, start_column=offset_p, end_row=current_row, end_column=offset_p + podium['cols'] - 1)
            ws1.cell(row=current_row, column=offset_p, value="主 席 台").font = Font(bold=True, color=podium_color)
            ws1.cell(row=current_row, column=offset_p).alignment = align_center
            current_row += 1

            for c in range(podium['cols']):
                ws1.cell(row=current_row, column=offset_p + c, value=get_label(podium, 'col', c)).font = Font(bold=True)
                ws1.cell(row=current_row, column=offset_p + c).alignment = align_center
            current_row += 1

            for r in range(podium['rows']):
                ws1.row_dimensions[current_row].height = 50
                r_label = get_label(podium, 'row', r)
                ws1.cell(row=current_row, column=offset_p - 1, value=r_label).font = Font(bold=True)
                for c in range(podium['cols']):
                    cell = ws1.cell(row=current_row, column=offset_p + c)
                    cell.border = border_thin
                    cell.alignment = align_center
                    s_data = seats_data.get(f"p-{r}-{c}", {})
                    c_label = get_seat_col_label(s_data, podium, c)
                    if s_data.get('type') == 'aisle':
                        cell.fill = fill_aisle; cell.value = "走廊"
                    elif s_data.get('type') == 'disabled':
                        cell.fill = fill_disabled; cell.value = "不可用"
                    else:
                        att = s_data.get('attendee')
                        cell.fill = fill_podium_seat; cell.font = Font(color=podium_text_color)
                        if att: cell.value = format_custom_text(att, export_template)
                        else: cell.value = f"{r_label}排\n{c_label}"
                current_row += 1
            current_row += 2

        if layout_type == 'classroom' and main.get('rows', 0) > 0:
            ws1.merge_cells(start_row=current_row, start_column=offset_m, end_row=current_row, end_column=offset_m + main['cols'] - 1)
            ws1.cell(row=current_row, column=offset_m, value="主 会 场").font = Font(bold=True, color=main_color)
            ws1.cell(row=current_row, column=offset_m).alignment = align_center
            current_row += 1

            for c in range(main['cols']):
                ws1.cell(row=current_row, column=offset_m + c, value=get_label(main, 'col', c)).font = Font(bold=True)
                ws1.cell(row=current_row, column=offset_m + c).alignment = align_center
            current_row += 1

            for r in range(main['rows']):
                ws1.row_dimensions[current_row].height = 50
                r_label = get_label(main, 'row', r)
                ws1.cell(row=current_row, column=offset_m - 1, value=r_label).font = Font(bold=True)
                for c in range(main['cols']):
                    cell = ws1.cell(row=current_row, column=offset_m + c)
                    cell.border = border_thin
                    cell.alignment = align_center
                    s_data = seats_data.get(f"m-{r}-{c}", {})
                    c_label = get_seat_col_label(s_data, main, c)
                    if s_data.get('type') == 'aisle':
                        cell.fill = fill_aisle; cell.value = "走廊"
                    elif s_data.get('type') == 'disabled':
                        cell.fill = fill_disabled; cell.value = "不可用"
                    else:
                        att = s_data.get('attendee')
                        region_name = s_data.get('region_name', '未分区')
                        region_color = main_color
                        region_text_color = main_text_color
                        if region_types and region_name != '未分区':
                            for reg in region_types:
                                if reg.get('name') == region_name:
                                    region_color = reg.get('color', main_color).replace('#', 'FF')
                                    region_text_color = get_contrast_color(region_color[2:])
                                    break
                        cell.fill = PatternFill(start_color=region_color, end_color=region_color, fill_type="solid")
                        cell.font = Font(color=region_text_color)
                        if att: cell.value = format_custom_text(att, export_template)
                        else: cell.value = f"{r_label}排\n{c_label}"
                current_row += 1

        if layout_type != 'classroom':
            # 长会议桌/圆桌/U型桌使用分组座位清单导出，位置名称与短信变量保持一致。
            ws1.merge_cells(start_row=1, start_column=1, end_row=2, end_column=6)
            cell_title = ws1.cell(row=1, column=1, value=conf_name + ' - ' + layout_type_name)
            cell_title.font = Font(size=20, bold=True)
            cell_title.alignment = align_center
            current_row = 4

            headers = ['区域', '位置', '姓名', '单位', '类别', '电话']
            for i, h in enumerate(headers):
                ws1.cell(row=current_row, column=i+1, value=h).font = Font(bold=True)
                ws1.cell(row=current_row, column=i+1).alignment = align_center
                ws1.column_dimensions[get_column_letter(i+1)].width = 18
            current_row += 1

            sorted_keys = sorted(seats_data.keys(), key=seat_sort_key)
            for key in sorted_keys:
                s_data = seats_data.get(key, {})
                zone, pos_label = parse_seat_key(key)
                att = s_data.get('attendee')
                if att:
                    row_data = [zone, pos_label, att.get('name',''), att.get('unit',''), att.get('category',''), att.get('phone','')]
                else:
                    row_data = [zone, pos_label, '(空位)', '', '', '']
                ws1.append(row_data)
                ws1.cell(row=current_row, column=1).alignment = align_center
                ws1.cell(row=current_row, column=2).alignment = align_center
                current_row += 1

        # --- Sheet 2: 签到表 ---
        ws2 = wb.create_sheet(title="参会人员签到表")

        base_headers = ["大区", "块区", "排号", "座/列号", "姓名", "单位", "类别", "电话"]
        extra_keys = set()
        for key, s_data in seats_data.items():
            att = s_data.get('attendee')
            if att:
                extra = att.get('extra', {})
                if isinstance(extra, str):
                    try: extra = json.loads(extra)
                    except: extra = {}
                if isinstance(extra, dict):
                    extra_keys.update(extra.keys())
        extra_keys_list = sorted(list(extra_keys))
        headers = base_headers + extra_keys_list + ["签到"]
        ws2.append(headers)
        for col in range(1, len(headers)+1):
            ws2.cell(row=1, column=col).font = Font(bold=True)
            ws2.column_dimensions[get_column_letter(col)].width = 15

        for key, s_data in seats_data.items():
            att = s_data.get('attendee')
            if att:
                zone, pos_label = parse_seat_key(key)
                if zone == 'p':
                    r = int(key.split('-')[1]); c = int(key.split('-')[2])
                    r_label = get_label(podium, 'row', r); c_label = get_seat_col_label(s_data, podium, c); zone_name = "主席台"
                elif zone == 'm':
                    r = int(key.split('-')[1]); c = int(key.split('-')[2])
                    r_label = get_label(main, 'row', r); c_label = get_seat_col_label(s_data, main, c); zone_name = "主会场"
                else:
                    r_label = zone; c_label = pos_label; zone_name = layout_type_name
                phone = att.get('phone', '')
                phone_str = str(phone) if phone else ''
                phone_num = ''.join(filter(str.isdigit, phone_str))

                extra = att.get('extra', {})
                if isinstance(extra, str):
                    try: extra = json.loads(extra)
                    except: extra = {}
                extra_values = []
                for ek in extra_keys_list:
                    extra_values.append(str(extra.get(ek, '')) if isinstance(extra, dict) else '')

                row_data = [zone_name, s_data.get('region_name', '未分区'), r_label, c_label,
                            att.get('name',''), att.get('unit',''), att.get('category',''), phone_num if phone_num else '']
                row_data.extend(extra_values)
                row_data.append("")
                ws2.append(row_data)

        # --- Sheet 3: 短信清单 ---
        ws3 = wb.create_sheet(title="短信群发清单")
        ws3.append(["姓名", "单位", "电话", "短信内容"])
        for col in range(1, 5): ws3.cell(row=1, column=col).font = Font(bold=True)
        ws3.column_dimensions['A'].width = 15; ws3.column_dimensions['B'].width = 25
        ws3.column_dimensions['C'].width = 15; ws3.column_dimensions['D'].width = 80

        for key, s_data in seats_data.items():
            att = s_data.get('attendee')
            if att:
                zone, pos_label = parse_seat_key(key)
                if zone == 'p':
                    r, c = int(key.split('-')[1]), int(key.split('-')[2])
                    r_label = get_label(podium, 'row', r); c_label = get_seat_col_label(s_data, podium, c)
                elif zone == 'm':
                    r, c = int(key.split('-')[1]), int(key.split('-')[2])
                    r_label = get_label(main, 'row', r); c_label = get_seat_col_label(s_data, main, c)
                else:
                    r_label = zone; c_label = pos_label

                phone = att.get('phone', '')
                phone_str = str(phone) if phone else ''
                phone_num = ''.join(filter(str.isdigit, phone_str))

                msg = sms_template
                msg = msg.replace('{姓名}', att.get('name','')).replace('{单位}', att.get('unit',''))\
                         .replace('{类别}', att.get('category','')).replace('{电话}', phone_num if phone_num else '')\
                         .replace('{排号}', r_label).replace('{座号}', c_label)
                extra = att.get('extra', {})
                if isinstance(extra, str):
                    try: extra = json.loads(extra)
                    except: extra = {}
                for k, v in extra.items():
                    msg = msg.replace(f'{{{k}}}', str(v))

                ws3.append([att.get('name',''), att.get('unit',''), phone_num if phone_num else '', msg])

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return send_file(output, as_attachment=True, download_name=f'{conf_name}_排座导出.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

# ==========================================
# 前端 HTML & Vue 3 SPA (防转义与沙盒保护)
# ==========================================
HTML_TEMPLATE = r"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>智能会场排座系统</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://unpkg.com/vue@3/dist/vue.global.js"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        [v-cloak] { display: none !important; }
        body { margin: 0; overflow: hidden; background-color: #f3f4f6; }
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-track { background: #f1f1f1; border-radius: 4px; }
        ::-webkit-scrollbar-thumb { background: #c1c1c1; border-radius: 4px; }
        ::-webkit-scrollbar-thumb:hover { background: #a8a8a8; }

        #canvas-container { overflow: auto; }
        .venue-wrapper { padding: 40px; display: grid; place-items: center; min-width: max-content; }
        .classroom-layout { display: flex; flex-direction: column; align-items: center; width: max-content; }

        .grid-container {
            display: grid; gap: 6px; background: white; padding: 20px; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.05); margin-bottom: 30px;
        }

        .seat {
            width: 85px; height: 75px; border-radius: 6px; display: flex; flex-direction: column;
            align-items: center; justify-content: center; cursor: pointer; transition: all 0.15s;
            border: 1px solid #e5e7eb; position: relative; user-select: none; background-color: #f9fafb;
        }
        .seat:hover { filter: brightness(0.95); transform: scale(1.05); z-index: 10; box-shadow: 0 4px 8px rgba(0,0,0,0.1); }
        .seat-occupied { border: none; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .seat-aisle { background-color: transparent !important; border: 1px dashed #cbd5e1; color: transparent; cursor: default;}
        .seat-aisle:hover { transform: none; box-shadow: none; }
        .seat-disabled { background-color: #cbd5e1 !important; color: white; border: none; cursor: not-allowed; }

        .axis-input { width: 100%; height: 100%; text-align: center; background: transparent; border: 1px solid transparent; font-weight: bold; color: #64748b; font-size: 13px; border-radius: 4px; cursor: text; }
        .axis-input:hover { background: #f1f5f9; border-color: #cbd5e1; }
        .axis-input:focus { outline: none; background: white; border-color: #3b82f6; box-shadow: 0 0 0 2px rgba(59,130,246,0.2); }
        .header-cell { display: flex; align-items: center; justify-content: center; background: #f8fafc; border-radius: 4px; cursor: context-menu;}
        .header-cell:hover { background: #e2e8f0; }

        .region-overlay { position: absolute; top: 0; left: 0; right: 0; bottom: 0; border-radius: 6px; opacity: 0.25; pointer-events: none; border: 2px solid transparent; }

        .drag-over { border-top: 2px solid #4f46e5 !important; background-color: #eef2ff !important; }
        .meeting-title-input { background: transparent; border: 1px dashed transparent; text-align: center; width: 100%; }
        .meeting-title-input:hover { border-color: #cbd5e1; }
        .meeting-title-input:focus { outline: none; border-color: #4f46e5; border-style: solid; background: white; border-radius: 8px;}

        /* ========== 长会议桌式布局样式（内部沿用 longtable 类名兼容旧数据） ========== */
        .longtable-wrapper { display: flex; flex-direction: column; justify-content: center; align-items: center; min-height: 500px; padding: 40px; }
        .longtable-scene { display: grid; gap: 8px; align-items: center; justify-items: center; }
        /* 桌子 */
        .longtable-table-wrap { display: flex; align-items: center; justify-content: center; }
        .longtable-table { background: var(--table-color, #8B4513); border-radius: 8px; box-shadow: 0 6px 20px rgba(0,0,0,0.25); width: 100%; height: 100%; display: flex; align-items: center; justify-content: center; }
        .longtable-table-label { color: rgba(255,255,255,0.7); font-size: 13px; font-weight: 700; letter-spacing: 2px; pointer-events: none; }
        /* 边容器 - 横边(top/bottom) */
        .longtable-side-h { display: flex; flex-direction: column; align-items: center; gap: 4px; }
        .longtable-side-h .longtable-side-label { font-weight: 700; font-size: 11px; color: #374151; background: #f3f4f6; padding: 2px 10px; border-radius: 4px; white-space: nowrap; }
        .longtable-side-h .longtable-row { display: flex; gap: 4px; justify-content: center; flex-wrap: nowrap; }
        .longtable-side-h .longtable-row-label { font-size: 9px; font-weight: 600; color: #9ca3af; padding: 1px 4px; text-align: center; }
        /* 边容器 - 竖边(left/right) */
        .longtable-side-v { display: flex; flex-direction: row; align-items: center; gap: 4px; }
        .longtable-side-v .longtable-side-label { font-weight: 700; font-size: 11px; color: #374151; background: #f3f4f6; padding: 4px 6px; border-radius: 4px; white-space: nowrap; writing-mode: vertical-lr; }
        .longtable-side-v .longtable-col { display: flex; flex-direction: column; gap: 4px; align-items: center; }
        .longtable-side-v .longtable-row-label { font-size: 9px; font-weight: 600; color: #9ca3af; padding: 1px 4px; text-align: center; writing-mode: vertical-lr; }
        /* 单个座位 */
        .longtable-seat { width: 74px; height: 74px; border: 2px solid #d1d5db; border-radius: 6px; background: #f9fafb; display: flex; align-items: center; justify-content: center; cursor: pointer; transition: all 0.15s; overflow: hidden; flex-shrink: 0; }
        .longtable-seat:hover { border-color: #6366f1; box-shadow: 0 2px 8px rgba(99,102,241,0.3); z-index: 10; }
        .longtable-seat.seat-aisle { background: #e5e7eb; border-style: dashed; }
        .longtable-seat.seat-disabled { background: #9ca3af; opacity: 0.5; }
        .longtable-seat.seat-occupied { border-color: #6366f1; }
        .longtable-seat .seat-empty-label { display: flex; align-items: center; justify-content: center; width: 100%; height: 100%; }
        .longtable-seat .seat-label-input { width: 80%; text-align: center; border: none; background: transparent; font-size: 11px; font-weight: 700; color: #9ca3af; outline: none; }
        .longtable-seat .seat-label-input:hover, .longtable-seat .seat-label-input:focus { color: #4b5563; background: rgba(0,0,0,0.05); border-radius: 3px; }
        .longtable-seat .seat-occupied-content { display: flex; flex-direction: column; align-items: center; justify-content: center; text-align: center; width: 100%; padding: 2px; }
        /* 桌外主位 */
        .longtable-head-row { display: flex; gap: 10px; align-items: center; }
        .longtable-head-group { display: flex; flex-direction: column; align-items: center; gap: 3px; }
        .longtable-head-group .head-label { font-size: 10px; font-weight: 700; color: #6b7280; background: #f3f4f6; padding: 2px 8px; border-radius: 3px; white-space: nowrap; }

        /* ========== 圆桌式布局样式 ========== */
        .roundtable-wrapper { display: flex; justify-content: center; align-items: flex-start; flex-wrap: wrap; gap: 40px; min-height: 500px; padding: 40px; }
        .roundtable-scene { position: relative; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
        .roundtable-table { position: absolute; border-radius: 50%; background: var(--table-color, #8B4513); box-shadow: 0 4px 16px rgba(0,0,0,0.3); display: flex; align-items: center; justify-content: center; z-index: 1; }
        .roundtable-seat { position: absolute; width: 68px; height: 68px; border: 2px solid #d1d5db; border-radius: 50%; background: #f9fafb; display: flex; align-items: center; justify-content: center; cursor: pointer; transition: all 0.15s; overflow: hidden; transform: translate(-50%, -50%); }
        .roundtable-seat:hover { border-color: #6366f1; box-shadow: 0 2px 8px rgba(99,102,241,0.3); z-index: 20; }
        .roundtable-seat.seat-aisle { background: #e5e7eb; border-style: dashed; }
        .roundtable-seat.seat-disabled { background: #9ca3af; opacity: 0.5; }
        .roundtable-seat.seat-occupied { border-color: #6366f1; }
        .roundtable-seat.roundtable-dual-principal { border-color: #f59e0b; box-shadow: 0 0 0 4px rgba(245,158,11,0.18), 0 3px 10px rgba(146,64,14,0.24); }
        .roundtable-seat.roundtable-dual-principal::after { content: '主位'; position: absolute; top: -1px; right: 5px; font-size: 7px; line-height: 12px; color: #92400e; background: #fef3c7; padding: 0 3px; border-radius: 0 0 4px 4px; }
        .roundtable-seat .seat-empty-label { font-size: 11px; color: #9ca3af; font-weight: 700; }
        .roundtable-seat .seat-occupied-content { display: flex; flex-direction: column; align-items: center; justify-content: center; text-align: center; width: 100%; padding: 1px; }

        /* ========== U 型桌式布局样式 ========== */
        .utable-wrapper { display: flex; justify-content: center; align-items: center; min-height: 500px; padding: 40px; }
        .utable-scene { display: grid; gap: 12px; align-items: center; justify-items: center; --u-accent: #475569; }
        .utable-leader-strip { display: flex; justify-content: center; align-items: center; width: 100%; }
        .utable-side-group { display: flex; align-items: stretch; justify-content: center; height: 100%; }
        .utable-side-group.left-side { flex-direction: row-reverse; }
        .utable-side-strip { display: flex; flex-direction: column; justify-content: flex-start; align-items: center; height: 100%; }
        .utable-side-strip.leader-at-bottom { flex-direction: column-reverse; }
        .utable-open-center { width: 100%; height: 100%; min-width: 240px; min-height: 240px; border: 3px dashed #cbd5e1; border-top-color: var(--u-accent); border-left-color: var(--u-accent); border-right-color: var(--u-accent); border-bottom-color: transparent; border-radius: 12px 12px 0 0; display: flex; align-items: center; justify-content: center; color: #94a3b8; font-size: 12px; font-weight: 700; letter-spacing: 2px; }
        .utable-scene.leader-bottom .utable-open-center { border-top-color: transparent; border-bottom-color: var(--u-accent); border-radius: 0 0 12px 12px; }
        .utable-stage { width: 100%; min-height: 70px; border: 2px solid #94a3b8; border-radius: 10px; background: linear-gradient(135deg, #f8fafc, #e2e8f0); color: #475569; display: flex; flex-direction: column; align-items: center; justify-content: center; font-weight: 800; letter-spacing: 2px; box-shadow: 0 3px 10px rgba(15,23,42,0.08); }
        .utable-stage i { font-size: 18px; margin-bottom: 4px; }
        .utable-side-title { color: #64748b; background: #f1f5f9; border-radius: 999px; padding: 2px 8px; font-size: 9px; font-weight: 700; white-space: nowrap; }
        .utable-leader-title { color: #92400e; background: #fef3c7; border-radius: 999px; padding: 2px 10px; font-size: 10px; font-weight: 800; white-space: nowrap; }
        .utable-seat { width: 74px; height: 74px; border: 2px solid #d1d5db; border-radius: 8px; background: #f9fafb; display: flex; align-items: center; justify-content: center; cursor: pointer; transition: all 0.15s; overflow: hidden; flex-shrink: 0; position: relative; }
        .utable-seat:hover { border-color: #6366f1; box-shadow: 0 2px 8px rgba(99,102,241,0.3); z-index: 10; }
        .utable-seat.seat-aisle { background: #e5e7eb; border-style: dashed; }
        .utable-seat.seat-disabled { background: #9ca3af; opacity: 0.5; }
        .utable-seat.seat-occupied { border-color: #6366f1; }
        .utable-seat .seat-empty-label { color: #9ca3af; font-size: 11px; font-weight: 700; }
        .utable-seat .seat-occupied-content { display: flex; flex-direction: column; align-items: center; justify-content: center; text-align: center; width: 100%; padding: 2px; }
    </style>
</head>
<body class="font-sans text-gray-800">
    <div id="app" class="flex flex-col h-screen" v-cloak @click="closeContextMenus">

        <!-- 导航栏 -->
        <div class="h-14 bg-indigo-900 text-white flex items-center px-6 shadow-md z-20 shrink-0">
            <div class="font-bold text-xl mr-8 flex items-center tracking-wider"><i class="fa-solid fa-layer-group mr-2"></i> 智能会场排座系统</div>

            <!-- 会议选择器 -->
            <div class="flex items-center bg-indigo-800 rounded px-2 mr-6 border border-indigo-700">
                <i class="fa-solid fa-briefcase text-indigo-300 mr-2"></i>
                <select v-model="currentMeetingId" @change="loadMeetingState" class="bg-transparent text-white font-bold py-1.5 focus:outline-none text-sm cursor-pointer w-48">
                    <option value="" disabled class="text-gray-500">请选择或新建会议...</option>
                    <option v-for="m in meetings" :key="m.id" :value="m.id" class="text-gray-800">{{m.name}}</option>
                </select>
                <button @click="createMeeting" class="ml-2 text-indigo-300 hover:text-white" title="新建会议"><i class="fa-solid fa-circle-plus"></i></button>
                <button v-if="currentMeetingId" @click="deleteCurrentMeeting" class="ml-2 text-red-400 hover:text-red-300" title="永久删除当前会议"><i class="fa-solid fa-trash-can"></i></button>
            </div>

            <div class="flex space-x-1" v-if="currentMeetingId">
                <button v-for="tab in tabs" :key="tab.id" @click="activeTab = tab.id"
                        :class="['px-4 py-2 rounded-md font-medium transition text-sm', activeTab === tab.id ? 'bg-indigo-600 text-white shadow-sm' : 'text-indigo-200 hover:bg-indigo-700']">
                    <i :class="tab.icon + ' mr-1'"></i> {{ tab.name }}
                </button>
            </div>

            <div class="ml-auto flex items-center space-x-4">
                <span class="text-xs text-indigo-300" v-if="saveStatus"><i class="fa-solid fa-cloud-arrow-up mr-1"></i> {{saveStatus}}</span>
                <button @click="undo" v-if="history.length > 0" class="text-sm bg-indigo-800 px-3 py-1 rounded-full text-indigo-200 hover:bg-indigo-700">
                    <i class="fa-solid fa-rotate-left mr-1"></i> 撤销 ({{history.length}})
                </button>
            </div>
        </div>

        <!-- 缺省页 -->
        <div v-if="!currentMeetingId" class="flex-1 flex items-center justify-center bg-gray-100">
            <div class="text-center">
                <i class="fa-solid fa-folder-open text-6xl text-gray-300 mb-4"></i>
                <h2 class="text-2xl font-bold text-gray-500 mb-2">欢迎使用会场排座系统</h2>
                <p class="text-gray-400 mb-6">所有数据按“会议”独立存储，不怕丢失。</p>
                <button @click="createMeeting" class="bg-indigo-600 text-white px-6 py-3 rounded-lg shadow-lg font-bold hover:bg-indigo-700 transition">立即新建一场会议</button>
            </div>
        </div>

        <div v-else class="flex flex-1 overflow-hidden">
            <!-- 左侧控制面板 -->
            <div class="w-[340px] bg-white border-r shadow-lg z-10 flex flex-col shrink-0">

                <!-- Tab 1: 会场设置 -->
                <div v-show="activeTab === 'layout'" class="p-5 flex flex-col h-full overflow-y-auto">

                    <h2 class="text-sm font-bold border-l-4 border-purple-500 pl-2 mb-3 bg-purple-50 py-1">🏛️ 布局类型</h2>
                    <select v-model="layout.layoutType" @focus="prevLayoutType = layout.layoutType" @change="onLayoutTypeChange" class="w-full border-2 border-purple-200 rounded p-2.5 text-sm font-bold text-purple-700 focus:ring-2 focus:ring-purple-500 outline-none bg-purple-50 mb-4">
                        <option value="classroom">🏫 课桌式（开大会）</option>
                        <option value="longtable">▭ 长会议桌式（对称围坐）</option>
                        <option value="roundtable">⭕ 圆桌式（围坐）</option>
                        <option value="utable">∪ U型桌式（领导席＋两侧参会席）</option>
                    </select>

                    <!-- 课桌式配置 -->
                    <div v-if="layout.layoutType === 'classroom'">

                    <div class="flex justify-between items-center mb-2">
                        <h2 class="text-sm font-bold border-l-4 border-orange-500 pl-2 bg-orange-50 py-1 flex-1">👑 主席台 (独立)</h2>
                        <div class="flex items-center space-x-1 ml-2">
                            <input type="color" v-model="layout.podium.color" class="w-6 h-6 p-0 border-0 rounded cursor-pointer" title="座席底色">
                        </div>
                    </div>
                    <div class="grid grid-cols-2 gap-3 mb-1">
                        <div><label class="text-[10px] text-gray-500 uppercase font-bold">排数 (自下而上增加)</label><input type="number" v-model.number="layout.podium.rows" class="w-full border rounded p-1.5 text-sm" @change="initSeats(true)"></div>
                        <div><label class="text-[10px] text-gray-500 uppercase font-bold">列数</label><input type="number" v-model.number="layout.podium.cols" class="w-full border rounded p-1.5 text-sm" @change="initSeats(true)"></div>
                    </div>
                    <p class="text-[9px] text-gray-400 mb-4">提示：主席台行号默认从下往上编号，1排最靠观众。</p>

                    <div class="flex justify-between items-center mb-2">
                        <h2 class="text-sm font-bold border-l-4 border-green-500 pl-2 bg-green-50 py-1 flex-1">👥 主会场 (全局)</h2>
                        <div class="flex items-center space-x-1 ml-2">
                            <input type="color" v-model="layout.main.color" class="w-6 h-6 p-0 border-0 rounded cursor-pointer" title="座席底色">
                        </div>
                    </div>
                    <div class="grid grid-cols-2 gap-3 mb-2">
                        <div><label class="text-[10px] text-gray-500 uppercase font-bold">总排数 (自上而下递增)</label><input type="number" v-model.number="layout.main.rows" class="w-full border rounded p-1.5 text-sm" @change="initSeats(true)"></div>
                        <div><label class="text-[10px] text-gray-500 uppercase font-bold">总列数</label><input type="number" v-model.number="layout.main.cols" class="w-full border rounded p-1.5 text-sm" @change="initSeats(true)"></div>
                    </div>
                    </div>

                    <!-- 长会议桌式配置 -->
                    <div v-if="layout.layoutType === 'longtable'" class="space-y-3">
                        <div class="flex items-center space-x-2">
                            <label class="text-[10px] text-gray-500 uppercase font-bold">桌子颜色</label>
                            <input type="color" v-model="layout.longtableConfig.tableColor" class="w-8 h-8 p-0 border rounded cursor-pointer">
                        </div>

                        <!-- 四边配置 -->
                        <div class="text-xs font-bold text-gray-500 border-b pb-1">📐 四边座位配置 <span class="text-[9px] font-normal text-gray-400">(座位数可为 0，桌边至少保持 3 座宽度)</span></div>
                        <div v-for="(sideCfg, sideKey) in layout.longtableConfig.sides" :key="sideKey" class="border p-3 rounded-lg bg-gray-50">
                            <div class="flex justify-between items-center mb-2">
                                <span class="text-xs font-bold" :class="{'text-blue-700':sideKey==='top','text-green-700':sideKey==='bottom','text-orange-700':sideKey==='left','text-purple-700':sideKey==='right'}">
                                    {{ sideKey === 'top' ? '🔝 上侧' : sideKey === 'bottom' ? '🔽 下侧' : sideKey === 'left' ? '⬅ 左侧' : '➡ 右侧' }}
                                </span>
                            </div>
                            <input type="text" v-model="sideCfg.label" class="w-full border rounded p-1.5 text-xs mb-2" placeholder="标签">
                            <div class="grid grid-cols-3 gap-2">
                                <div><label class="text-[9px] text-gray-500">排数</label><input type="number" v-model.number="sideCfg.rows" class="w-full border rounded p-1.5 text-sm" @change="initSeats()" min="0" max="5"></div>
                                <div><label class="text-[9px] text-gray-500">每排座位数</label><input type="number" v-model.number="sideCfg.seatsPerRow" class="w-full border rounded p-1.5 text-sm" @change="initSeats()" min="0" max="20"></div>
                                <div><label class="text-[9px] text-gray-500">座间距(px)</label><input type="number" v-model.number="sideCfg.seatGap" class="w-full border rounded p-1.5 text-sm" @change="initSeats()" min="0" max="800" step="2"></div>
                            </div>
                        </div>

                        <div class="text-xs font-bold text-gray-500 border-b pb-1 mt-2">🤝 主宾对坐设置</div>
                        <div class="grid grid-cols-2 gap-2 border p-3 rounded-lg bg-amber-50">
                            <div>
                                <label class="text-[9px] text-amber-700 font-bold">主方（上级）所在边</label>
                                <select :value="layout.longtableConfig.hostSide" @change="setLongtablePartySide('host', $event.target.value)" class="w-full border rounded p-1.5 text-xs bg-white">
                                    <option value="top">上侧</option>
                                    <option value="bottom">下侧</option>
                                </select>
                            </div>
                            <div>
                                <label class="text-[9px] text-blue-700 font-bold">宾方（下级）所在边</label>
                                <select :value="layout.longtableConfig.guestSide" @change="setLongtablePartySide('guest', $event.target.value)" class="w-full border rounded p-1.5 text-xs bg-white">
                                    <option value="top">上侧</option>
                                    <option value="bottom">下侧</option>
                                </select>
                            </div>
                            <div class="col-span-2 pt-2 border-t border-amber-200 text-[9px] leading-relaxed text-amber-800">
                                双方在各自的上／下侧座位不足时，都会同时向左、右两侧延伸。靠上方的一方从两侧上端向下排，靠下方的一方从两侧下端向上排。
                            </div>
                        </div>

                        <!-- 桌外主位 -->
                        <div class="text-xs font-bold text-gray-500 border-b pb-1 mt-2">👤 桌外主位（可选）</div>
                        <div class="flex items-center space-x-2">
                            <input type="checkbox" v-model="layout.longtableConfig.headSeats.enabled" @change="initSeats()" class="w-4 h-4">
                            <label class="text-xs text-gray-600">增加左右桌外主位</label>
                            <input v-if="layout.longtableConfig.headSeats.enabled" type="number" v-model.number="layout.longtableConfig.headSeats.seatCount" class="w-16 border rounded p-1 text-xs" @change="initSeats()" min="1" max="3">
                        </div>
                    </div>

                    <!-- 圆桌式配置 -->
                    <div v-if="layout.layoutType === 'roundtable'" class="space-y-3">
                        <div class="flex items-center space-x-2">
                            <label class="text-[10px] text-gray-500 uppercase font-bold">桌子颜色</label>
                            <input type="color" v-model="layout.roundtableConfig.tableColor" class="w-8 h-8 p-0 border rounded cursor-pointer">
                        </div>
                        <div><label class="text-[10px] text-gray-500 uppercase font-bold">圆桌数量</label><input type="number" v-model.number="layout.roundtableConfig.tables" class="w-full border rounded p-1.5 text-sm" @change="initSeats()" min="1" max="8"></div>
                        <label class="flex items-start space-x-2 border border-amber-200 bg-amber-50 rounded-lg p-3 cursor-pointer">
                            <input type="checkbox" v-model="layout.roundtableConfig.dualPrincipal" class="w-4 h-4 mt-0.5">
                            <span>
                                <span class="block text-xs font-bold text-amber-800">主宾双主位</span>
                                <span class="block text-[9px] text-amber-700 mt-0.5">主人1号与宾客1号在12点方向左右并列，其余人员仍按各自名单次序展开。</span>
                            </span>
                        </label>
                        <div v-for="(ring, ri) in layout.roundtableConfig.rings" :key="ring.id" class="border p-3 rounded-lg bg-gray-50">
                            <div class="flex justify-between items-center mb-2">
                                <span class="text-xs font-bold text-gray-700">第{{ ri + 1 }}圈</span>
                                <button @click="removeRoundtableRing(ri)" v-if="layout.roundtableConfig.rings.length > 1" class="text-red-400 hover:text-red-600 text-[10px]"><i class="fa-solid fa-trash-can"></i></button>
                            </div>
                            <input type="text" v-model="ring.label" class="w-full border rounded p-1.5 text-xs mb-2" placeholder="圈标签">
                            <div class="grid grid-cols-2 gap-2">
                                <div><label class="text-[9px] text-gray-500">座位数</label><input type="number" v-model.number="ring.seatCount" class="w-full border rounded p-1.5 text-sm" @change="initSeats()" min="4" max="30"></div>
                                <div><label class="text-[9px] text-gray-500">半径 (px)</label><input type="number" v-model.number="ring.radius" class="w-full border rounded p-1.5 text-sm" @change="initSeats()" min="100" max="360" step="10"></div>
                            </div>
                        </div>
                        <button @click="addRoundtableRing" v-if="layout.roundtableConfig.rings.length < 3" class="w-full text-[10px] bg-gray-100 border border-dashed border-gray-400 py-1.5 rounded hover:bg-gray-200 text-gray-600"><i class="fa-solid fa-plus mr-1"></i> 增加一圈</button>
                    </div>

                    <!-- U 型桌式配置 -->
                    <div v-if="layout.layoutType === 'utable'" class="space-y-3">
                        <div class="border border-slate-200 rounded-lg p-3 bg-slate-50">
                            <label class="text-[10px] text-gray-500 uppercase font-bold block mb-1">领导席位置</label>
                            <select v-model="layout.utableConfig.leaderSide" class="w-full border rounded p-2 text-sm bg-white">
                                <option value="top">领导坐上边（下边为 PPT／演讲台）</option>
                                <option value="bottom">领导坐下边（上边为 PPT／演讲台）</option>
                            </select>
                        </div>
                        <div><label class="text-[9px] text-gray-500">领导席数</label><input type="number" v-model.number="layout.utableConfig.leaderSeatCount" @input="scheduleUtableSeatSync" @change="initSeats()" min="1" max="20" class="w-full border rounded p-1.5 text-sm"></div>
                        <div class="grid grid-cols-2 gap-2 border rounded-lg p-2 bg-orange-50">
                            <div><label class="text-[9px] text-orange-700 font-bold">左侧列数</label><input type="number" v-model.number="layout.utableConfig.leftColumnCount" @input="scheduleUtableSeatSync" @change="initSeats()" min="0" max="10" class="w-full border rounded p-1.5 text-sm"></div>
                            <div><label class="text-[9px] text-orange-700 font-bold">左侧每列席数</label><input type="number" v-model.number="layout.utableConfig.leftSeatCount" @input="scheduleUtableSeatSync" @change="initSeats()" min="0" max="30" class="w-full border rounded p-1.5 text-sm"></div>
                        </div>
                        <div class="grid grid-cols-2 gap-2 border rounded-lg p-2 bg-purple-50">
                            <div><label class="text-[9px] text-purple-700 font-bold">右侧列数</label><input type="number" v-model.number="layout.utableConfig.rightColumnCount" @input="scheduleUtableSeatSync" @change="initSeats()" min="0" max="10" class="w-full border rounded p-1.5 text-sm"></div>
                            <div><label class="text-[9px] text-purple-700 font-bold">右侧每列席数</label><input type="number" v-model.number="layout.utableConfig.rightSeatCount" @input="scheduleUtableSeatSync" @change="initSeats()" min="0" max="30" class="w-full border rounded p-1.5 text-sm"></div>
                        </div>
                        <div class="grid grid-cols-2 gap-2">
                            <div><label class="text-[9px] text-gray-500">领导席间距(px)</label><input type="number" v-model.number="layout.utableConfig.leaderGap" min="0" max="100" step="2" class="w-full border rounded p-1.5 text-sm"></div>
                            <div><label class="text-[9px] text-gray-500">两侧座间距(px)</label><input type="number" v-model.number="layout.utableConfig.sideGap" min="0" max="100" step="2" class="w-full border rounded p-1.5 text-sm"></div>
                        </div>
                        <div><label class="text-[9px] text-gray-500">对侧区域名称</label><input type="text" v-model="layout.utableConfig.stageLabel" class="w-full border rounded p-1.5 text-sm" placeholder="PPT／演讲台"></div>
                        <div class="flex items-center space-x-2"><label class="text-[9px] text-gray-500">U 型边线颜色</label><input type="color" v-model="layout.utableConfig.accentColor" class="w-8 h-8 p-0 border rounded cursor-pointer"></div>
                        <p class="text-[9px] text-slate-500 bg-slate-50 border rounded p-2">领导在上方时按主位居左、先右再左展开；领导在下方时左右镜像。两侧第1列为靠近中间开放区的内列，座位从靠近领导的一端开始，再按距离和内外列依次延伸。</p>
                    </div>

                    <h2 class="text-sm font-bold border-l-4 border-indigo-500 pl-2 mb-2 mt-4 bg-indigo-50 py-1">🎨 画布交互模式</h2>
                    <select v-model="clickMode" class="w-full border rounded p-2 text-sm bg-gray-50 mb-1 font-bold text-indigo-700 focus:ring-2 focus:ring-indigo-500 outline-none">
                        <option value="none">鼠标仅查看信息 / 右键座位管理</option>
                        <option value="toggle_aisle">🟩 点击挖空设为走廊</option>
                        <option value="toggle_disabled">🚫 点击设为禁用废座</option>
                        <option v-if="layout.layoutType === 'classroom'" value="draw_region">🖌️ 框选色彩区域 (用于分组排座)</option>
                        <option v-if="layout.layoutType === 'classroom'" value="draw_region_batch">🖌️🖌️ 批量框选色彩区域 (点击首尾座位批量连线)</option>
                        <option v-if="layout.layoutType === 'classroom'" value="define_sequence">🔢 自定义区域排座次序</option>
                    </select>
                    <p v-if="layout.layoutType === 'classroom'" class="text-[9px] text-red-500 mb-3">🔥 技巧：鼠标<b>右键点击</b>外圈行/列号，可插入、删除、设为走廊。</p>
                    <p v-else class="text-[9px] text-indigo-500 mb-3">提示：可直接点击座位设为走廊或禁用，右键可强制安排、清空或设置显示字段。</p>

                    <div v-if="clickMode === 'draw_region' || clickMode === 'draw_region_batch'" class="mb-4 border p-3 rounded-lg bg-gray-50 shadow-inner">
                        <div class="flex justify-between items-center mb-2">
                            <label class="text-xs font-bold">可用色彩区块：</label>
                            <button @click="addRegionType" class="text-[10px] bg-white border px-2 py-0.5 rounded hover:bg-gray-100"><i class="fa-solid fa-plus"></i> 新增</button>
                        </div>
                        <div class="space-y-2 max-h-32 overflow-y-auto pr-1">
                            <div v-for="reg in layout.regionTypes" :key="reg.id"
                                 @click="activeRegionId = reg.id"
                                 :class="['flex items-center p-1.5 rounded cursor-pointer border', activeRegionId === reg.id ? 'border-indigo-500 bg-white shadow-sm' : 'border-transparent hover:bg-gray-100']">
                                <input type="color" v-model="reg.color" class="w-5 h-5 p-0 border-0 mr-2 rounded cursor-pointer shadow-sm" @click.stop>
                                <input type="text" v-model="reg.name" class="flex-1 text-xs bg-transparent border-b border-dashed focus:outline-none focus:border-indigo-500">
                                <button @click.stop="removeRegionType(reg.id)" class="text-red-400 hover:text-red-600 ml-2"><i class="fa-solid fa-trash-can"></i></button>
                            </div>
                        </div>
                    </div>
                    <div v-if="clickMode === 'draw_region_batch'" class="mb-3 p-2 bg-indigo-50 border border-indigo-200 rounded text-xs">
                        <p v-if="!batchSelectState.hasStart" class="text-indigo-700 font-bold"><i class="fa-solid fa-circle-info mr-1"></i> 批量模式：点击第一个座位作为起点</p>
                        <p v-else class="text-indigo-700 font-bold"><i class="fa-solid fa-hand-pointer mr-1"></i> 已选中起点，请点击同行的终点或同列的终点</p>
                    </div>

                    <div class="mb-4 border p-3 rounded-lg bg-gray-50 shadow-inner">
                        <label class="text-xs font-bold text-gray-700 block mb-2">📋 座席显示字段配置</label>
                        <div class="grid grid-cols-2 gap-2">
                            <div>
                                <label class="text-[9px] text-gray-500">主字段 (大字显示)</label>
                                <select v-model="layout.seatDisplayConfig.mainField" @change="autoSave" class="w-full border rounded p-1.5 text-xs bg-white">
                                    <option value="姓名">姓名</option>
                                    <option value="单位">单位</option>
                                    <option value="类别">类别</option>
                                    <option value="电话">电话</option>
                                    <option v-for="key in extraKeys" :key="'mf_'+key" :value="key">{{key}}</option>
                                </select>
                            </div>
                            <div>
                                <label class="text-[9px] text-gray-500">副字段 (小字显示)</label>
                                <select v-model="layout.seatDisplayConfig.subField" @change="autoSave" class="w-full border rounded p-1.5 text-xs bg-white">
                                    <option value="">无</option>
                                    <option value="姓名">姓名</option>
                                    <option value="单位">单位</option>
                                    <option value="类别">类别</option>
                                    <option value="电话">电话</option>
                                    <option v-for="key in extraKeys" :key="'sf_'+key" :value="key">{{key}}</option>
                                </select>
                            </div>
                        </div>
                    </div>

                    <div v-if="clickMode === 'define_sequence'" class="mb-4 border p-3 rounded-lg bg-indigo-50 shadow-inner border-indigo-200">
                        <div class="text-xs font-bold text-indigo-800 mb-1">自定义排序连线：</div>
                        <p class="text-[9px] text-indigo-600 mb-2">直接在右侧点击带有颜色的座位。系统会记录您的点击顺序（角标序号）。排座时可选择按此顺序依次入座。</p>
                        <div class="flex items-center space-x-2">
                            <span class="text-[10px] font-bold">当前区：</span>
                            <span class="text-[10px] px-2 py-0.5 rounded text-white shadow-sm" :style="{backgroundColor: layout.regionTypes.find(r=>r.id===activeRegionId)?.color || '#9ca3af'}">
                                {{ layout.regionTypes.find(r=>r.id===activeRegionId)?.name || '请先点击右侧区域' }}
                            </span>
                            <button @click="clearCustomSequence(activeRegionId)" class="text-[10px] bg-white border border-indigo-300 text-indigo-600 px-2 py-0.5 rounded hover:bg-indigo-100 shadow-sm ml-auto">清空此区次序</button>
                        </div>
                    </div>

                    <div class="mt-auto pt-4 border-t">
                        <label class="text-xs text-gray-500 font-bold block mb-1">模板库 (复用给其他会议)</label>
                        <div class="flex space-x-1 mb-2">
                            <input type="text" v-model="layoutName" placeholder="存为新模板名..." class="flex-1 border rounded px-2 py-1.5 text-sm">
                            <button @click="saveTemplate" class="bg-gray-800 text-white px-3 rounded text-sm hover:bg-gray-900 shadow">保存</button>
                        </div>
                        <div class="flex space-x-1 mt-1">
                            <select v-model="selectedTemplateId" class="flex-1 border rounded p-1.5 text-sm bg-gray-50">
                                <option value="">-- 选择历史模板 --</option>
                                <option v-for="l in templates" :key="l.id" :value="l.id">{{l.name}}</option>
                            </select>
                            <button @click="loadTemplate(selectedTemplateId)" class="bg-indigo-100 text-indigo-700 px-2 rounded text-xs font-bold hover:bg-indigo-200 shadow-sm">加载</button>
                            <button @click="deleteTemplate(selectedTemplateId)" class="bg-red-50 text-red-600 px-2.5 rounded hover:bg-red-100 border border-red-200 shadow-sm" title="删除选中模板"><i class="fa-solid fa-trash-can"></i></button>
                        </div>
                    </div>
                </div>

                <!-- Tab 2: 人员管理 -->
                <div v-show="activeTab === 'attendees'" class="p-4 flex flex-col h-full bg-gray-50">
                    <div class="mb-3 bg-white p-3 rounded-lg border shadow-sm">
                        <label class="block text-xs font-bold text-gray-700 mb-2"><i class="fa-solid fa-file-csv text-green-600 mr-1"></i> 名单导入</label>
                        <div class="flex space-x-2 items-center">
                            <input type="file" @change="uploadFile" accept=".xlsx,.xls,.csv" class="flex-1 text-xs border rounded file:bg-indigo-50 file:border-0 file:text-indigo-700 file:px-2 file:py-1 file:rounded cursor-pointer"/>
                        </div>
                        <p class="text-[9px] text-gray-400 mt-2">包含标题行: <span class="font-bold">序号, 姓名, 单位, 主宾, 类别, 电话, 备注等</span>；“主宾”可填写主方/上级或宾方/下级。</p>
                    </div>

                    <div v-if="isPartyLayout" class="mb-3 border border-amber-200 bg-amber-50 p-3 rounded-lg shadow-sm">
                        <div class="flex items-center justify-between mb-2">
                            <span class="text-xs font-bold text-amber-800">🤝 先分主宾，再分别排序</span>
                            <select v-model="attendeeRoleFilter" class="border rounded bg-white px-2 py-1 text-[10px]">
                                <option value="all">全部人员</option>
                                <option value="host">只看主方</option>
                                <option value="guest">只看宾方</option>
                                <option value="unassigned">只看未分类</option>
                            </select>
                        </div>
                        <div class="flex space-x-1 text-[10px] mb-2">
                            <span class="bg-amber-100 text-amber-800 border border-amber-200 px-2 py-1 rounded">主方 {{ partyCounts.host }} 人</span>
                            <span class="bg-blue-100 text-blue-800 border border-blue-200 px-2 py-1 rounded">宾方 {{ partyCounts.guest }} 人</span>
                            <span class="bg-gray-100 text-gray-600 border px-2 py-1 rounded">未分类 {{ partyCounts.unassigned }} 人</span>
                        </div>
                        <div v-if="selectedAttIds.length" class="flex space-x-1">
                            <button @click="setSelectedPartyRole('host')" class="flex-1 bg-amber-600 text-white rounded py-1 text-[10px] font-bold">设为主方</button>
                            <button @click="setSelectedPartyRole('guest')" class="flex-1 bg-blue-600 text-white rounded py-1 text-[10px] font-bold">设为宾方</button>
                            <button @click="setSelectedPartyRole('unassigned')" class="px-2 bg-gray-200 text-gray-600 rounded py-1 text-[10px]">取消分类</button>
                        </div>
                    </div>

                    <div class="flex justify-between items-center mb-2">
                        <span class="text-sm font-bold text-gray-700"><span class="bg-indigo-100 text-indigo-800 px-1.5 rounded">{{ attendees.length }}</span> 人</span>
                        <div class="space-x-1">
                            <button @click="bulkDeleteAttendees" class="text-[10px] bg-red-100 text-red-600 px-2 py-1 rounded hover:bg-red-200 border border-red-200" v-if="selectedAttIds.length>0">删选中</button>
                            <button @click="openAddModal" class="text-[10px] bg-indigo-500 text-white px-2 py-1 rounded hover:bg-indigo-600 shadow"><i class="fa-solid fa-plus"></i> 单个加</button>
                        </div>
                    </div>

                    <input type="text" v-model="attendeeSearch" placeholder="搜索姓名/单位..." class="w-full border rounded p-2 text-sm mb-2 focus:ring-1 focus:ring-indigo-500 outline-none shadow-sm">

                    <div class="flex items-center mb-1 pl-1">
                        <input type="checkbox" :checked="isAllSelected" @change="toggleSelectAllFiltered" class="w-3.5 h-3.5 mr-2 cursor-pointer text-indigo-600">
                        <span class="text-[10px] text-gray-500 font-bold flex-1" v-if="!attendeeSearch">{{ isPartyLayout ? '全选 / 主方与宾方分别拖拽排序' : '全选 / 拖拽项可改变排序同步至该排座批次' }}</span>
                        <span class="text-[10px] text-blue-600 font-bold flex-1" v-else>全选当前搜索出的结果</span>
                    </div>

                    <div class="flex-1 overflow-y-auto border rounded bg-white shadow-inner">
                        <div v-for="(att, index) in filteredAttendees" :key="att.id"
                             class="p-2 border-b text-sm flex items-center hover:bg-indigo-50 group transition"
                             :draggable="!attendeeSearch"
                             @dragstart="onDragStart($event, att.id)"
                             @dragover.prevent="onDragOver($event, att.id)"
                             @dragleave="onDragLeave($event)"
                             @drop="onDrop($event, att.id)"
                             :class="{'drag-over': dragTargetIndex === att.id}">

                            <input type="checkbox" v-model="selectedAttIds" :value="att.id" class="mr-2 w-3.5 h-3.5 text-indigo-600 cursor-pointer">
                            <div class="flex flex-col items-center mr-1">
                                <button @click="moveAttendee(att.id, -1)" class="text-[8px] text-gray-400 hover:text-indigo-600" :disabled="isFirstInAttendeeScope(att)" title="上移"><i class="fa-solid fa-chevron-up"></i></button>
                                <input type="number" :value="getAttendeeOrder(att)" @change="setAttendeePosition(att.id, $event.target.value)" class="w-8 text-[10px] text-gray-400 font-bold text-center border-0 focus:outline-none focus:bg-indigo-50 focus:text-indigo-600" title="输入本方序号调整位置">
                                <button @click="moveAttendee(att.id, 1)" class="text-[8px] text-gray-400 hover:text-indigo-600" :disabled="isLastInAttendeeScope(att)" title="下移"><i class="fa-solid fa-chevron-down"></i></button>
                            </div>

                            <div class="flex-1 truncate">
                                <div class="font-bold text-gray-800">{{att.name}} <span class="text-[9px] bg-gray-100 border px-1 rounded text-gray-500 ml-1 font-normal">{{att.category}}</span></div>
                                <div class="text-[10px] text-gray-500 mt-0.5 truncate">{{att.unit}}</div>
                            </div>
                            <select v-if="isPartyLayout" :value="normalizePartyRole(att.party_role)" @change="setAttendeePartyRole(att.id, $event.target.value)" class="w-[74px] border rounded px-1 py-1 text-[9px] font-bold mr-1" :class="partyRoleClass(att.party_role)">
                                <option value="unassigned">未分类</option>
                                <option value="host">主方</option>
                                <option value="guest">宾方</option>
                            </select>
                            <div class="opacity-0 group-hover:opacity-100 flex flex-col space-y-1 ml-1">
                                <button @click="editAttendee(att)" class="text-blue-500 hover:bg-blue-100 rounded px-1"><i class="fa-solid fa-pen"></i></button>
                            </div>
                        </div>
                        <div v-if="filteredAttendees.length===0" class="p-4 text-center text-xs text-gray-400">暂无人员</div>
                    </div>
                </div>

                <!-- Tab 3: 排座管理 -->
                <div v-show="activeTab === 'assign'" class="p-5 flex flex-col h-full overflow-y-auto">

                    <div v-if="layout.layoutType === 'classroom'" class="bg-blue-50 border border-blue-200 p-3 rounded-lg mb-4 shadow-sm">
                        <p class="text-[10px] text-blue-800 font-bold mb-1"><i class="fa-solid fa-lightbulb"></i> 标准排座四步曲：</p>
                        <ol class="text-[9px] text-blue-700 list-decimal pl-4 space-y-0.5">
                            <li>先选<b>主席台</b>区块填入核心领导；</li>
                            <li>鼠标<b>右键图上座位</b>强制绑定零星大咖；</li>
                            <li>选定<b>色彩区域</b>，填入所属分类人员；</li>
                            <li>选定<b>主会场</b>，一键填补其余所有人。</li>
                        </ol>
                    </div>

                    <div v-else-if="isPartyLayout" class="bg-amber-50 border border-amber-200 p-3 rounded-lg mb-4 shadow-sm">
                        <div class="flex justify-between items-center">
                            <p class="text-[10px] text-amber-800 font-bold">🤝 主宾分序：主方 {{ partyCounts.host }} 人，宾方 {{ partyCounts.guest }} 人，未分类 {{ partyCounts.unassigned }} 人</p>
                            <button @click="activeTab='attendees'" class="text-[9px] bg-white border border-amber-300 text-amber-700 px-2 py-1 rounded">前往分类排序</button>
                        </div>
                        <p v-if="partyCounts.unassigned > 0" class="text-[9px] text-red-600 mt-1">执行“主宾自动排座”前，所选人员必须全部完成主宾分类。</p>
                    </div>

                    <div class="mb-4">
                        <label class="text-xs font-bold text-gray-700 uppercase block mb-1">选目标区块</label>
                        <select v-model="assignConfig.targetZone" class="w-full border rounded p-2 bg-white text-sm font-bold text-indigo-700 shadow-sm focus:ring-1">
                            <template v-if="layout.layoutType === 'classroom'">
                                <option value="podium">⭐ 主席台 (严格按图上排/列号顺排)</option>
                                <optgroup label="主会场特定区:">
                                    <option v-for="reg in layout.regionTypes" :key="reg.id" :value="'reg_' + reg.id">区域: {{reg.name}}</option>
                                </optgroup>
                                <option value="all_m">主会场全局 (跳过已被排好的功能区)</option>
                            </template>
                            <template v-else-if="layout.layoutType === 'longtable'">
                                <option value="party_auto">🤝 主宾自动对坐（推荐）</option>
                                <option value="all">全部座位（不分主宾）</option>
                                <option v-if="layout.longtableConfig.sides.top.rows > 0" value="lt-top">上侧 ({{ layout.longtableConfig.sides.top.label }})</option>
                                <option v-if="layout.longtableConfig.sides.bottom.rows > 0" value="lt-bottom">下侧 ({{ layout.longtableConfig.sides.bottom.label }})</option>
                                <option v-if="longtableHasLeft" value="lt-left">左侧 ({{ layout.longtableConfig.sides.left.label }})</option>
                                <option v-if="longtableHasRight" value="lt-right">右侧 ({{ layout.longtableConfig.sides.right.label }})</option>
                                <option v-if="layout.longtableConfig.headSeats.enabled" value="lt-head-left">桌外主位-左</option>
                                <option v-if="layout.longtableConfig.headSeats.enabled" value="lt-head-right">桌外主位-右</option>
                            </template>
                            <template v-else-if="layout.layoutType === 'roundtable'">
                                <option value="party_auto">🤝 主宾分序自动排座（推荐）</option>
                                <option value="all">全部座位（不分主宾）</option>
                                <optgroup v-for="t in layout.roundtableConfig.tables" :key="'rt-target-'+t" :label="'第 '+t+' 桌'">
                                    <option :value="'rt-'+(t-1)">第 {{ t }} 桌全部座位</option>
                                    <option v-for="ring in layout.roundtableConfig.rings" :key="t+'-'+ring.id" :value="'rt-'+(t-1)+'-'+ring.id">{{ ring.label }}</option>
                                </optgroup>
                            </template>
                            <template v-else-if="layout.layoutType === 'utable'">
                                <option value="u_auto">∪ U型桌自动排座（推荐）</option>
                                <option value="all">全部座位</option>
                                <option value="u-leader">领导席</option>
                                <option value="u-left">左侧参会席</option>
                                <option value="u-right">右侧参会席</option>
                            </template>
                        </select>
                    </div>

                    <div class="mb-4" v-if="layout.layoutType === 'classroom' && assignConfig.targetZone !== 'podium'">
                        <label class="text-xs font-bold text-gray-700 uppercase block mb-1">选填充路线</label>
                        <select v-model="assignConfig.rule" class="w-full border rounded p-2 bg-white text-sm shadow-sm">
                            <option value="center_lr">主位居左，先右再左依次排序</option>
                            <option value="center_rl">主位居右，先左再右依次排序</option>
                            <option value="ltr">从左往右逐行填</option>
                            <option value="rtl">从右往左逐行填</option>
                            <option value="snake">蛇形折返填</option>
                            <option value="custom_seq" class="font-bold text-indigo-600">🔢 自定义连线顺序 (仅限色彩功能区)</option>
                        </select>
                    </div>
                    <div class="mb-4" v-if="layout.layoutType === 'longtable'">
                        <label class="text-xs font-bold text-gray-700 uppercase block mb-1">选填充路线</label>
                        <select v-model="assignConfig.rule" class="w-full border rounded p-2 bg-white text-sm shadow-sm">
                            <option value="center_out">🎯 居中向外扩展 (领导居中，先左后右)</option>
                            <option value="ltr">从左往右依次</option>
                            <option value="rtl">从右往左依次</option>
                            <option value="snake">蛇形折返填 (多边时)</option>
                        </select>
                        <p v-if="assignConfig.targetZone === 'party_auto'" class="text-[9px] text-amber-700 mt-2 p-2 bg-amber-50 border border-amber-200 rounded">主方先排 {{ longtableHostSideLabel }}、宾方先排 {{ longtableGuestSideLabel }}；任一方本侧不足时，都从靠近自己的一端同时向左、右两侧延伸，并保持各自名单次序。</p>
                    </div>
                    <div class="mb-4" v-if="layout.layoutType === 'roundtable'">
                        <label class="text-xs font-bold text-gray-700 uppercase block mb-1">选填充路线</label>
                        <select v-model="assignConfig.rule" class="w-full border rounded p-2 bg-white text-sm shadow-sm">
                            <option value="banquet">🍽️ 主宾分序 (主方逆时针、宾方顺时针)</option>
                            <option v-if="assignConfig.targetZone !== 'party_auto'" value="clockwise">🕐 顺时针 (从12点方向起，不分主宾)</option>
                            <option v-if="assignConfig.targetZone !== 'party_auto'" value="counterclockwise">🕐 逆时针 (从12点方向起，不分主宾)</option>
                        </select>
                        <p class="text-[9px] text-gray-400 mt-1" v-if="!layout.roundtableConfig.dualPrincipal">12点钟方向为主方第1位；主方按逆时针方向、宾方按顺时针方向，分别依照各自名单次序落座。</p>
                        <p class="text-[9px] text-amber-700 mt-1" v-else>主人1号在12点方向左侧、宾客1号在右侧并列双主位；两方再分别向各自一侧展开。</p>
                        <p v-if="assignConfig.targetZone === 'party_auto' && assignConfig.rule === 'banquet'" class="text-[9px] text-amber-700 mt-2 p-2 bg-amber-50 border border-amber-200 rounded">系统会先拆分主方和宾方两份名单，再分别占用圆桌两侧，不会混用次序。</p>
                    </div>
                    <div class="mb-4" v-if="layout.layoutType === 'utable'">
                        <label class="text-xs font-bold text-gray-700 uppercase block mb-1">填充路线</label>
                        <div class="text-[10px] text-slate-700 p-2 bg-slate-50 border border-slate-200 rounded">自动排座时，名单前 {{ layout.utableConfig.leaderSeatCount }} 位优先进入领导席并居中展开；其余人员从靠近领导的一端开始，同一距离按“左内列、右内列、左外列、右外列”依次排，再向 PPT／演讲台方向延伸。</div>
                    </div>
                    <div class="mb-3 flex-1 flex flex-col min-h-[220px]">
                        <div class="flex justify-between items-center mb-1">
                            <label class="text-xs font-bold text-gray-700 uppercase">选人入座 (选了 {{assignSelectedAtts.length}} 人)</label>
                            <button @click="selectAllUnassigned" class="text-[10px] bg-gray-200 px-2 rounded hover:bg-gray-300">全选未排</button>
                        </div>
                        <div class="flex flex-col space-y-1 mb-1">
                            <input type="text" v-model="assignSearch" placeholder="搜索姓名/单位..." class="w-full border rounded px-2 py-1 text-sm outline-none focus:ring-1 focus:ring-indigo-500">
                            <select v-if="isPartyLayout" v-model="assignRoleFilter" class="w-full border rounded px-2 py-1 text-[11px] bg-white">
                                <option value="all">显示全部主宾</option>
                                <option value="host">只选主方（上级）</option>
                                <option value="guest">只选宾方（下级）</option>
                                <option value="unassigned">只看未分类</option>
                            </select>
                            <div class="flex space-x-1">
                                <input type="text" v-model="assignRangeInput" @keyup.enter="applyRangeSelection" placeholder="序号范围(如1-3,5,9-12)" class="flex-1 border rounded px-2 py-1 text-[11px] outline-none focus:ring-1 focus:ring-indigo-500" title="按回车或点击按钮勾选">
                                <button @click="applyRangeSelection" class="bg-indigo-100 text-indigo-700 px-2 rounded text-xs font-bold hover:bg-indigo-200">选取</button>
                            </div>
                        </div>
                        <div class="flex-1 overflow-y-auto border rounded bg-white shadow-inner p-1">
                            <label v-for="att in unassignedFilteredAttendees" :key="att.id" class="flex items-center p-1 hover:bg-indigo-50 rounded cursor-pointer border-b border-gray-50 last:border-0 transition">
                                <input type="checkbox" :value="att" v-model="assignSelectedAtts" class="mr-2 w-3.5 h-3.5 text-indigo-600">
                                <span class="text-[9px] text-gray-400 w-6 text-right mr-1 font-bold">{{ getAttendeeOrder(att) }}</span>
                                <div class="text-[11px] truncate flex-1">
                                    <span class="font-bold text-gray-800">{{att.name}}</span>
                                    <span class="text-[9px] text-gray-500 ml-1">{{att.unit}}</span>
                                </div>
                                <span v-if="isPartyLayout" class="text-[8px] border rounded px-1 py-0.5 ml-1" :class="partyRoleClass(att.party_role)">{{ partyRoleLabel(att.party_role) }}</span>
                            </label>
                        </div>
                    </div>

                    <div class="flex space-x-2 mt-auto">
                        <button @click="clearTargetZone" class="w-1/3 bg-white border border-red-200 text-red-500 font-bold py-2.5 rounded shadow-sm hover:bg-red-50 text-xs">清除该区</button>
                        <button @click="executeAssignment" class="flex-1 bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-2.5 rounded shadow-md transition transform hover:scale-[1.02] text-sm">
                            <i class="fa-solid fa-bolt mr-1"></i> 执行排座
                        </button>
                    </div>
                </div>

                <!-- Tab 4: 导出发布 -->
                <div v-show="activeTab === 'export'" class="p-5 flex flex-col h-full overflow-y-auto bg-gray-50">
                    <div class="mb-4 bg-white border p-4 rounded-lg shadow-sm">
                        <h3 class="font-bold text-gray-800 mb-2"><i class="fa-solid fa-file-excel text-green-600 mr-1"></i> 导出综合表格 (Excel)</h3>
                        <div class="mb-3">
                            <label class="text-[10px] text-gray-500 font-bold block mb-1">座席方块显示格式 (点击组合)：</label>
                            <div class="flex flex-wrap gap-1 mb-2">
                                <span class="text-[9px] bg-green-50 border border-green-200 text-green-700 px-1 py-0.5 rounded cursor-pointer hover:bg-green-100 font-bold shadow-sm" @click="insertExportVar('{姓名}')">{姓名}</span>
                                <span class="text-[9px] bg-green-50 border border-green-200 text-green-700 px-1 py-0.5 rounded cursor-pointer hover:bg-green-100 font-bold shadow-sm" @click="insertExportVar('{单位}')">{单位}</span>
                                <span class="text-[9px] bg-green-50 border border-green-200 text-green-700 px-1 py-0.5 rounded cursor-pointer hover:bg-green-100 font-bold shadow-sm" @click="insertExportVar('{类别}')">{类别}</span>
                                <span class="text-[9px] bg-green-50 border border-green-200 text-green-700 px-1 py-0.5 rounded cursor-pointer hover:bg-green-100 font-bold shadow-sm" @click="insertExportVar('{电话}')">{电话}</span>
                                <span v-for="key in extraKeys" :key="'exp_'+key" class="text-[9px] bg-gray-50 border border-gray-200 text-gray-600 px-1 py-0.5 rounded cursor-pointer hover:bg-gray-100 shadow-sm" @click="insertExportVar(`{${key}}`)">{{"{"+key+"}"}}</span>
                                <span class="text-[9px] bg-red-50 border border-red-200 text-red-700 px-1 py-0.5 rounded cursor-pointer hover:bg-red-100 font-bold shadow-sm" @click="insertExportVar('\n')">↵ 换行</span>
                            </div>
                            <textarea id="exportTemplate" v-model="layout.export_template" @change="autoSave" class="w-full border rounded p-1.5 text-xs bg-gray-50" rows="3" placeholder="如：{姓名}\n({单位})"></textarea>
                        </div>
                        <button @click="exportExcel" class="w-full bg-green-600 hover:bg-green-700 text-white py-2 rounded shadow text-sm font-bold transition">下载本场排座表</button>
                    </div>

                    <div class="bg-white border p-4 rounded-lg shadow-sm flex-1 flex flex-col">
                        <h3 class="font-bold text-gray-800 mb-2"><i class="fa-solid fa-comment-sms text-blue-500 mr-1"></i> 短信模板配置</h3>
                        <p class="text-[9px] text-gray-400 mb-2">点击下方变量插入模板，已全量扫描所有列字段。</p>
                        <div class="flex flex-wrap gap-1 mb-2">
                            <span class="text-[9px] bg-blue-50 border border-blue-200 text-blue-700 px-1 py-0.5 rounded cursor-pointer hover:bg-blue-100 font-bold shadow-sm" @click="insertVar('{姓名}')">{姓名}</span>
                            <span class="text-[9px] bg-blue-50 border border-blue-200 text-blue-700 px-1 py-0.5 rounded cursor-pointer hover:bg-blue-100 font-bold shadow-sm" @click="insertVar('{单位}')">{单位}</span>
                            <span class="text-[9px] bg-blue-50 border border-blue-200 text-blue-700 px-1 py-0.5 rounded cursor-pointer hover:bg-blue-100 font-bold shadow-sm" @click="insertVar('{类别}')">{类别}</span>
                            <span class="text-[9px] bg-blue-50 border border-blue-200 text-blue-700 px-1 py-0.5 rounded cursor-pointer hover:bg-blue-100 font-bold shadow-sm" @click="insertVar('{电话}')">{电话}</span>
                            <span class="text-[9px] bg-blue-50 border border-blue-200 text-blue-700 px-1 py-0.5 rounded cursor-pointer hover:bg-blue-100 font-bold shadow-sm" @click="insertVar('{排号}')">{排号}</span>
                            <span class="text-[9px] bg-blue-50 border border-blue-200 text-blue-700 px-1 py-0.5 rounded cursor-pointer hover:bg-blue-100 font-bold shadow-sm" @click="insertVar('{座号}')">{座号}</span>
                            <span v-for="key in extraKeys" :key="'sms_'+key" class="text-[9px] bg-gray-50 border border-gray-200 text-gray-600 px-1 py-0.5 rounded cursor-pointer hover:bg-gray-100 shadow-sm" @click="insertVar(`{${key}}`)">{{"{"+key+"}"}}</span>
                        </div>
                        <textarea id="smsTemplate" v-model="layout.sms_template" @change="autoSave" class="w-full flex-1 border rounded p-2 text-xs focus:outline-none focus:ring-1 focus:ring-blue-500 resize-none mb-2 bg-gray-50" placeholder="请输入短信模板..."></textarea>
                        <button @click="autoSave" class="w-full bg-blue-50 border border-blue-200 text-blue-600 hover:bg-blue-100 py-1.5 rounded text-xs font-bold">保存模板</button>
                    </div>
                </div>
            </div>

            <!-- 右侧主画布区 -->
            <div class="flex-1 overflow-auto bg-[#e2e8f0] relative" id="canvas-container">
                <div class="venue-wrapper">

                    <input type="text" v-model="meetingNameInput" @change="renameMeeting" class="text-3xl font-extrabold text-gray-800 mb-8 tracking-widest drop-shadow-sm meeting-title-input" placeholder="请输入会议名称">

                    <!-- ========== 课桌式布局 ========== -->
                    <div v-if="layout.layoutType === 'classroom'" class="classroom-layout">

                    <!-- 主席台 Grid -->
                    <div v-if="layout.podium.rows > 0 && layout.podium.cols > 0" class="grid-container relative border-t-8" :style="{ gridTemplateColumns: `45px repeat(${layout.podium.cols}, 85px) 45px`, borderColor: layout.podium.color }">
                        <div class="absolute -top-3 left-1/2 transform -translate-x-1/2 px-4 py-0.5 rounded-full text-xs font-bold shadow" :style="{backgroundColor: layout.podium.color, color: layout.podium.textColor}">主席台</div>

                        <div></div>
                        <div v-for="c in layout.podium.cols" :key="'ptcol'+c" class="header-cell h-7" @contextmenu.prevent="handleAxisRightClick($event, 'col', c-1, 'p')">
                            <input type="text" v-model="layout.podium.colLabels[c-1]" class="axis-input" @change="handleLabelChange('col', c-1, 'p'); syncPodiumLabels()" placeholder="号">
                        </div>
                        <div></div>

                        <template v-for="r in layout.podium.rows" :key="'prow'+r">
                            <div class="header-cell" @contextmenu.prevent="handleAxisRightClick($event, 'row', r-1, 'p')">
                                <input type="text" v-model="layout.podium.rowLabels[r-1]" class="axis-input" @change="handleLabelChange('row', r-1, 'p'); syncPodiumLabels()" placeholder="排">
                            </div>

                            <div v-for="c in layout.podium.cols" :key="'pcell'+r+c"
                                 class="seat" :style="seatStyle('p', r-1, c-1)"
                                 :class="seatClass('p', r-1, c-1)"
                                 @click="handleSeatClick('p', r-1, c-1)"
                                 @contextmenu.prevent="handleSeatRightClick($event, 'p', r-1, c-1)">

                                 <div v-if="seats[`p-${r-1}-${c-1}`]?.type === 'aisle'"></div>
                                 <div v-else-if="seats[`p-${r-1}-${c-1}`]?.type === 'disabled'"><i class="fa-solid fa-ban text-gray-300 text-xl"></i></div>

                                 <!-- 空位情况下的排座号 -->
                                 <div v-else-if="!seats[`p-${r-1}-${c-1}`]?.attendee" class="absolute flex flex-col items-center justify-center w-full h-full font-bold pointer-events-none transition-all text-center">
                                     <span class="text-[12px] mb-0.5 whitespace-nowrap flex-shrink-0" :style="{ color: isDuplicateSeat('p', r-1, c-1) ? '#ef4444' : (isSeatBgDark('p', r-1, c-1) ? 'rgba(255,255,255,0.8)' : 'rgba(0,0,0,0.65)') }">{{ layout.podium.rowLabels[r-1] }}排</span>
                                     <input type="text"
                                            :value="getSeatColLabelUI('p', r-1, c-1)"
                                            @change="setCustomLabel('p-'+(r-1)+'-'+(c-1), $event.target.value)"
                                            @click.stop
                                            :class="{'text-red-500 font-extrabold': isDuplicateSeat('p', r-1, c-1)}"
                                            class="w-11/12 bg-transparent border-none text-center font-bold outline-none hover:bg-black/10 focus:bg-white/50 rounded pointer-events-auto min-w-0"
                                            :style="{ color: isDuplicateSeat('p', r-1, c-1) ? '#ef4444' : (isSeatBgDark('p', r-1, c-1) ? 'rgba(255,255,255,0.95)' : 'rgba(0,0,0,0.85)'), fontSize: '13px' }">
                                 </div>

                                 <!-- 有人落座的情况 -->
                                 <div v-else class="w-full h-full relative flex flex-col items-center justify-center p-0.5">
                                     <div class="absolute top-0.5 left-1 flex items-center z-20 w-[90%]">
                                         <span class="text-[8px] font-bold pointer-events-none whitespace-nowrap flex-shrink-0" :style="{ color: isSeatBgDark('p', r-1, c-1) ? 'rgba(255,255,255,0.7)' : 'rgba(0,0,0,0.5)' }">{{ layout.podium.rowLabels[r-1] }}排&nbsp;</span>
                                         <input type="text"
                                                :value="getSeatColLabelUI('p', r-1, c-1)"
                                                @change="setCustomLabel('p-'+(r-1)+'-'+(c-1), $event.target.value)"
                                                @click.stop
                                                class="flex-1 min-w-0 bg-transparent border-none outline-none text-[8px] font-bold p-0 m-0 pointer-events-auto hover:bg-black/10 focus:bg-white/50 rounded"
                                                :style="{ color: isDuplicateSeat('p', r-1, c-1) ? '#ef4444' : (isSeatBgDark('p', r-1, c-1) ? 'rgba(255,255,255,0.9)' : 'rgba(0,0,0,0.7)') }">
                                     </div>
                                     <div class="text-center w-full mt-2 pointer-events-none">
                                        <div class="font-extrabold text-[13px] truncate leading-tight">{{ getAttendeeFieldValue(seats[`p-${r-1}-${c-1}`].attendee, getSeatDisplayConfig('p', r-1, c-1).mainField) }}</div>
                                        <div v-if="getSeatDisplayConfig('p', r-1, c-1).subField" class="text-[8px] opacity-90 truncate leading-tight mt-0.5" :style="{ backgroundColor: isSeatBgDark('p', r-1, c-1) ? 'rgba(0,0,0,0.1)' : 'rgba(255,255,255,0.3)', padding: '0 2px', borderRadius: '2px', display: 'inline-block', maxWidth: '100%' }">{{ getAttendeeFieldValue(seats[`p-${r-1}-${c-1}`].attendee, getSeatDisplayConfig('p', r-1, c-1).subField) }}</div>
                                     </div>
                                 </div>
                            </div>

                            <div class="header-cell" @contextmenu.prevent="handleAxisRightClick($event, 'row', r-1, 'p')">
                                <input type="text" v-model="layout.podium.rowLabels[r-1]" class="axis-input" @change="handleLabelChange('row', r-1, 'p'); syncPodiumLabels()">
                            </div>
                        </template>

                        <div></div>
                        <div v-for="c in layout.podium.cols" :key="'pbcol'+c" class="header-cell h-7" @contextmenu.prevent="handleAxisRightClick($event, 'col', c-1, 'p')">
                            <input type="text" v-model="layout.podium.colLabels[c-1]" class="axis-input" @change="handleLabelChange('col', c-1, 'p'); syncPodiumLabels()">
                        </div>
                        <div></div>
                    </div>

                    <!-- 主会场 Grid -->
                    <div v-if="layout.main.rows > 0 && layout.main.cols > 0" class="grid-container relative border-t-8" :style="{ gridTemplateColumns: `45px repeat(${layout.main.cols}, 85px) 45px`, borderColor: layout.main.color }">
                        <div class="absolute -top-3 left-1/2 transform -translate-x-1/2 px-4 py-0.5 rounded-full text-xs font-bold shadow" :style="{backgroundColor: layout.main.color, color: layout.main.textColor}">主会场</div>

                        <div></div>
                        <div v-for="c in layout.main.cols" :key="'mtcol'+c" class="header-cell h-7" @contextmenu.prevent="handleAxisRightClick($event, 'col', c-1, 'm')">
                            <input type="text" v-model="layout.main.colLabels[c-1]" class="axis-input" @change="handleLabelChange('col', c-1, 'm')" placeholder="列">
                        </div>
                        <div></div>

                        <template v-for="r in layout.main.rows" :key="'mrow'+r">
                            <div class="header-cell" @contextmenu.prevent="handleAxisRightClick($event, 'row', r-1, 'm')">
                                <input type="text" v-model="layout.main.rowLabels[r-1]" class="axis-input" @change="handleLabelChange('row', r-1, 'm')" placeholder="排">
                            </div>

                            <div v-for="c in layout.main.cols" :key="'mcell'+r+c"
                                 class="seat" :style="seatStyle('m', r-1, c-1)"
                                 :class="seatClass('m', r-1, c-1)"
                                 @click="handleSeatClick('m', r-1, c-1)"
                                 @contextmenu.prevent="handleSeatRightClick($event, 'm', r-1, c-1)">

                                 <div v-if="getSeatRegionColor('m', r-1, c-1)" class="region-overlay" :style="{backgroundColor: getSeatRegionColor('m', r-1, c-1), borderColor: getSeatRegionColor('m', r-1, c-1)}"></div>
                                 <div v-if="seats[`m-${r-1}-${c-1}`]?.type === 'aisle'"></div>
                                 <div v-else-if="seats[`m-${r-1}-${c-1}`]?.type === 'disabled'"><i class="fa-solid fa-ban text-gray-300 text-xl"></i></div>

                                 <!-- 空位情况下的排座号 -->
                                 <div v-else-if="!seats[`m-${r-1}-${c-1}`]?.attendee" class="absolute flex flex-col items-center justify-center w-full h-full font-bold pointer-events-none transition-all text-center">
                                     <span class="text-[12px] mb-0.5 whitespace-nowrap flex-shrink-0" :style="{ color: isDuplicateSeat('m', r-1, c-1) ? '#ef4444' : (isSeatBgDark('m', r-1, c-1) ? 'rgba(255,255,255,0.8)' : 'rgba(0,0,0,0.65)') }">{{ layout.main.rowLabels[r-1] }}排</span>
                                     <input type="text"
                                            :value="getSeatColLabelUI('m', r-1, c-1)"
                                            @change="setCustomLabel('m-'+(r-1)+'-'+(c-1), $event.target.value)"
                                            @click.stop
                                            :class="{'text-red-500 font-extrabold': isDuplicateSeat('m', r-1, c-1)}"
                                            class="w-11/12 bg-transparent border-none text-center font-bold outline-none hover:bg-black/10 focus:bg-white/50 rounded pointer-events-auto min-w-0"
                                            :style="{ color: isDuplicateSeat('m', r-1, c-1) ? '#ef4444' : (isSeatBgDark('m', r-1, c-1) ? 'rgba(255,255,255,0.95)' : 'rgba(0,0,0,0.85)'), fontSize: '13px' }">
                                 </div>

                                 <!-- 有人落座的情况 -->
                                 <div v-else class="w-full h-full relative flex flex-col items-center justify-center z-10 p-0.5">
                                     <!-- 连线排号角标 -->
                                     <div v-if="clickMode === 'define_sequence' && seats[`m-${r-1}-${c-1}`]?.regionId === activeRegionId && layout.customSequences?.[activeRegionId]?.includes('m-'+(r-1)+'-'+(c-1))"
                                          class="absolute -top-2 -right-2 bg-indigo-600 text-white rounded-full w-[18px] h-[18px] flex items-center justify-center text-[9px] font-extrabold z-50 shadow-md border border-white">
                                         {{ layout.customSequences[activeRegionId].indexOf('m-'+(r-1)+'-'+(c-1)) + 1 }}
                                     </div>

                                     <div class="absolute top-0.5 left-1 flex items-center z-20 w-[90%]">
                                         <span class="text-[8px] font-bold pointer-events-none whitespace-nowrap flex-shrink-0" :style="{ color: isSeatBgDark('m', r-1, c-1) ? 'rgba(255,255,255,0.7)' : 'rgba(0,0,0,0.5)' }">{{ layout.main.rowLabels[r-1] }}排&nbsp;</span>
                                         <input type="text"
                                                :value="getSeatColLabelUI('m', r-1, c-1)"
                                                @change="setCustomLabel('m-'+(r-1)+'-'+(c-1), $event.target.value)"
                                                @click.stop
                                                class="flex-1 min-w-0 bg-transparent border-none outline-none text-[8px] font-bold p-0 m-0 pointer-events-auto hover:bg-black/10 focus:bg-white/50 rounded"
                                                :style="{ color: isDuplicateSeat('m', r-1, c-1) ? '#ef4444' : (isSeatBgDark('m', r-1, c-1) ? 'rgba(255,255,255,0.9)' : 'rgba(0,0,0,0.7)') }">
                                     </div>
                                     <div class="text-center w-full mt-2 pointer-events-none">
                                        <div class="font-bold text-[13px] truncate leading-tight">{{ getAttendeeFieldValue(seats[`m-${r-1}-${c-1}`].attendee, getSeatDisplayConfig('m', r-1, c-1).mainField) }}</div>
                                        <div v-if="getSeatDisplayConfig('m', r-1, c-1).subField" class="text-[8px] opacity-90 truncate leading-tight mt-0.5" :style="{ backgroundColor: isSeatBgDark('m', r-1, c-1) ? 'rgba(0,0,0,0.1)' : 'rgba(255,255,255,0.3)', padding: '0 2px', borderRadius: '2px', display: 'inline-block', maxWidth: '100%' }">{{ getAttendeeFieldValue(seats[`m-${r-1}-${c-1}`].attendee, getSeatDisplayConfig('m', r-1, c-1).subField) }}</div>
                                     </div>
                                 </div>

                                 <!-- 空位且定义连线排座时的角标 -->
                                 <div v-if="!seats[`m-${r-1}-${c-1}`]?.attendee && clickMode === 'define_sequence' && seats[`m-${r-1}-${c-1}`]?.regionId === activeRegionId && layout.customSequences?.[activeRegionId]?.includes('m-'+(r-1)+'-'+(c-1))"
                                      class="absolute -top-2 -right-2 bg-indigo-600 text-white rounded-full w-[18px] h-[18px] flex items-center justify-center text-[9px] font-extrabold z-50 shadow-md border border-white">
                                     {{ layout.customSequences[activeRegionId].indexOf('m-'+(r-1)+'-'+(c-1)) + 1 }}
                                 </div>
                            </div>

                            <div class="header-cell" @contextmenu.prevent="handleAxisRightClick($event, 'row', r-1, 'm')">
                                <input type="text" v-model="layout.main.rowLabels[r-1]" class="axis-input" @change="handleLabelChange('row', r-1, 'm')">
                            </div>
                        </template>

                        <div></div>
                        <div v-for="c in layout.main.cols" :key="'mbcol'+c" class="header-cell h-7" @contextmenu.prevent="handleAxisRightClick($event, 'col', c-1, 'm')">
                            <input type="text" v-model="layout.main.colLabels[c-1]" class="axis-input" @change="handleLabelChange('col', c-1, 'm')">
                        </div>
                        <div></div>
                    </div>

                    </div>  <!-- end 课桌式 classroom -->

                    <!-- ========== 长会议桌式布局 ========== -->
                    <div v-if="layout.layoutType === 'longtable'" class="longtable-wrapper" :style="{ '--table-color': layout.longtableConfig.tableColor }">
                        <!-- 桌外主位行（Grid 上方） -->
                        <div v-if="layout.longtableConfig.headSeats.enabled" class="longtable-head-row" style="margin-bottom:8px;">
                            <div class="longtable-head-group">
                                <span class="head-label">桌外主位左</span>
                                <div v-for="pos in layout.longtableConfig.headSeats.seatCount" :key="'hlt-'+pos"
                                     class="longtable-seat" style="width:64px;height:64px;"
                                     :class="seatClass('lt-head-left', pos-1, 0)"
                                     :style="seatStyleLongtable('head-left', pos-1)"
                                     @click="handleSeatClick('lt-head-left', pos-1, 0)"
                                     @contextmenu.prevent="handleSeatRightClick($event, 'lt-head-left', pos-1, 0)">
                                    <div v-if="!seats['lt-head-left-'+(pos-1)]?.attendee && seats['lt-head-left-'+(pos-1)]?.type === 'normal'" class="seat-empty-label">L{{ pos }}</div>
                                    <div v-else-if="seats['lt-head-left-'+(pos-1)]?.attendee" class="seat-occupied-content">
                                        <div class="font-bold text-[10px] truncate">{{ getAttendeeFieldValue(seats['lt-head-left-'+(pos-1)].attendee, getSeatDisplayConfig('lt-head-left', pos-1, 0).mainField) }}</div>
                                    </div>
                                </div>
                            </div>
                            <div class="longtable-head-group">
                                <span class="head-label">桌外主位右</span>
                                <div v-for="pos in layout.longtableConfig.headSeats.seatCount" :key="'hrt-'+pos"
                                     class="longtable-seat" style="width:64px;height:64px;"
                                     :class="seatClass('lt-head-right', pos-1, 0)"
                                     :style="seatStyleLongtable('head-right', pos-1)"
                                     @click="handleSeatClick('lt-head-right', pos-1, 0)"
                                     @contextmenu.prevent="handleSeatRightClick($event, 'lt-head-right', pos-1, 0)">
                                    <div v-if="!seats['lt-head-right-'+(pos-1)]?.attendee && seats['lt-head-right-'+(pos-1)]?.type === 'normal'" class="seat-empty-label">R{{ pos }}</div>
                                    <div v-else-if="seats['lt-head-right-'+(pos-1)]?.attendee" class="seat-occupied-content">
                                        <div class="font-bold text-[10px] truncate">{{ getAttendeeFieldValue(seats['lt-head-right-'+(pos-1)].attendee, getSeatDisplayConfig('lt-head-right', pos-1, 0).mainField) }}</div>
                                    </div>
                                </div>
                            </div>
                        </div>

                        <!-- Grid 主布局: 3 cols × 3 rows -->
                        <div class="longtable-scene" :style="longtableGridStyle">
                            <!-- Row 1: 上侧座位 (top), col 2 -->
                            <div v-if="layout.longtableConfig.sides.top.rows > 0" class="longtable-side-h" :style="{ gridRow: longtableTopRow, gridColumn: longtableTableGridCol, justifySelf: 'center' }">
                                <div class="longtable-side-label" style="margin-bottom:2px;">{{ layout.longtableConfig.sides.top.label }}</div>
                                <!-- row表示第几排, row 0 最靠近桌子(在下方), row N-1 最远(在上方) -->
                                <template v-for="rowIdx in layout.longtableConfig.sides.top.rows" :key="'t-r'+rowIdx">
                                    <div v-if="rowIdx-1 < layout.longtableConfig.sides.top.rows" style="display:flex;flex-direction:column;align-items:center;">
                                        <div class="longtable-row" :style="{ gap: layout.longtableConfig.sides.top.seatGap + 'px' }">
                                            <div v-for="colIdx in layout.longtableConfig.sides.top.seatsPerRow" :key="'ts-'+rowIdx+'-'+colIdx"
                                                 class="longtable-seat"
                                                 :class="seatClass('lt-top', (rowIdx-1), colIdx-1)"
                                                 :style="seatStyleLongtableRow('top', rowIdx-1, colIdx-1)"
                                                 @click="handleSeatClick('lt-top', (rowIdx-1), colIdx-1)"
                                                 @contextmenu.prevent="handleSeatRightClick($event, 'lt-top', (rowIdx-1), colIdx-1)">
                                                <div v-if="!seats['lt-top-'+(rowIdx-1)+'-'+(colIdx-1)]?.attendee && (!seats['lt-top-'+(rowIdx-1)+'-'+(colIdx-1)] || seats['lt-top-'+(rowIdx-1)+'-'+(colIdx-1)].type==='normal')" class="seat-empty-label">
                                                    <input type="text" :value="getSeatColLabelUI('lt-top', (rowIdx-1), colIdx-1)" @change="setCustomLabel('lt-top-'+(rowIdx-1)+'-'+(colIdx-1), $event.target.value)" @click.stop class="seat-label-input" placeholder="号">
                                                </div>
                                                <div v-else-if="seats['lt-top-'+(rowIdx-1)+'-'+(colIdx-1)]?.type === 'aisle'"></div>
                                                <div v-else-if="seats['lt-top-'+(rowIdx-1)+'-'+(colIdx-1)]?.type === 'disabled'"><i class="fa-solid fa-ban text-gray-300 text-lg"></i></div>
                                                <div v-else-if="seats['lt-top-'+(rowIdx-1)+'-'+(colIdx-1)]?.attendee" class="seat-occupied-content">
                                                    <div class="font-bold text-[11px] truncate leading-tight">{{ getAttendeeFieldValue(seats['lt-top-'+(rowIdx-1)+'-'+(colIdx-1)].attendee, getSeatDisplayConfig('lt-top', (rowIdx-1), colIdx-1).mainField) }}</div>
                                                </div>
                                            </div>
                                        </div>
                                        <div class="longtable-row-label" v-if="rowIdx-1 < layout.longtableConfig.sides.top.rows">第{{ rowIdx }}排</div>
                                    </div>
                                </template>
                            </div>

                            <!-- Row 2: 左侧座位 (left), col 1 -->
                            <div v-if="longtableHasLeft" class="longtable-side-v" :style="{ gridRow: longtableTableRow, gridColumn: '1' }">
                                <div class="longtable-side-label">{{ layout.longtableConfig.sides.left.label }}</div>
                                <!-- 每"排"是一个竖立座位列，row 0 最靠近桌子(在右侧)，row N-1 最远(在左侧) -->
                                <template v-for="rowIdx in layout.longtableConfig.sides.left.rows" :key="'l-r'+rowIdx">
                                    <div v-if="rowIdx-1 < layout.longtableConfig.sides.left.rows" style="display:flex;flex-direction:row;align-items:center;">
                                        <div class="longtable-col" :style="{ gap: layout.longtableConfig.sides.left.seatGap + 'px' }">
                                            <div v-for="colIdx in layout.longtableConfig.sides.left.seatsPerRow" :key="'ls-'+rowIdx+'-'+colIdx"
                                                 class="longtable-seat"
                                                 :class="seatClass('lt-left', (rowIdx-1), colIdx-1)"
                                                 :style="seatStyleLongtableRow('left', rowIdx-1, colIdx-1)"
                                                 @click="handleSeatClick('lt-left', (rowIdx-1), colIdx-1)"
                                                 @contextmenu.prevent="handleSeatRightClick($event, 'lt-left', (rowIdx-1), colIdx-1)">
                                                <div v-if="!seats['lt-left-'+(rowIdx-1)+'-'+(colIdx-1)]?.attendee && (!seats['lt-left-'+(rowIdx-1)+'-'+(colIdx-1)] || seats['lt-left-'+(rowIdx-1)+'-'+(colIdx-1)].type==='normal')" class="seat-empty-label">
                                                    <input type="text" :value="getSeatColLabelUI('lt-left', (rowIdx-1), colIdx-1)" @change="setCustomLabel('lt-left-'+(rowIdx-1)+'-'+(colIdx-1), $event.target.value)" @click.stop class="seat-label-input" placeholder="号">
                                                </div>
                                                <div v-else-if="seats['lt-left-'+(rowIdx-1)+'-'+(colIdx-1)]?.type === 'aisle'"></div>
                                                <div v-else-if="seats['lt-left-'+(rowIdx-1)+'-'+(colIdx-1)]?.type === 'disabled'"><i class="fa-solid fa-ban text-gray-300 text-sm"></i></div>
                                                <div v-else-if="seats['lt-left-'+(rowIdx-1)+'-'+(colIdx-1)]?.attendee" class="seat-occupied-content">
                                                    <div class="font-bold text-[9px] truncate leading-tight">{{ getAttendeeFieldValue(seats['lt-left-'+(rowIdx-1)+'-'+(colIdx-1)].attendee, getSeatDisplayConfig('lt-left', (rowIdx-1), colIdx-1).mainField) }}</div>
                                                </div>
                                            </div>
                                        </div>
                                        <div class="longtable-row-label" v-if="rowIdx-1 < layout.longtableConfig.sides.left.rows">第{{ rowIdx }}排</div>
                                    </div>
                                </template>
                            </div>

                            <!-- Row 2: 桌子本体, col 2 -->
                            <div class="longtable-table-wrap" :style="{ gridRow: longtableTableRow, gridColumn: longtableTableGridCol }">
                                <div class="longtable-table" :style="{ width: longtableTableWidth + 'px', height: longtableTableHeight + 'px' }">
                                    <span class="longtable-table-label">长 会 议 桌</span>
                                </div>
                            </div>

                            <!-- Row 2: 右侧座位 (right), col 3 -->
                            <div v-if="longtableHasRight" class="longtable-side-v" :style="{ gridRow: longtableTableRow, gridColumn: longtableGridCols }">
                                <template v-for="rowIdx in layout.longtableConfig.sides.right.rows" :key="'r-r'+rowIdx">
                                    <div v-if="rowIdx-1 < layout.longtableConfig.sides.right.rows" style="display:flex;flex-direction:row;align-items:center;">
                                        <div class="longtable-row-label" v-if="rowIdx-1 < layout.longtableConfig.sides.right.rows">第{{ rowIdx }}排</div>
                                        <div class="longtable-col" :style="{ gap: layout.longtableConfig.sides.right.seatGap + 'px' }">
                                            <div v-for="colIdx in layout.longtableConfig.sides.right.seatsPerRow" :key="'rs-'+rowIdx+'-'+colIdx"
                                                 class="longtable-seat"
                                                 :class="seatClass('lt-right', (rowIdx-1), colIdx-1)"
                                                 :style="seatStyleLongtableRow('right', rowIdx-1, colIdx-1)"
                                                 @click="handleSeatClick('lt-right', (rowIdx-1), colIdx-1)"
                                                 @contextmenu.prevent="handleSeatRightClick($event, 'lt-right', (rowIdx-1), colIdx-1)">
                                                <div v-if="!seats['lt-right-'+(rowIdx-1)+'-'+(colIdx-1)]?.attendee && (!seats['lt-right-'+(rowIdx-1)+'-'+(colIdx-1)] || seats['lt-right-'+(rowIdx-1)+'-'+(colIdx-1)].type==='normal')" class="seat-empty-label">
                                                    <input type="text" :value="getSeatColLabelUI('lt-right', (rowIdx-1), colIdx-1)" @change="setCustomLabel('lt-right-'+(rowIdx-1)+'-'+(colIdx-1), $event.target.value)" @click.stop class="seat-label-input" placeholder="号">
                                                </div>
                                                <div v-else-if="seats['lt-right-'+(rowIdx-1)+'-'+(colIdx-1)]?.type === 'aisle'"></div>
                                                <div v-else-if="seats['lt-right-'+(rowIdx-1)+'-'+(colIdx-1)]?.type === 'disabled'"><i class="fa-solid fa-ban text-gray-300 text-sm"></i></div>
                                                <div v-else-if="seats['lt-right-'+(rowIdx-1)+'-'+(colIdx-1)]?.attendee" class="seat-occupied-content">
                                                    <div class="font-bold text-[9px] truncate leading-tight">{{ getAttendeeFieldValue(seats['lt-right-'+(rowIdx-1)+'-'+(colIdx-1)].attendee, getSeatDisplayConfig('lt-right', (rowIdx-1), colIdx-1).mainField) }}</div>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </template>
                                <div class="longtable-side-label">{{ layout.longtableConfig.sides.right.label }}</div>
                            </div>

                            <!-- Row 3: 下侧座位 (bottom), col 2 -->
                            <div v-if="layout.longtableConfig.sides.bottom.rows > 0" class="longtable-side-h" :style="{ gridRow: longtableBottomRow, gridColumn: longtableTableGridCol, justifySelf: 'center' }">
                                <template v-for="rowIdx in layout.longtableConfig.sides.bottom.rows" :key="'b-r'+rowIdx">
                                    <div v-if="rowIdx-1 < layout.longtableConfig.sides.bottom.rows" style="display:flex;flex-direction:column;align-items:center;">
                                        <div class="longtable-row-label" v-if="rowIdx-1 < layout.longtableConfig.sides.bottom.rows">第{{ rowIdx }}排</div>
                                        <div class="longtable-row" :style="{ gap: layout.longtableConfig.sides.bottom.seatGap + 'px' }">
                                            <div v-for="colIdx in layout.longtableConfig.sides.bottom.seatsPerRow" :key="'bs-'+rowIdx+'-'+colIdx"
                                                 class="longtable-seat"
                                                 :class="seatClass('lt-bottom', (rowIdx-1), colIdx-1)"
                                                 :style="seatStyleLongtableRow('bottom', rowIdx-1, colIdx-1)"
                                                 @click="handleSeatClick('lt-bottom', (rowIdx-1), colIdx-1)"
                                                 @contextmenu.prevent="handleSeatRightClick($event, 'lt-bottom', (rowIdx-1), colIdx-1)">
                                                <div v-if="!seats['lt-bottom-'+(rowIdx-1)+'-'+(colIdx-1)]?.attendee && (!seats['lt-bottom-'+(rowIdx-1)+'-'+(colIdx-1)] || seats['lt-bottom-'+(rowIdx-1)+'-'+(colIdx-1)].type==='normal')" class="seat-empty-label">
                                                    <input type="text" :value="getSeatColLabelUI('lt-bottom', (rowIdx-1), colIdx-1)" @change="setCustomLabel('lt-bottom-'+(rowIdx-1)+'-'+(colIdx-1), $event.target.value)" @click.stop class="seat-label-input" placeholder="号">
                                                </div>
                                                <div v-else-if="seats['lt-bottom-'+(rowIdx-1)+'-'+(colIdx-1)]?.type === 'aisle'"></div>
                                                <div v-else-if="seats['lt-bottom-'+(rowIdx-1)+'-'+(colIdx-1)]?.type === 'disabled'"><i class="fa-solid fa-ban text-gray-300 text-lg"></i></div>
                                                <div v-else-if="seats['lt-bottom-'+(rowIdx-1)+'-'+(colIdx-1)]?.attendee" class="seat-occupied-content">
                                                    <div class="font-bold text-[11px] truncate leading-tight">{{ getAttendeeFieldValue(seats['lt-bottom-'+(rowIdx-1)+'-'+(colIdx-1)].attendee, getSeatDisplayConfig('lt-bottom', (rowIdx-1), colIdx-1).mainField) }}</div>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </template>
                                <div class="longtable-side-label" style="margin-top:2px;">{{ layout.longtableConfig.sides.bottom.label }}</div>
                            </div>
                        </div>
                    </div>

                    <!-- ========== 圆桌式布局 ========== -->
                    <div v-if="layout.layoutType === 'roundtable'" class="roundtable-wrapper">
                        <div v-for="t in layout.roundtableConfig.tables" :key="'rtable-'+t" class="roundtable-scene" :style="{ '--table-color': layout.roundtableConfig.tableColor, width: (Math.max(...layout.roundtableConfig.rings.map(r=>r.radius)) * 2 + 120) + 'px', height: (Math.max(...layout.roundtableConfig.rings.map(r=>r.radius)) * 2 + 120) + 'px' }">
                            <div class="roundtable-table" :style="{ width: ((layout.roundtableConfig.rings[0]?.radius || 130) * 2 - 40) + 'px', height: ((layout.roundtableConfig.rings[0]?.radius || 130) * 2 - 40) + 'px' }">
                                <span class="text-xs font-bold text-white opacity-80">{{ t > 1 ? '桌'+t : '圆桌' }}</span>
                            </div>
                            <template v-for="ring in layout.roundtableConfig.rings" :key="ring.id">
                                <div v-for="pos in ring.seatCount" :key="`${ring.id}-${pos-1}`"
                                     class="roundtable-seat"
                                     :class="[seatClass('rt-'+(t-1)+'-'+ring.id, pos-1, 0), {'roundtable-dual-principal': layout.roundtableConfig.dualPrincipal && t === 1 && ring.id === layout.roundtableConfig.rings[0].id && (pos === 1 || pos === 2)}]"
                                     :style="roundtableSeatStyle(t-1, ring, pos-1)"
                                     @click="handleSeatClick('rt-'+(t-1)+'-'+ring.id, pos-1, 0)"
                                     @contextmenu.prevent="handleSeatRightClick($event, 'rt-'+(t-1)+'-'+ring.id, pos-1, 0)">
                                    <div v-if="!seats['rt-'+(t-1)+'-'+ring.id+'-'+(pos-1)]?.attendee && seats['rt-'+(t-1)+'-'+ring.id+'-'+(pos-1)]?.type === 'normal'" class="seat-empty-label">{{ pos }}</div>
                                    <div v-else-if="seats['rt-'+(t-1)+'-'+ring.id+'-'+(pos-1)]?.type === 'aisle'"></div>
                                    <div v-else-if="seats['rt-'+(t-1)+'-'+ring.id+'-'+(pos-1)]?.type === 'disabled'"><i class="fa-solid fa-ban text-gray-300 text-lg"></i></div>
                                    <div v-else-if="seats['rt-'+(t-1)+'-'+ring.id+'-'+(pos-1)]?.attendee" class="seat-occupied-content">
                                        <div class="font-bold text-[11px] truncate leading-tight">{{ getAttendeeFieldValue(seats['rt-'+(t-1)+'-'+ring.id+'-'+(pos-1)].attendee, getSeatDisplayConfig('rt-'+(t-1)+'-'+ring.id, pos-1, 0).mainField) }}</div>
                                        <div v-if="getSeatDisplayConfig('rt-'+(t-1)+'-'+ring.id, pos-1, 0).subField" class="text-[7px] opacity-90 truncate">{{ getAttendeeFieldValue(seats['rt-'+(t-1)+'-'+ring.id+'-'+(pos-1)].attendee, getSeatDisplayConfig('rt-'+(t-1)+'-'+ring.id, pos-1, 0).subField) }}</div>
                                    </div>
                                </div>
                            </template>
                        </div>
                    </div>

                    <!-- ========== U 型桌式布局 ========== -->
                    <div v-if="layout.layoutType === 'utable'" class="utable-wrapper">
                        <div class="utable-scene" :class="{'leader-bottom': layout.utableConfig.leaderSide === 'bottom'}" :style="utableGridStyle">
                            <div class="utable-leader-strip" :style="{ gridColumn: '2', gridRow: layout.utableConfig.leaderSide === 'top' ? '1' : '3', gap: layout.utableConfig.leaderGap + 'px', flexDirection: 'column' }">
                                <div class="utable-leader-title">领导席 · 主位居中</div>
                                <div class="flex justify-center" :style="{gap: layout.utableConfig.leaderGap + 'px', flexDirection: layout.utableConfig.leaderSide === 'bottom' ? 'row-reverse' : 'row'}">
                                    <div v-for="pos in layout.utableConfig.leaderSeatCount" :key="'u-leader-'+(pos-1)"
                                         class="utable-seat" :class="seatClass('u-leader', pos-1, 0)" :style="seatStyle('u-leader', pos-1, 0)"
                                         @click="handleSeatClick('u-leader', pos-1, 0)" @contextmenu.prevent="handleSeatRightClick($event, 'u-leader', pos-1, 0)">
                                        <div v-if="!seats['u-leader-'+(pos-1)]?.attendee && seats['u-leader-'+(pos-1)]?.type === 'normal'" class="seat-empty-label">领导{{ getUtableLeaderOrderNumber(pos-1) }}</div>
                                        <div v-else-if="seats['u-leader-'+(pos-1)]?.type === 'aisle'"></div>
                                        <div v-else-if="seats['u-leader-'+(pos-1)]?.type === 'disabled'"><i class="fa-solid fa-ban text-gray-300 text-lg"></i></div>
                                        <div v-else-if="seats['u-leader-'+(pos-1)]?.attendee" class="seat-occupied-content">
                                            <div class="font-bold text-[11px] truncate leading-tight">{{ getAttendeeFieldValue(seats['u-leader-'+(pos-1)].attendee, getSeatDisplayConfig('u-leader', pos-1, 0).mainField) }}</div>
                                            <div v-if="getSeatDisplayConfig('u-leader', pos-1, 0).subField" class="text-[7px] opacity-90 truncate">{{ getAttendeeFieldValue(seats['u-leader-'+(pos-1)].attendee, getSeatDisplayConfig('u-leader', pos-1, 0).subField) }}</div>
                                        </div>
                                    </div>
                                </div>
                            </div>

                            <div class="utable-side-group left-side" :style="{ gridColumn: '1', gridRow: '2', gap: layout.utableConfig.sideGap + 'px' }">
                                <div v-for="column in layout.utableConfig.leftColumnCount" :key="'u-left-column-'+(column-1)" class="utable-side-strip" :class="{'leader-at-bottom': layout.utableConfig.leaderSide === 'bottom'}" :style="{gap: layout.utableConfig.sideGap + 'px'}">
                                    <div v-for="pos in layout.utableConfig.leftSeatCount" :key="'u-left-'+(column-1)+'-'+(pos-1)"
                                         class="utable-seat" :class="seatClass('u-left-'+(column-1), pos-1, 0)" :style="seatStyle('u-left-'+(column-1), pos-1, 0)"
                                         @click="handleSeatClick('u-left-'+(column-1), pos-1, 0)" @contextmenu.prevent="handleSeatRightClick($event, 'u-left-'+(column-1), pos-1, 0)">
                                        <div v-if="!seats['u-left-'+(column-1)+'-'+(pos-1)]?.attendee && seats['u-left-'+(column-1)+'-'+(pos-1)]?.type === 'normal'" class="seat-empty-label">左{{ column }}-{{ pos }}</div>
                                        <div v-else-if="seats['u-left-'+(column-1)+'-'+(pos-1)]?.type === 'aisle'"></div>
                                        <div v-else-if="seats['u-left-'+(column-1)+'-'+(pos-1)]?.type === 'disabled'"><i class="fa-solid fa-ban text-gray-300 text-lg"></i></div>
                                        <div v-else-if="seats['u-left-'+(column-1)+'-'+(pos-1)]?.attendee" class="seat-occupied-content">
                                            <div class="font-bold text-[11px] truncate leading-tight">{{ getAttendeeFieldValue(seats['u-left-'+(column-1)+'-'+(pos-1)].attendee, getSeatDisplayConfig('u-left-'+(column-1), pos-1, 0).mainField) }}</div>
                                            <div v-if="getSeatDisplayConfig('u-left-'+(column-1), pos-1, 0).subField" class="text-[7px] opacity-90 truncate">{{ getAttendeeFieldValue(seats['u-left-'+(column-1)+'-'+(pos-1)].attendee, getSeatDisplayConfig('u-left-'+(column-1), pos-1, 0).subField) }}</div>
                                        </div>
                                    </div>
                                </div>
                            </div>

                            <div class="utable-open-center" :style="{ gridColumn: '2', gridRow: '2' }">开放区域 · 无桌面</div>

                            <div class="utable-side-group right-side" :style="{ gridColumn: '3', gridRow: '2', gap: layout.utableConfig.sideGap + 'px' }">
                                <div v-for="column in layout.utableConfig.rightColumnCount" :key="'u-right-column-'+(column-1)" class="utable-side-strip" :class="{'leader-at-bottom': layout.utableConfig.leaderSide === 'bottom'}" :style="{gap: layout.utableConfig.sideGap + 'px'}">
                                    <div v-for="pos in layout.utableConfig.rightSeatCount" :key="'u-right-'+(column-1)+'-'+(pos-1)"
                                         class="utable-seat" :class="seatClass('u-right-'+(column-1), pos-1, 0)" :style="seatStyle('u-right-'+(column-1), pos-1, 0)"
                                         @click="handleSeatClick('u-right-'+(column-1), pos-1, 0)" @contextmenu.prevent="handleSeatRightClick($event, 'u-right-'+(column-1), pos-1, 0)">
                                        <div v-if="!seats['u-right-'+(column-1)+'-'+(pos-1)]?.attendee && seats['u-right-'+(column-1)+'-'+(pos-1)]?.type === 'normal'" class="seat-empty-label">右{{ column }}-{{ pos }}</div>
                                        <div v-else-if="seats['u-right-'+(column-1)+'-'+(pos-1)]?.type === 'aisle'"></div>
                                        <div v-else-if="seats['u-right-'+(column-1)+'-'+(pos-1)]?.type === 'disabled'"><i class="fa-solid fa-ban text-gray-300 text-lg"></i></div>
                                        <div v-else-if="seats['u-right-'+(column-1)+'-'+(pos-1)]?.attendee" class="seat-occupied-content">
                                            <div class="font-bold text-[11px] truncate leading-tight">{{ getAttendeeFieldValue(seats['u-right-'+(column-1)+'-'+(pos-1)].attendee, getSeatDisplayConfig('u-right-'+(column-1), pos-1, 0).mainField) }}</div>
                                            <div v-if="getSeatDisplayConfig('u-right-'+(column-1), pos-1, 0).subField" class="text-[7px] opacity-90 truncate">{{ getAttendeeFieldValue(seats['u-right-'+(column-1)+'-'+(pos-1)].attendee, getSeatDisplayConfig('u-right-'+(column-1), pos-1, 0).subField) }}</div>
                                        </div>
                                    </div>
                                </div>
                            </div>

                            <div class="utable-stage" :style="{ gridColumn: '2', gridRow: layout.utableConfig.leaderSide === 'top' ? '3' : '1' }">
                                <i class="fa-solid fa-display"></i><span>{{ layout.utableConfig.stageLabel || 'PPT／演讲台' }}</span>
                            </div>
                        </div>
                    </div>

                </div>
            </div>

            <!-- 座位右键菜单 -->
            <div v-if="seatMenu.show" class="absolute bg-white border border-gray-200 shadow-2xl rounded-lg z-50 w-52 py-1 font-sans" :style="{ top: seatMenu.y + 'px', left: seatMenu.x + 'px' }">
                <div class="px-3 py-2 border-b bg-gray-100 flex justify-between items-center text-xs font-bold text-indigo-800">
                    <span>📍 {{seatMenuLabel}}</span>
                    <i class="fa-solid fa-xmark cursor-pointer text-gray-400 hover:text-gray-800" @click="closeContextMenus"></i>
                </div>
                <template v-if="seatMenu.hasAttendee">
                    <button @click="clearSeat(seatMenu.zone, seatMenu.r, seatMenu.c)" class="w-full text-left px-4 py-2.5 text-sm hover:bg-gray-100 font-medium"><i class="fa-solid fa-eraser text-gray-400 w-5 text-center"></i> 清空该座位</button>
                    <button @click="removeAndShift(seatMenu.zone, seatMenu.r, seatMenu.c)" class="w-full text-left px-4 py-2.5 text-sm text-red-600 hover:bg-red-50 font-bold border-t"><i class="fa-solid fa-angles-left text-red-400 w-5 text-center"></i> 缺席并让后方顺移补位</button>
                </template>
                <template v-else>
                    <div class="px-2 py-1.5 border-b"><input type="text" v-model="seatMenuSearch" @click.stop placeholder="搜索未排座人员强绑..." class="w-full border border-gray-300 rounded px-2 py-1.5 text-xs focus:outline-none focus:border-indigo-500"></div>
                    <div class="max-h-48 overflow-y-auto">
                        <div v-for="att in seatMenuFilteredAtts" :key="att.id" @click="forceAssign(seatMenu.zone, seatMenu.r, seatMenu.c, att)" class="px-4 py-2 hover:bg-indigo-50 cursor-pointer flex flex-col border-b border-gray-50 last:border-0">
                            <span class="font-bold text-sm text-gray-800">{{att.name}}</span><span class="text-[9px] text-gray-500">{{att.unit}}</span>
                        </div>
                        <div v-if="seatMenuFilteredAtts.length===0" class="p-3 text-center text-xs text-gray-400">无匹配人员</div>
                    </div>
                </template>
                <div class="border-t my-1"></div>
                <button @click="openSeatDisplayEdit(seatMenu.zone, seatMenu.r, seatMenu.c)" class="w-full text-left px-4 py-2.5 text-sm hover:bg-indigo-50 font-medium text-indigo-700"><i class="fa-solid fa-eye text-indigo-400 w-5 text-center"></i> 单独设置显示字段</button>
                <button v-if="seatMenuHasDisplayConfig" @click="resetSeatDisplayConfig(seatMenu.zone, seatMenu.r, seatMenu.c)" class="w-full text-left px-4 py-2.5 text-xs hover:bg-gray-100 text-gray-500"><i class="fa-solid fa-rotate-left text-red-400 w-5 text-center"></i> 恢复使用全局设置</button>
            </div>

            <!-- 座位显示字段编辑弹窗 -->
            <div v-if="seatDisplayEdit.show" class="absolute bg-white border border-gray-300 shadow-2xl rounded-lg z-50 w-64 p-4 font-sans" :style="{ top: (seatDisplayEdit.show ? Math.min((seatMenu.show ? seatMenu.y : 200), document.body.clientHeight - 320) : 0) + 'px', left: (seatDisplayEdit.show ? Math.min((seatMenu.show ? seatMenu.x + 220 : 300), document.body.clientWidth - 280) : 0) + 'px' }">
                <div class="flex justify-between items-center mb-3">
                    <span class="text-sm font-bold text-indigo-800">📋 座位显示字段</span>
                    <i class="fa-solid fa-xmark cursor-pointer text-gray-400 hover:text-gray-800" @click="seatDisplayEdit.show = false"></i>
                </div>
                <div class="space-y-2">
                    <div>
                        <label class="text-[10px] text-gray-500 block">主字段 (大字显示)</label>
                        <select v-model="seatDisplayEdit.mainField" class="w-full border rounded p-1.5 text-xs">
                            <option value="姓名">姓名</option>
                            <option value="单位">单位</option>
                            <option value="类别">类别</option>
                            <option value="电话">电话</option>
                            <option v-for="key in extraKeys" :key="'se_mf_'+key" :value="key">{{key}}</option>
                        </select>
                    </div>
                    <div>
                        <label class="text-[10px] text-gray-500 block">副字段 (小字显示)</label>
                        <select v-model="seatDisplayEdit.subField" class="w-full border rounded p-1.5 text-xs">
                            <option value="">无</option>
                            <option value="姓名">姓名</option>
                            <option value="单位">单位</option>
                            <option value="类别">类别</option>
                            <option value="电话">电话</option>
                            <option v-for="key in extraKeys" :key="'se_sf_'+key" :value="key">{{key}}</option>
                        </select>
                    </div>
                    <div class="flex space-x-2 pt-1">
                        <button @click="saveSeatDisplayConfig" class="flex-1 bg-indigo-500 text-white py-1.5 rounded text-xs font-bold hover:bg-indigo-600">保存</button>
                        <button @click="seatDisplayEdit.show = false" class="flex-1 bg-gray-100 text-gray-600 py-1.5 rounded text-xs hover:bg-gray-200">取消</button>
                    </div>
                </div>
            </div>

            <!-- 坐标轴右键菜单 -->
            <div v-if="axisMenu.show" class="absolute bg-white border border-gray-200 shadow-2xl rounded-lg z-50 w-48 py-1 font-sans" :style="{ top: axisMenu.y + 'px', left: axisMenu.x + 'px' }">
                <div class="px-3 py-2 border-b bg-gray-100 flex justify-between items-center text-[10px] font-bold text-gray-500">
                    <span>操作该{{axisMenu.type==='row'?'排':'列'}}</span>
                    <i class="fa-solid fa-xmark cursor-pointer text-gray-400 hover:text-gray-800" @click="closeContextMenus"></i>
                </div>

                <button @click="modifyAxis('insert_before')" class="w-full text-left px-4 py-2 text-xs hover:bg-gray-100">➕ 在前方插入一{{axisMenu.type==='row'?'排':'列'}}</button>
                <button @click="modifyAxis('insert_after')" class="w-full text-left px-4 py-2 text-xs hover:bg-gray-100">➕ 在后方插入一{{axisMenu.type==='row'?'排':'列'}}</button>
                <div class="border-t my-1"></div>
                <button @click="setAxisStatus('aisle')" class="w-full text-left px-4 py-2 text-xs hover:bg-gray-100">🟩 全设为走廊</button>
                <button @click="setAxisStatus('disabled')" class="w-full text-left px-4 py-2 text-xs hover:bg-gray-100 text-red-600">🚫 全设为废座</button>
                <button @click="setAxisStatus('normal')" class="w-full text-left px-4 py-2 text-xs hover:bg-gray-100 font-bold">🔄 恢复正常排座</button>
                <div class="border-t my-1"></div>
                <button @click="modifyAxis('delete')" class="w-full text-left px-4 py-2 text-xs hover:bg-red-50 text-red-600 font-bold"><i class="fa-solid fa-trash-can mr-1"></i> 删除本{{axisMenu.type==='row'?'排':'列'}} (自动并拢)</button>
            </div>

            <!-- 人员 Modal -->
            <div v-if="showModal" class="fixed inset-0 bg-gray-900 bg-opacity-50 flex items-center justify-center z-50 backdrop-blur-sm">
                <div class="bg-white p-6 rounded-xl shadow-2xl w-[500px] transform transition-all max-h-[90vh] overflow-y-auto">
                    <h3 class="font-extrabold text-xl mb-4 text-gray-800">{{ modalForm.id ? '✏️ 编辑人员' : '✨ 增加人员' }}</h3>
                    <div class="space-y-3">
                        <div><label class="text-[10px] font-bold text-gray-500">姓名 *</label><input type="text" v-model="modalForm.name" class="w-full border-b-2 border-gray-300 focus:border-indigo-500 bg-gray-50 p-2 text-sm outline-none"></div>
                        <div><label class="text-[10px] font-bold text-gray-500">单位</label><input type="text" v-model="modalForm.unit" class="w-full border-b-2 border-gray-300 focus:border-indigo-500 bg-gray-50 p-2 text-sm outline-none"></div>
                        <div class="grid grid-cols-2 gap-2">
                            <div><label class="text-[10px] font-bold text-gray-500">类别</label><input type="text" v-model="modalForm.category" class="w-full border-b-2 border-gray-300 focus:border-indigo-500 bg-gray-50 p-2 text-sm outline-none"></div>
                            <div><label class="text-[10px] font-bold text-gray-500">电话</label><input type="text" v-model="modalForm.phone" class="w-full border-b-2 border-gray-300 focus:border-indigo-500 bg-gray-50 p-2 text-sm outline-none"></div>
                        </div>
                        <div>
                            <label class="text-[10px] font-bold text-gray-500">主宾身份（长桌/圆桌使用）</label>
                            <select v-model="modalForm.party_role" class="w-full border-b-2 border-gray-300 focus:border-indigo-500 bg-gray-50 p-2 text-sm outline-none">
                                <option value="unassigned">未分类</option>
                                <option value="host">主方（上级）</option>
                                <option value="guest">宾方（下级）</option>
                            </select>
                        </div>
                        <div v-if="extraKeys.length > 0" class="border-t pt-3 mt-3">
                            <label class="text-[10px] font-bold text-gray-500 block mb-2">额外字段</label>
                            <div v-for="key in extraKeys" :key="key" class="mb-2">
                                <label class="text-[9px] font-bold text-gray-400">{{key}}</label>
                                <input type="text" v-model="modalForm.extra[key]" class="w-full border-b border-gray-200 focus:border-indigo-500 bg-gray-50 p-1.5 text-xs outline-none">
                            </div>
                        </div>
                    </div>
                    <div class="mt-6 flex justify-end space-x-2">
                        <button @click="showModal = false" class="px-4 py-2 bg-gray-100 text-gray-700 font-bold rounded hover:bg-gray-200">取消</button>
                        <button @click="saveModalForm" class="px-4 py-2 bg-indigo-600 text-white font-bold rounded shadow hover:bg-indigo-700">保存</button>
                    </div>
                </div>
            </div>

        </div>
    </div>

    <script>
        const { createApp, ref, computed, reactive, onMounted, watch } = Vue;

        createApp({
            setup() {
                const tabs = [
                    { id: 'layout', name: '会场设置', icon: 'fa-solid fa-border-all' },
                    { id: 'attendees', name: '人员管理', icon: 'fa-solid fa-users' },
                    { id: 'assign', name: '区域排座', icon: 'fa-solid fa-bolt' },
                    { id: 'export', name: '导出发布', icon: 'fa-solid fa-paper-plane' }
                ];
                const activeTab = ref('layout');

                const meetings = ref([]);
                const currentMeetingId = ref('');
                const templates = ref([]);
                const selectedTemplateId = ref('');
                const saveStatus = ref('');
                const history = ref([]);

                const layout = reactive({
                    layoutType: 'classroom',
                    podium: { rows: 1, cols: 8, rowLabels: [], colLabels: [], color: '#f97316', textColor: '#ffffff' },
                    main: { rows: 10, cols: 20, rowLabels: [], colLabels: [], color: '#ffffff', textColor: '#000000' },
                    longtableConfig: {
                        sides: {
                            top:    { rows: 1, seatsPerRow: 6, seatGap: 8, label: '上侧' },
                            bottom: { rows: 1, seatsPerRow: 6, seatGap: 8, label: '下侧' },
                            left:   { rows: 1, seatsPerRow: 3, seatGap: 8, label: '左侧' },
                            right:  { rows: 1, seatsPerRow: 3, seatGap: 8, label: '右侧' }
                        },
                        hostSide: 'bottom',
                        guestSide: 'top',
                        partySeatRule: 'center_out',
                        headSeats: { enabled: false, seatCount: 1 },
                        tableColor: '#8B4513'
                    },
                    roundtableConfig: {
                        rings: [
                            { id: 'ring-0', label: '主桌', seatCount: 12, radius: 150 }
                        ],
                        tableColor: '#8B4513',
                        tables: 1,
                        dualPrincipal: false
                    },
                    utableConfig: {
                        leaderSide: 'top',
                        leaderSeatCount: 5,
                        leftColumnCount: 1,
                        leftSeatCount: 6,
                        rightColumnCount: 1,
                        rightSeatCount: 6,
                        leaderGap: 8,
                        sideGap: 8,
                        stageLabel: 'PPT／演讲台',
                        accentColor: '#475569'
                    },
                    regionTypes: [
                        { id: '1', name: '核心内场', color: '#fca5a5' },
                        { id: '2', name: '嘉宾区', color: '#fef08a' }
                    ],
                    customSequences: {},
                    sms_template: '(参会通知){姓名}您好，您的座位在 {排号}排 {座号}，请准时入场。',
                    export_template: '{姓名}\\n({单位})',
                    seatDisplayConfig: { mainField: '姓名', subField: '单位' }
                });
                const seats = reactive({});
                const regionAssignmentRecords = reactive({});
                const attendees = ref([]);

                const clickMode = ref('none');
                const activeRegionId = ref('1');
                const layoutName = ref('');
                const meetingNameInput = ref('');
                const currentMeetingObj = computed(() => meetings.value.find(m => m.id === currentMeetingId.value) || {});

                const batchSelectState = reactive({
                    hasStart: false,
                    startZone: '',
                    startR: -1,
                    startC: -1
                });

                const attendeeSearch = ref('');
                const attendeeRoleFilter = ref('all');
                const selectedAttIds = ref([]);
                const dragTargetIndex = ref(-1);

                const assignSearch = ref('');
                const assignRoleFilter = ref('all');
                const assignConfig = reactive({ rule: 'center_lr', targetZone: 'all_m' });
                const resetAssignmentForLayout = () => {
                    assignSelectedAtts.value = [];
                    assignRoleFilter.value = 'all';
                    if (layout.layoutType === 'longtable') {
                        assignConfig.targetZone = 'party_auto';
                        assignConfig.rule = layout.longtableConfig.partySeatRule || 'center_out';
                    } else if (layout.layoutType === 'roundtable') {
                        assignConfig.targetZone = 'party_auto';
                        assignConfig.rule = 'banquet';
                    } else if (layout.layoutType === 'utable') {
                        assignConfig.targetZone = 'u_auto';
                        assignConfig.rule = 'leader_near';
                    } else {
                        assignConfig.targetZone = 'all_m';
                        assignConfig.rule = 'center_lr';
                    }
                };
                const assignSelectedAtts = ref([]);
                const assignRangeInput = ref('');

                const showModal = ref(false);
                const modalForm = reactive({id: null, name: '', unit: '', category: '', phone: '', party_role: 'unassigned', extra: {}});

                const seatMenu = reactive({ show: false, x: 0, y: 0, zone: '', r: -1, c: -1, hasAttendee: false });
                const seatMenuSearch = ref('');
                const axisMenu = reactive({ show: false, x: 0, y: 0, zone: '', type: '', index: -1 });
                const seatDisplayEdit = reactive({ show: false, zone: '', r: -1, c: -1, mainField: '姓名', subField: '' });

                const getContrastYIQ = (hexcolor) => {
                    if(!hexcolor) return '#ffffff';
                    hexcolor = hexcolor.replace("#", "");
                    if(hexcolor.length === 3) hexcolor = hexcolor.split('').map(x=>x+x).join('');
                    var r = parseInt(hexcolor.substr(0,2),16);
                    var g = parseInt(hexcolor.substr(2,2),16);
                    var b = parseInt(hexcolor.substr(4,2),16);
                    var yiq = ((r*299)+(g*587)+(b*114))/1000;
                    return (yiq >= 128) ? '#1f2937' : '#ffffff';
                };

                const isSeatBgDark = (zone, r, c) => {
                    const s = seats[makeSeatKey(zone, r, c)];
                    if(!s) return false;
                    let hex = '#f9fafb';
                    if(s.attendee) {
                        if (zone === 'p') hex = layout.podium.color;
                        else if (zone === 'm') hex = layout.main.color;
                        else hex = '#6366f1';
                    } else if (s.regionId) {
                        const reg = layout.regionTypes.find(rg => rg.id === s.regionId);
                        if(reg) hex = reg.color;
                    }
                    return getContrastYIQ(hex) === '#ffffff';
                };

                watch(() => layout.podium.color, (newVal) => { layout.podium.textColor = getContrastYIQ(newVal); });
                watch(() => layout.main.color, (newVal) => { layout.main.textColor = getContrastYIQ(newVal); });

                let saveTimer = null;
                const autoSave = () => {
                    if(!currentMeetingId.value) return;
                    saveStatus.value = '保存中...';
                    clearTimeout(saveTimer);
                    saveTimer = setTimeout(async () => {
                        await fetch(`/api/meetings/${currentMeetingId.value}/save`, {
                            method: 'POST', headers: {'Content-Type': 'application/json'},
                            body: JSON.stringify({layout: layout, seats: seats, records: regionAssignmentRecords, sms_template: layout.sms_template})
                        });
                        saveStatus.value = '已保存';
                        setTimeout(() => saveStatus.value='', 2000);
                    }, 800);
                };

                watch(seats, autoSave, {deep: true});
                watch(layout, autoSave, {deep: true});

                const loadGlobalData = async () => {
                    let res = await fetch('/api/meetings'); meetings.value = await res.json();
                    res = await fetch('/api/templates'); templates.value = await res.json();
                };

                const createMeeting = async () => {
                    const name = prompt('请输入新会议名称：', '新建会议 ' + new Date().toLocaleDateString());
                    if(!name) return;
                    const res = await fetch('/api/meetings', {method: 'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({name})});
                    const data = await res.json();
                    await loadGlobalData();
                    currentMeetingId.value = data.id;
                    loadMeetingState();
                };

                const deleteCurrentMeeting = async () => {
                    if(!currentMeetingId.value) return;
                    if(!confirm('确定永久删除当前会议及所有排座数据吗？此操作不可恢复！')) return;
                    await fetch(`/api/meetings/${currentMeetingId.value}/delete`, { method: 'DELETE' });
                    currentMeetingId.value = '';
                    await loadGlobalData();
                };

                const renameMeeting = async () => {
                    if(!currentMeetingId.value || !meetingNameInput.value) return;
                    await fetch(`/api/meetings/${currentMeetingId.value}/rename`, {
                        method: 'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({name: meetingNameInput.value})
                    });
                    const m = meetings.value.find(x => x.id === currentMeetingId.value);
                    if (m) m.name = meetingNameInput.value;
                };

                const loadMeetingState = async () => {
                    if(!currentMeetingId.value) return;
                    activeTab.value = 'layout';
                    const res = await fetch(`/api/meetings/${currentMeetingId.value}`);
                    const data = await res.json();

                    meetingNameInput.value = data.name;
                    const l = data.layout ? JSON.parse(data.layout) : null;
                    if(l && l.main) Object.assign(layout, l);

                    // 向后兼容：旧数据没有 layoutType 和新布局配置
                    if(!layout.layoutType) layout.layoutType = 'classroom';
                    if(layout.layoutType === 'squaretable') layout.layoutType = 'longtable';
                    if(!layout.longtableConfig) layout.longtableConfig = {
                        sides: { top: { rows: 1, seatsPerRow: 6, seatGap: 8, label: '上侧' }, bottom: { rows: 1, seatsPerRow: 6, seatGap: 8, label: '下侧' }, left: { rows: 1, seatsPerRow: 3, seatGap: 8, label: '左侧' }, right: { rows: 1, seatsPerRow: 3, seatGap: 8, label: '右侧' } },
                        hostSide: 'bottom', guestSide: 'top', partySeatRule: 'center_out', headSeats: { enabled: false, seatCount: 1 }, tableColor: '#8B4513'
                    };
                    // 向后兼容：旧数组格式转新对象格式
                    if(Array.isArray(layout.longtableConfig.sides)) {
                        const old = layout.longtableConfig.sides;
                        layout.longtableConfig.sides = {
                            top: { rows: 1, seatsPerRow: 6, seatGap: 8, label: '上侧' },
                            bottom: { rows: 1, seatsPerRow: 6, seatGap: 8, label: '下侧' },
                            left: { rows: 1, seatsPerRow: 3, seatGap: 8, label: '左侧' },
                            right: { rows: 1, seatsPerRow: 3, seatGap: 8, label: '右侧' }
                        };
                        old.forEach(s => {
                            if(s.direction === 'top' || s.direction === 'bottom') {
                                layout.longtableConfig.sides[s.direction].seatsPerRow = s.seatCount || 5;
                                layout.longtableConfig.sides[s.direction].label = s.label || '';
                            }
                        });
                    }
                    const longtableSideDefaults = {
                        top: { rows: 1, seatsPerRow: 6, seatGap: 8, label: '上侧' },
                        bottom: { rows: 1, seatsPerRow: 6, seatGap: 8, label: '下侧' },
                        left: { rows: 1, seatsPerRow: 3, seatGap: 8, label: '左侧' },
                        right: { rows: 1, seatsPerRow: 3, seatGap: 8, label: '右侧' }
                    };
                    if(!layout.longtableConfig.sides || Array.isArray(layout.longtableConfig.sides)) {
                        layout.longtableConfig.sides = {};
                    }
                    Object.keys(longtableSideDefaults).forEach(side => {
                        if(!layout.longtableConfig.sides[side]) {
                            layout.longtableConfig.sides[side] = { ...longtableSideDefaults[side] };
                        } else if(layout.longtableConfig.sides[side].seatGap === undefined || layout.longtableConfig.sides[side].seatGap === null) {
                            layout.longtableConfig.sides[side].seatGap = longtableSideDefaults[side].seatGap;
                        }
                    });
                    if(!['top', 'bottom'].includes(layout.longtableConfig.hostSide)) {
                        layout.longtableConfig.hostSide = 'bottom';
                    }
                    if(!['top', 'bottom'].includes(layout.longtableConfig.guestSide) || layout.longtableConfig.guestSide === layout.longtableConfig.hostSide) {
                        layout.longtableConfig.guestSide = layout.longtableConfig.hostSide === 'top' ? 'bottom' : 'top';
                    }
                    if(!['center_out', 'ltr', 'rtl', 'snake'].includes(layout.longtableConfig.partySeatRule)) {
                        layout.longtableConfig.partySeatRule = 'center_out';
                    }
                    if(!layout.roundtableConfig) layout.roundtableConfig = {
                        rings: [{ id: 'ring-0', label: '主桌', seatCount: 12, radius: 150 }],
                        tableColor: '#8B4513', tables: 1, dualPrincipal: false
                    };
                    if(!Array.isArray(layout.roundtableConfig.rings) || layout.roundtableConfig.rings.length === 0) {
                        layout.roundtableConfig.rings = [{ id: 'ring-0', label: '主桌', seatCount: 12, radius: 150 }];
                    }
                    if(layout.roundtableConfig.dualPrincipal === undefined) layout.roundtableConfig.dualPrincipal = false;
                    const utableDefaults = {
                        leaderSide: 'top', leaderSeatCount: 5, leftColumnCount: 1, leftSeatCount: 6, rightColumnCount: 1, rightSeatCount: 6,
                        leaderGap: 8, sideGap: 8, stageLabel: 'PPT／演讲台', accentColor: '#475569'
                    };
                    if(!layout.utableConfig) layout.utableConfig = { ...utableDefaults };
                    Object.keys(utableDefaults).forEach(key => {
                        if(layout.utableConfig[key] === undefined || layout.utableConfig[key] === null) layout.utableConfig[key] = utableDefaults[key];
                    });
                    if(!['top', 'bottom'].includes(layout.utableConfig.leaderSide)) layout.utableConfig.leaderSide = 'top';
                    if(!layout.podium.color) layout.podium.color = '#f97316';
                    if(!layout.podium.textColor) layout.podium.textColor = '#ffffff';
                    if(!layout.main.color) layout.main.color = '#ffffff';
                    if(!layout.main.textColor) layout.main.textColor = '#000000';
                    if(!layout.export_template) layout.export_template = '{姓名}\\n({单位})';
                    if(!layout.customSequences) layout.customSequences = {};
                    if(!layout.longtableConfig.headSeats) layout.longtableConfig.headSeats = { enabled: false, seatCount: 1 };
                    if(!layout.seatDisplayConfig) layout.seatDisplayConfig = { mainField: '姓名', subField: '单位' };

                    for(let k in seats) delete seats[k];
                    const st = data.seats ? JSON.parse(data.seats) : null;
                    if(st) Object.assign(seats, st);

                    for(let k in regionAssignmentRecords) delete regionAssignmentRecords[k];
                    const rec = data.records ? JSON.parse(data.records) : null;
                    if(rec) Object.assign(regionAssignmentRecords, rec);

                    history.value = [];
                    await fetchAttendees();
                    initSeats(false);
                    resetAssignmentForLayout();
                };

                const fetchAttendees = async () => {
                    if(!currentMeetingId.value) return;
                    const res = await fetch(`/api/meetings/${currentMeetingId.value}/attendees`);
                    attendees.value = await res.json();
                };

                const getGlobalIndex = (att) => {
                    return attendees.value.findIndex(a => a.id === att.id) + 1;
                };

                const normalizePartyRole = (value) => {
                    if (value === 'host' || value === 'guest') return value;
                    return 'unassigned';
                };
                const partyRoleLabel = (value) => ({
                    host: '主方', guest: '宾方', unassigned: '未分类'
                })[normalizePartyRole(value)];
                const partyRoleClass = (value) => ({
                    host: 'bg-amber-100 text-amber-800 border-amber-300',
                    guest: 'bg-blue-100 text-blue-800 border-blue-300',
                    unassigned: 'bg-gray-100 text-gray-500 border-gray-300'
                })[normalizePartyRole(value)];
                const isPartyLayout = computed(() => layout.layoutType === 'longtable' || layout.layoutType === 'roundtable');
                const partyRoleRank = { host: 0, guest: 1, unassigned: 2 };
                const partySortedAttendees = computed(() => {
                    if (!isPartyLayout.value) return attendees.value;
                    const originalOrder = new Map(attendees.value.map((att, index) => [att.id, index]));
                    return [...attendees.value].sort((a, b) => {
                        const roleDiff = partyRoleRank[normalizePartyRole(a.party_role)] - partyRoleRank[normalizePartyRole(b.party_role)];
                        if (roleDiff !== 0) return roleDiff;
                        return originalOrder.get(a.id) - originalOrder.get(b.id);
                    });
                });
                const partyCounts = computed(() => {
                    const counts = { host: 0, guest: 0, unassigned: 0 };
                    attendees.value.forEach(att => counts[normalizePartyRole(att.party_role)]++);
                    return counts;
                });
                const getAttendeeScope = (att) => {
                    if (!isPartyLayout.value) return attendees.value;
                    const role = normalizePartyRole(att.party_role);
                    return partySortedAttendees.value.filter(item => normalizePartyRole(item.party_role) === role);
                };
                const getAttendeeOrder = (att) => getAttendeeScope(att).findIndex(item => item.id === att.id) + 1;
                const isFirstInAttendeeScope = (att) => getAttendeeOrder(att) <= 1;
                const isLastInAttendeeScope = (att) => getAttendeeOrder(att) >= getAttendeeScope(att).length;

                const getAttendeeFieldValue = (att, fieldName) => {
                    if (!att) return '';
                    if (fieldName === '姓名') return att.name || '';
                    if (fieldName === '单位') return att.unit || '';
                    if (fieldName === '类别') return att.category || '';
                    if (fieldName === '电话') return att.phone || '';
                    if (att.extra && typeof att.extra === 'object') {
                        return att.extra[fieldName] || '';
                    }
                    if (att.extra && typeof att.extra === 'string') {
                        try {
                            const extraObj = JSON.parse(att.extra);
                            if (extraObj && extraObj[fieldName]) return extraObj[fieldName];
                        } catch(e) {}
                    }
                    return '';
                };

                const getSeatDisplayConfig = (zone, r, c) => {
                    const seat = seats[makeSeatKey(zone, r, c)];
                    if (seat && seat.displayConfig) return seat.displayConfig;
                    return layout.seatDisplayConfig;
                };

                // 主席台默认以中心偏左为主位，随后按“右、左”交替向外展开。
                const getPodiumColumnOrder = (count) => {
                    if(count <= 0) return [];
                    const centerLeft = Math.floor((count - 1) / 2);
                    const order = [centerLeft];
                    for(let offset = 1; order.length < count; offset++) {
                        const right = centerLeft + offset;
                        const left = centerLeft - offset;
                        if(right < count) order.push(right);
                        if(left >= 0) order.push(left);
                    }
                    return order;
                };
                const sortUtableSeatKeys = (keys) => {
                    const keySet = new Set(keys);
                    const leaderPositions = keys
                        .filter(key => key.startsWith('u-leader-'))
                        .map(key => Number.parseInt(key.split('-').pop(), 10))
                        .filter(pos => !Number.isNaN(pos));
                    const leaderCount = leaderPositions.length ? Math.max(...leaderPositions) + 1 : 0;
                    const ordered = getPodiumColumnOrder(leaderCount)
                        .map(pos => `u-leader-${pos}`)
                        .filter(key => keySet.has(key));
                    const sideMaps = {left: new Map(), right: new Map()};
                    let maxSidePosition = -1;
                    let maxSideColumn = -1;
                    keys.forEach(key => {
                        const match = key.match(/^u-(left|right)-(\d+)(?:-(\d+))?$/);
                        if(!match) return;
                        const side = match[1];
                        const column = match[3] === undefined ? 0 : Number.parseInt(match[2], 10);
                        const pos = Number.parseInt(match[3] === undefined ? match[2] : match[3], 10);
                        sideMaps[side].set(`${column}-${pos}`, key);
                        maxSideColumn = Math.max(maxSideColumn, column);
                        maxSidePosition = Math.max(maxSidePosition, pos);
                    });
                    for(let pos = 0; pos <= maxSidePosition; pos++) {
                        for(let column = 0; column <= maxSideColumn; column++) {
                            const leftKey = sideMaps.left.get(`${column}-${pos}`);
                            const rightKey = sideMaps.right.get(`${column}-${pos}`);
                            if(leftKey) ordered.push(leftKey);
                            if(rightKey) ordered.push(rightKey);
                        }
                    }
                    return ordered;
                };
                const getUtableLeaderOrderNumber = (physicalPosition) => {
                    const orderIndex = getPodiumColumnOrder(layout.utableConfig.leaderSeatCount).indexOf(physicalPosition);
                    return orderIndex >= 0 ? orderIndex + 1 : physicalPosition + 1;
                };
                const buildDefaultPodiumColumnLabels = (count) => {
                    const labels = Array(count).fill('');
                    getPodiumColumnOrder(count).forEach((columnIndex, orderIndex) => {
                        labels[columnIndex] = String(orderIndex + 1);
                    });
                    return labels;
                };

                const isFullMainAisleColumn = (columnIndex) => {
                    if(layout.main.rows <= 0) return false;
                    for(let row = 0; row < layout.main.rows; row++) {
                        if(seats[`m-${row}-${columnIndex}`]?.type !== 'aisle') return false;
                    }
                    return true;
                };
                const renumberMainColumnsAfterAisles = () => {
                    let nextSeatNumber = 1;
                    for(let column = 0; column < layout.main.cols; column++) {
                        if(isFullMainAisleColumn(column)) {
                            layout.main.colLabels[column] = '走廊';
                        } else {
                            layout.main.colLabels[column] = String(nextSeatNumber++);
                        }
                    }
                };

                onMounted(() => loadGlobalData());

                const initSeats = (overwriteLabels = true) => {
                    const makeEmptySeat = () => ({ type: 'normal', attendee: null, regionId: null, customLabel: null, displayConfig: null });
                    const evictedNames = [];
                    const seatGroupKey = (key) => {
                        const longtableMatch = key.match(/^lt-(top|bottom|left|right)-/);
                        if(longtableMatch) return `lt-${longtableMatch[1]}`;
                        if(key.startsWith('lt-head-left-')) return 'lt-head-left';
                        if(key.startsWith('lt-head-right-')) return 'lt-head-right';
                        if(key.startsWith('rt-')) return key.substring(0, key.lastIndexOf('-'));
                        return key.split('-')[0];
                    };
                    const removeRecordKey = (key) => {
                        for(let batchKey in regionAssignmentRecords) {
                            const sequence = regionAssignmentRecords[batchKey];
                            let index = sequence.indexOf(key);
                            while(index > -1) {
                                sequence.splice(index, 1);
                                index = sequence.indexOf(key);
                            }
                        }
                    };
                    const replaceRecordKey = (oldKey, newKey) => {
                        for(let batchKey in regionAssignmentRecords) {
                            const sequence = regionAssignmentRecords[batchKey];
                            const oldIndex = sequence.indexOf(oldKey);
                            if(oldIndex < 0) continue;
                            const newIndex = sequence.indexOf(newKey);
                            if(newIndex > -1) sequence.splice(oldIndex, 1);
                            else sequence[oldIndex] = newKey;
                        }
                    };
                    const syncSeatKeys = (desiredKeys, preserveOccupied = false) => {
                        desiredKeys.forEach(key => {
                            if (!seats[key]) seats[key] = makeEmptySeat();
                        });

                        if(preserveOccupied) {
                            const availableKeys = Array.from(desiredKeys).filter(key => seats[key].type === 'normal' && !seats[key].attendee);
                            const staleOccupiedKeys = Object.keys(seats).filter(key => !desiredKeys.has(key) && seats[key]?.attendee);
                            staleOccupiedKeys.forEach(oldKey => {
                                const preferredGroup = seatGroupKey(oldKey);
                                let destinationIndex = availableKeys.findIndex(key => seatGroupKey(key) === preferredGroup);
                                if(destinationIndex < 0) destinationIndex = 0;
                                if(availableKeys.length > 0 && destinationIndex >= 0) {
                                    const newKey = availableKeys.splice(destinationIndex, 1)[0];
                                    seats[newKey].attendee = seats[oldKey].attendee;
                                    replaceRecordKey(oldKey, newKey);
                                    seats[oldKey].attendee = null;
                                } else {
                                    evictedNames.push(seats[oldKey].attendee.name || '未命名人员');
                                    removeRecordKey(oldKey);
                                    seats[oldKey].attendee = null;
                                }
                            });
                        }

                        for (let key in seats) {
                            if (!desiredKeys.has(key)) delete seats[key];
                        }
                    };
                    const attendeePartyRole = (attendee) => {
                        const latest = attendees.value.find(item => String(item.id) === String(attendee?.id));
                        return normalizePartyRole(latest?.party_role ?? attendee?.party_role);
                    };
                    const collectRoleOccupants = (orderedKeys, role) => {
                        const result = [];
                        const visited = new Set();
                        orderedKeys.forEach(key => {
                            if(visited.has(key)) return;
                            visited.add(key);
                            const attendee = seats[key]?.attendee;
                            if(attendee && attendeePartyRole(attendee) === role) result.push({oldKey: key, attendee});
                        });
                        return result;
                    };
                    const sortRoundtableRoleKeys = (keys, role) => {
                        const groups = {};
                        keys.forEach(key => {
                            const prefix = key.substring(0, key.lastIndexOf('-'));
                            if(!groups[prefix]) groups[prefix] = [];
                            groups[prefix].push({key, pos: Number.parseInt(key.substring(key.lastIndexOf('-') + 1), 10)});
                        });
                        const ordered = [];
                        Object.keys(groups).sort((a, b) => a.localeCompare(b, undefined, {numeric: true})).forEach(prefix => {
                            const items = groups[prefix];
                            if(role === 'host') {
                                items.sort((a, b) => {
                                    if(a.pos === 0) return -1;
                                    if(b.pos === 0) return 1;
                                    return b.pos - a.pos;
                                });
                            } else {
                                items.sort((a, b) => {
                                    if(a.pos === 0) return 1;
                                    if(b.pos === 0) return -1;
                                    return a.pos - b.pos;
                                });
                            }
                            ordered.push(...items.map(item => item.key));
                        });
                        return ordered;
                    };
                    const clearReflowOccupants = (groups) => {
                        Object.values(groups).flat().forEach(item => {
                            if(seats[item.oldKey]) seats[item.oldKey].attendee = null;
                        });
                    };
                    const hasTopologyChanged = (oldKeys, desiredKeys) => {
                        if(oldKeys.length !== desiredKeys.size) return true;
                        return oldKeys.some(key => !desiredKeys.has(key));
                    };
                    const placeReflowOccupants = (groups, seatPlan) => {
                        const keyMapping = new Map();
                        const removedKeys = new Set();
                        ['host', 'guest'].forEach(role => {
                            const people = groups[role] || [];
                            const keys = seatPlan[role] || [];
                            people.forEach((item, index) => {
                                const newKey = keys[index];
                                if(newKey && seats[newKey] && seats[newKey].type === 'normal' && !seats[newKey].attendee) {
                                    seats[newKey].attendee = item.attendee;
                                    keyMapping.set(item.oldKey, newKey);
                                } else {
                                    evictedNames.push(item.attendee.name || '未命名人员');
                                    removedKeys.add(item.oldKey);
                                }
                            });
                        });
                        for(let batchKey in regionAssignmentRecords) {
                            const nextSequence = [];
                            regionAssignmentRecords[batchKey].forEach(key => {
                                if(removedKeys.has(key)) return;
                                const mappedKey = keyMapping.get(key) || key;
                                if(seats[mappedKey] && !nextSequence.includes(mappedKey)) nextSequence.push(mappedKey);
                            });
                            regionAssignmentRecords[batchKey] = nextSequence;
                        }
                    };
                    const placeOrderedOccupants = (people, orderedKeys) => {
                        const keyMapping = new Map();
                        const removedKeys = new Set();
                        people.forEach((item, index) => {
                            const newKey = orderedKeys[index];
                            if(newKey && seats[newKey] && seats[newKey].type === 'normal' && !seats[newKey].attendee) {
                                seats[newKey].attendee = item.attendee;
                                keyMapping.set(item.oldKey, newKey);
                            } else {
                                evictedNames.push(item.attendee.name || '未命名人员');
                                removedKeys.add(item.oldKey);
                            }
                        });
                        for(let batchKey in regionAssignmentRecords) {
                            const nextSequence = [];
                            regionAssignmentRecords[batchKey].forEach(key => {
                                if(removedKeys.has(key)) return;
                                const mappedKey = keyMapping.get(key) || key;
                                if(seats[mappedKey] && !nextSequence.includes(mappedKey)) nextSequence.push(mappedKey);
                            });
                            regionAssignmentRecords[batchKey] = nextSequence;
                        }
                    };
                    const clampInt = (value, min, max, fallback) => {
                        const parsed = Number.parseInt(value, 10);
                        if (Number.isNaN(parsed)) return fallback;
                        return Math.min(max, Math.max(min, parsed));
                    };

                    if (layout.layoutType === 'classroom') {
                        layout.podium.rows = clampInt(layout.podium.rows, 0, 20, 0);
                        layout.podium.cols = clampInt(layout.podium.cols, 0, 40, 0);
                        layout.main.rows = clampInt(layout.main.rows, 0, 40, 0);
                        layout.main.cols = clampInt(layout.main.cols, 0, 60, 0);
                        while(layout.main.rowLabels.length < layout.main.rows) layout.main.rowLabels.push(String(layout.main.rowLabels.length + 1));
                        while(layout.main.colLabels.length < layout.main.cols) layout.main.colLabels.push(String(layout.main.colLabels.length + 1));
                        layout.main.rowLabels.length = layout.main.rows; layout.main.colLabels.length = layout.main.cols;

                        if (overwriteLabels) {
                            let pRows = [];
                            for(let i=layout.podium.rows; i>0; i--) pRows.push(String(i));
                            layout.podium.rowLabels = pRows;
                        }
                        const usesLegacySequentialPodiumLabels = layout.podium.colLabels.length === layout.podium.cols
                            && layout.podium.colLabels.every((label, index) => String(label) === String(index + 1));
                        if(layout.podium.colLabels.length !== layout.podium.cols || (!overwriteLabels && usesLegacySequentialPodiumLabels)) {
                            layout.podium.colLabels = buildDefaultPodiumColumnLabels(layout.podium.cols);
                        } else {
                            while(layout.podium.colLabels.length < layout.podium.cols) layout.podium.colLabels.push(String(layout.podium.colLabels.length + 1));
                            layout.podium.colLabels.length = layout.podium.cols;
                        }

                        const desiredKeys = new Set();
                        for(let r=0; r<layout.podium.rows; r++) {
                            for(let c=0; c<layout.podium.cols; c++) {
                                desiredKeys.add(`p-${r}-${c}`);
                            }
                        }
                        for(let r=0; r<layout.main.rows; r++) {
                            for(let c=0; c<layout.main.cols; c++) {
                                desiredKeys.add(`m-${r}-${c}`);
                            }
                        }
                        syncSeatKeys(desiredKeys);
                        if(overwriteLabels && Array.from({length: layout.main.cols}, (_, column) => column).some(isFullMainAisleColumn)) {
                            renumberMainColumnsAfterAisles();
                        }
                    } else if (layout.layoutType === 'longtable') {
                        const cfg = layout.longtableConfig;
                        if(!['top', 'bottom'].includes(cfg.hostSide)) cfg.hostSide = 'bottom';
                        if(!['top', 'bottom'].includes(cfg.guestSide) || cfg.guestSide === cfg.hostSide) cfg.guestSide = cfg.hostSide === 'top' ? 'bottom' : 'top';
                        if(!['center_out', 'ltr', 'rtl', 'snake'].includes(cfg.partySeatRule)) cfg.partySeatRule = 'center_out';
                        const oldLongtableKeys = Object.keys(seats).filter(key => /^lt-(top|bottom|left|right)-/.test(key));
                        const oldExtensionKeys = oldLongtableKeys.filter(key => key.startsWith('lt-left-') || key.startsWith('lt-right-'));
                        const oldHostOrder = [
                            ...sortLongtableSeatKeys(oldLongtableKeys.filter(key => key.startsWith(`lt-${cfg.hostSide}-`)), cfg.partySeatRule),
                            ...sortLongtableExtensionSeatKeys(oldExtensionKeys, cfg.hostSide)
                        ];
                        const oldGuestOrder = [
                            ...sortLongtableSeatKeys(oldLongtableKeys.filter(key => key.startsWith(`lt-${cfg.guestSide}-`)), cfg.partySeatRule),
                            ...sortLongtableExtensionSeatKeys(oldExtensionKeys, cfg.guestSide)
                        ];
                        const reflowGroups = {
                            host: collectRoleOccupants(oldHostOrder, 'host'),
                            guest: collectRoleOccupants(oldGuestOrder, 'guest')
                        };
                        const oldPartySeatKeys = oldLongtableKeys;
                        const desiredKeys = new Set();
                        // 四边座位：lt-{side}-{row}-{col}
                        for(let side of ['top','bottom','left','right']) {
                            const sc = cfg.sides[side];
                            if(!sc) continue;
                            sc.rows = clampInt(sc.rows, 0, 5, 0);
                            sc.seatsPerRow = clampInt(sc.seatsPerRow, 0, 20, side === 'top' || side === 'bottom' ? 6 : 3);
                            sc.seatGap = clampInt(sc.seatGap, 0, 800, 8);
                            if(!sc.rows) continue;
                            for(let row = 0; row < sc.rows; row++) {
                                for(let col = 0; col < sc.seatsPerRow; col++) {
                                    desiredKeys.add(`lt-${side}-${row}-${col}`);
                                }
                            }
                        }
                        // 可选桌外主位
                        if (cfg.headSeats.enabled) {
                            cfg.headSeats.seatCount = clampInt(cfg.headSeats.seatCount, 1, 3, 1);
                            for(let pos = 0; pos < cfg.headSeats.seatCount; pos++) {
                                for(let end of ['head-left', 'head-right']) {
                                    desiredKeys.add(`lt-${end}-${pos}`);
                                }
                            }
                        }
                        const desiredPartySeatKeys = new Set(Array.from(desiredKeys).filter(key => /^lt-(top|bottom|left|right)-/.test(key)));
                        const shouldReflow = hasTopologyChanged(oldPartySeatKeys, desiredPartySeatKeys);
                        if(shouldReflow) clearReflowOccupants(reflowGroups);
                        syncSeatKeys(desiredKeys, true);
                        if(shouldReflow) {
                            const availablePartyKeys = Object.keys(seats).filter(key => /^lt-(top|bottom|left|right)-/.test(key) && seats[key].type === 'normal' && !seats[key].attendee);
                            placeReflowOccupants(
                                reflowGroups,
                                buildLongtablePartyPlan(availablePartyKeys, reflowGroups.host.length, reflowGroups.guest.length, cfg.partySeatRule, cfg.hostSide, cfg.guestSide)
                            );
                        }
                    } else if (layout.layoutType === 'roundtable') {
                        const cfg = layout.roundtableConfig;
                        const oldRoundtableKeys = Object.keys(seats).filter(key => key.startsWith('rt-'));
                        const reflowGroups = {
                            host: collectRoleOccupants(sortRoundtableRoleKeys(oldRoundtableKeys, 'host'), 'host'),
                            guest: collectRoleOccupants(sortRoundtableRoleKeys(oldRoundtableKeys, 'guest'), 'guest')
                        };
                        const desiredKeys = new Set();
                        cfg.tables = clampInt(cfg.tables, 1, 8, 1);
                        for(let t = 0; t < cfg.tables; t++) {
                            cfg.rings.forEach(ring => {
                                ring.seatCount = clampInt(ring.seatCount, 4, 30, 12);
                                ring.radius = clampInt(ring.radius, 100, 360, 150);
                                for(let pos = 0; pos < ring.seatCount; pos++) {
                                    desiredKeys.add(`rt-${t}-${ring.id}-${pos}`);
                                }
                            });
                        }
                        const shouldReflow = hasTopologyChanged(oldRoundtableKeys, desiredKeys);
                        if(shouldReflow) clearReflowOccupants(reflowGroups);
                        syncSeatKeys(desiredKeys, true);
                        if(shouldReflow) {
                            const availableRoundtableKeys = Object.keys(seats).filter(key => key.startsWith('rt-') && seats[key].type === 'normal' && !seats[key].attendee);
                            placeReflowOccupants(
                                reflowGroups,
                                buildRoundtablePartyPlan(availableRoundtableKeys, reflowGroups.host.length, reflowGroups.guest.length)
                            );
                        }
                    } else if (layout.layoutType === 'utable') {
                        const cfg = layout.utableConfig;
                        const oldUtableKeys = Object.keys(seats).filter(key => key.startsWith('u-'));
                        const logicalRank = new Map(sortUtableSeatKeys(oldUtableKeys).map((key, index) => [key, index]));
                        const attendeeRank = new Map(attendees.value.map((attendee, index) => [String(attendee.id), index]));
                        const occupants = oldUtableKeys
                            .map(key => ({oldKey: key, attendee: seats[key]?.attendee}))
                            .filter(item => !!item.attendee)
                            .sort((a, b) => {
                                const attendeeDiff = (attendeeRank.get(String(a.attendee.id)) ?? Number.MAX_SAFE_INTEGER)
                                    - (attendeeRank.get(String(b.attendee.id)) ?? Number.MAX_SAFE_INTEGER);
                                if(attendeeDiff !== 0) return attendeeDiff;
                                return (logicalRank.get(a.oldKey) ?? Number.MAX_SAFE_INTEGER) - (logicalRank.get(b.oldKey) ?? Number.MAX_SAFE_INTEGER);
                            });
                        cfg.leaderSeatCount = clampInt(cfg.leaderSeatCount, 1, 20, 5);
                        cfg.leftColumnCount = clampInt(cfg.leftColumnCount, 0, 10, 1);
                        cfg.leftSeatCount = clampInt(cfg.leftSeatCount, 0, 30, 6);
                        cfg.rightColumnCount = clampInt(cfg.rightColumnCount, 0, 10, 1);
                        cfg.rightSeatCount = clampInt(cfg.rightSeatCount, 0, 30, 6);
                        cfg.leaderGap = clampInt(cfg.leaderGap, 0, 100, 8);
                        cfg.sideGap = clampInt(cfg.sideGap, 0, 100, 8);
                        if(!['top', 'bottom'].includes(cfg.leaderSide)) cfg.leaderSide = 'top';
                        const desiredKeys = new Set();
                        for(let pos = 0; pos < cfg.leaderSeatCount; pos++) desiredKeys.add(`u-leader-${pos}`);
                        for(let column = 0; column < cfg.leftColumnCount; column++) {
                            for(let pos = 0; pos < cfg.leftSeatCount; pos++) desiredKeys.add(`u-left-${column}-${pos}`);
                        }
                        for(let column = 0; column < cfg.rightColumnCount; column++) {
                            for(let pos = 0; pos < cfg.rightSeatCount; pos++) desiredKeys.add(`u-right-${column}-${pos}`);
                        }
                        const shouldReflow = hasTopologyChanged(oldUtableKeys, desiredKeys);
                        if(shouldReflow) occupants.forEach(item => { if(seats[item.oldKey]) seats[item.oldKey].attendee = null; });
                        syncSeatKeys(desiredKeys);
                        if(shouldReflow) {
                            const availableKeys = sortUtableSeatKeys(Object.keys(seats).filter(key => key.startsWith('u-') && seats[key].type === 'normal' && !seats[key].attendee));
                            placeOrderedOccupants(occupants, availableKeys);
                        }
                    }
                    if(overwriteLabels !== false && evictedNames.length) {
                        const uniqueNames = [...new Set(evictedNames)];
                        alert(`座位数量减少后容量不足，已移除 ${uniqueNames.length} 人：${uniqueNames.join('、')}`);
                    }
                };
                let utableSeatSyncTimer = null;
                const scheduleUtableSeatSync = () => {
                    if(utableSeatSyncTimer) clearTimeout(utableSeatSyncTimer);
                    utableSeatSyncTimer = setTimeout(() => {
                        utableSeatSyncTimer = null;
                        initSeats();
                    }, 180);
                };

                const prevLayoutType = ref('classroom');

                const onLayoutTypeChange = (event) => {
                    if (Object.values(seats).some(s => s && s.attendee)) {
                        if (!confirm('切换布局类型将清除所有座位数据（不影响人员名单），确定继续吗？')) {
                            layout.layoutType = prevLayoutType.value;
                            return;
                        }
                    }
                    for(let k in seats) delete seats[k];
                    for(let k in regionAssignmentRecords) delete regionAssignmentRecords[k];
                    clickMode.value = 'none';
                    resetAssignmentForLayout();
                    initSeats();
                    autoSave();
                };

                const addRoundtableRing = () => {
                    const rings = layout.roundtableConfig.rings;
                    const maxRadius = rings.length > 0 ? Math.max(...rings.map(r => r.radius)) + 100 : 130;
                    rings.push({ id: `ring-${Date.now()}`, label: `第${rings.length + 1}圈`, seatCount: 16, radius: maxRadius });
                    initSeats();
                };
                const removeRoundtableRing = (idx) => {
                    layout.roundtableConfig.rings.splice(idx, 1);
                    initSeats();
                };

                const handleLabelChange = (type, index, zone) => {
                    const obj = zone === 'p' ? layout.podium : layout.main;
                    const list = type === 'row' ? obj.rowLabels : obj.colLabels;
                    if (type === 'col') return; // 列号不顺延
                    const num = parseInt(list[index]);
                    if (!isNaN(num)) {
                        let currentNum = num;
                        for (let i = index + 1; i < list.length; i++) {
                            const nextVal = list[i];
                            if (isNaN(parseInt(nextVal)) && nextVal.trim() !== '') continue;
                            currentNum++;
                            list[i] = String(currentNum);
                        }
                    }
                };

                const syncPodiumLabels = () => {
                    saveHistory();
                    let podiumAtts = [];
                    for(let r=0; r<layout.podium.rows; r++){
                        for(let c=0; c<layout.podium.cols; c++){
                            const s = seats[`p-${r}-${c}`];
                            if(s && s.attendee) {
                                podiumAtts.push(s.attendee);
                                s.attendee = null;
                            }
                        }
                    }
                    if(podiumAtts.length === 0) return;

                    let targetCoords = [];
                    for(let r=0; r<layout.podium.rows; r++){
                        for(let c=0; c<layout.podium.cols; c++){
                            const s = seats[`p-${r}-${c}`];
                            if(s && s.type === 'normal') targetCoords.push({r, c, z:'p'});
                        }
                    }
                    targetCoords.sort((a,b) => {
                        let valA_r = parseFloat(layout.podium.rowLabels[a.r]); if(isNaN(valA_r)) valA_r = a.r * 1000;
                        let valB_r = parseFloat(layout.podium.rowLabels[b.r]); if(isNaN(valB_r)) valB_r = b.r * 1000;
                        if(valA_r !== valB_r) return valA_r - valB_r;

                        let valA_c = parseFloat(layout.podium.colLabels[a.c]); if(isNaN(valA_c)) valA_c = a.c * 1000;
                        let valB_c = parseFloat(layout.podium.colLabels[b.c]); if(isNaN(valB_c)) valB_c = b.c * 1000;
                        return valA_c - valB_c;
                    });

                    podiumAtts.sort((a,b) => {
                        return attendees.value.findIndex(x=>x.id===a.id) - attendees.value.findIndex(x=>x.id===b.id);
                    });

                    let newSeq = [];
                    for(let i=0; i<targetCoords.length && i<podiumAtts.length; i++) {
                        const pos = targetCoords[i];
                        const k = `p-${pos.r}-${pos.c}`;
                        seats[k].attendee = podiumAtts[i];
                        newSeq.push(k);
                    }
                    regionAssignmentRecords['podium_auto_sync'] = newSeq;
                };

                const formatColLabel = (label) => {
                    if (!label) return '';
                    if (label.endsWith('座') || label.endsWith('号') || isNaN(parseInt(label))) return label;
                    return label + '座';
                };

                const getSeatColLabelUI = (zone, r, c) => {
                    const key = makeSeatKey(zone, r, c);
                    const s = seats[key];
                    if(s && s.customLabel !== undefined && s.customLabel !== null) return s.customLabel;
                    if (zone === 'p' || zone === 'm') {
                        const obj = zone === 'p' ? layout.podium : layout.main;
                        return formatColLabel(obj.colLabels[c]);
                    }
                    if (zone.startsWith('lt-head-')) return formatColLabel(String(r + 1));
                    if (zone.startsWith('lt-')) return formatColLabel(String(c + 1));
                    return formatColLabel(String(r + 1));
                };

                const setCustomLabel = (key, val) => {
                    saveHistory();
                    if(!val || val.trim() === '') seats[key].customLabel = null;
                    else seats[key].customLabel = val;
                    autoSave();
                };

                const isDuplicateSeat = (zone, r, c) => {
                    const s = seats[makeSeatKey(zone, r, c)];
                    if (!s || s.type !== 'normal') return false;
                    const myLabel = getSeatColLabelUI(zone, r, c);
                    // 新布局类型不使用重复检测（每行每边座位都不同）
                    if (zone !== 'p' && zone !== 'm') return false;
                    const cols = zone === 'p' ? layout.podium.cols : layout.main.cols;
                    let count = 0;
                    for(let i=0; i<cols; i++) {
                        const other = seats[`${zone}-${r}-${i}`];
                        if(other && other.type === 'normal') {
                            if (getSeatColLabelUI(zone, r, i) === myLabel) {
                                count++;
                                if (count > 1) return true;
                            }
                        }
                    }
                    return false;
                };

                const clearCustomSequence = (regId) => {
                    if(!regId) return;
                    saveHistory();
                    if(!layout.customSequences) layout.customSequences = {};
                    layout.customSequences[regId] = [];
                    autoSave();
                };

                const modifyAxis = (action) => {
                    saveHistory();
                    const z = axisMenu.zone;
                    const obj = z === 'm' ? layout.main : layout.podium;
                    const t = axisMenu.type;
                    const idx = axisMenu.index;
                    const hadMainAisleColumn = z === 'm' && t === 'col'
                        && Array.from({length: layout.main.cols}, (_, column) => column).some(isFullMainAisleColumn);

                    if (action === 'delete') {
                        if (t === 'row') {
                            if (obj.rows <= 1) return alert('最后一行无法删除');
                            obj.rows--; obj.rowLabels.splice(idx, 1);
                            for (let r = idx; r < obj.rows; r++) {
                                for (let c = 0; c < obj.cols; c++) seats[`${z}-${r}-${c}`] = seats[`${z}-${r+1}-${c}`];
                            }
                            for (let c = 0; c < obj.cols; c++) delete seats[`${z}-${obj.rows}-${c}`];
                        } else {
                            if (obj.cols <= 1) return alert('最后一列无法删除');
                            obj.cols--; obj.colLabels.splice(idx, 1);
                            for (let c = idx; c < obj.cols; c++) {
                                for (let r = 0; r < obj.rows; r++) seats[`${z}-${r}-${c}`] = seats[`${z}-${r}-${c+1}`];
                            }
                            for (let r = 0; r < obj.rows; r++) delete seats[`${z}-${r}-${obj.cols}`];
                        }
                    } else {
                        const newIdx = action === 'insert_before' ? idx : idx + 1;
                        if (t === 'row') {
                            obj.rows++; obj.rowLabels.splice(newIdx, 0, '新排');
                            for (let r = obj.rows - 1; r > newIdx; r--) {
                                for (let c = 0; c < obj.cols; c++) seats[`${z}-${r}-${c}`] = seats[`${z}-${r-1}-${c}`] || {type:'normal', attendee:null, regionId:null, customLabel:null, displayConfig:null};
                            }
                            for (let c = 0; c < obj.cols; c++) seats[`${z}-${newIdx}-${c}`] = {type:'normal', attendee:null, regionId:null, customLabel:null, displayConfig:null};
                        } else {
                            obj.cols++; obj.colLabels.splice(newIdx, 0, '新列');
                            for (let c = obj.cols - 1; c > newIdx; c--) {
                                for (let r = 0; r < obj.rows; r++) seats[`${z}-${r}-${c}`] = seats[`${z}-${r}-${c-1}`] || {type:'normal', attendee:null, regionId:null, customLabel:null, displayConfig:null};
                            }
                            for (let r = 0; r < obj.rows; r++) seats[`${z}-${r}-${newIdx}`] = {type:'normal', attendee:null, regionId:null, customLabel:null, displayConfig:null};
                        }
                    }
                    if(z === 'm' && t === 'col' && (hadMainAisleColumn || Array.from({length: layout.main.cols}, (_, column) => column).some(isFullMainAisleColumn))) {
                        renumberMainColumnsAfterAisles();
                    }
                    closeContextMenus();
                    autoSave();
                };

                const setAxisStatus = (status) => {
                    saveHistory();
                    const obj = axisMenu.zone === 'p' ? layout.podium : layout.main;
                    if(status === 'aisle') {
                        if(axisMenu.type === 'row') obj.rowLabels[axisMenu.index] = '走廊';
                        else obj.colLabels[axisMenu.index] = '走廊';
                    }
                    if(axisMenu.type === 'row') {
                        for(let c=0; c<obj.cols; c++) {
                            const k = `${axisMenu.zone}-${axisMenu.index}-${c}`;
                            if(seats[k]) { seats[k].type = status; seats[k].attendee = null; seats[k].regionId = null; seats[k].customLabel = null; seats[k].displayConfig = null; }
                        }
                    } else {
                        for(let r=0; r<obj.rows; r++) {
                            const k = `${axisMenu.zone}-${r}-${axisMenu.index}`;
                            if(seats[k]) { seats[k].type = status; seats[k].attendee = null; seats[k].regionId = null; seats[k].customLabel = null; seats[k].displayConfig = null; }
                        }
                    }
                    if(axisMenu.zone === 'm' && axisMenu.type === 'col') renumberMainColumnsAfterAisles();
                    closeContextMenus();
                };

                const persistAttendeeOrder = async (nextOrder) => {
                    attendees.value.splice(0, attendees.value.length, ...nextOrder);
                    await fetch(`/api/meetings/${currentMeetingId.value}/attendees/reorder`, {
                        method: 'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({ids: nextOrder.map(a => a.id)})
                    });
                    syncSeatingOrder();
                };
                const mergePartyScopeOrder = (att, orderedScope) => {
                    if (!isPartyLayout.value) return orderedScope;
                    const targetRole = normalizePartyRole(att.party_role);
                    const result = [];
                    ['host', 'guest', 'unassigned'].forEach(role => {
                        if (role === targetRole) result.push(...orderedScope);
                        else result.push(...partySortedAttendees.value.filter(item => normalizePartyRole(item.party_role) === role));
                    });
                    return result;
                };
                const onDragStart = (e, attId) => { e.dataTransfer.setData('sourceAttId', String(attId)); };
                const onDragOver = (e, attId) => { dragTargetIndex.value = attId; };
                const onDragLeave = (e) => { dragTargetIndex.value = -1; };
                const onDrop = async (e, targetAttId) => {
                    dragTargetIndex.value = -1;
                    const sourceAttId = parseInt(e.dataTransfer.getData('sourceAttId'));
                    if(isNaN(sourceAttId) || sourceAttId === targetAttId) return;

                    const sourceAtt = attendees.value.find(att => att.id === sourceAttId);
                    const targetAtt = attendees.value.find(att => att.id === targetAttId);
                    if(!sourceAtt || !targetAtt) return;
                    if(isPartyLayout.value && normalizePartyRole(sourceAtt.party_role) !== normalizePartyRole(targetAtt.party_role)) {
                        return alert('主方、宾方和未分类人员需要在各自组内分别排序。');
                    }

                    const scope = [...getAttendeeScope(sourceAtt)];
                    const sourceIdx = scope.findIndex(att => att.id === sourceAttId);
                    const targetIdx = scope.findIndex(att => att.id === targetAttId);
                    if(sourceIdx < 0 || targetIdx < 0) return;

                    saveHistory();
                    const item = scope.splice(sourceIdx, 1)[0];
                    scope.splice(targetIdx, 0, item);
                    await persistAttendeeOrder(mergePartyScopeOrder(sourceAtt, scope));
                };

                const moveAttendee = async (attId, direction) => {
                    const att = attendees.value.find(a => a.id === attId);
                    if (!att) return;
                    const scope = [...getAttendeeScope(att)];
                    const idx = scope.findIndex(a => a.id === attId);
                    if (idx < 0) return;

                    const newIdx = idx + direction;
                    if (newIdx < 0 || newIdx >= scope.length) return;

                    saveHistory();
                    const item = scope.splice(idx, 1)[0];
                    scope.splice(newIdx, 0, item);
                    await persistAttendeeOrder(mergePartyScopeOrder(att, scope));
                };

                const setAttendeePosition = async (attId, newPos) => {
                    const targetPos = parseInt(newPos);
                    if (isNaN(targetPos) || targetPos <= 0) return;

                    const att = attendees.value.find(a => a.id === attId);
                    if (!att) return;
                    const scope = [...getAttendeeScope(att)];
                    const idx = scope.findIndex(a => a.id === attId);
                    if (idx < 0) return;

                    const insertIdx = Math.min(targetPos - 1, scope.length - 1);
                    if (insertIdx === idx) return;

                    saveHistory();
                    const item = scope.splice(idx, 1)[0];
                    scope.splice(insertIdx, 0, item);
                    await persistAttendeeOrder(mergePartyScopeOrder(att, scope));
                };

                const syncSeatingOrder = () => {
                    for(let batchKey in regionAssignmentRecords) {
                        const seq = regionAssignmentRecords[batchKey];
                        let seatedAtts = [];
                        let occupiedIndices = [];

                        seq.forEach((key, idx) => {
                            if(seats[key] && seats[key].attendee) {
                                seatedAtts.push(seats[key].attendee);
                                occupiedIndices.push(idx);
                            }
                        });

                        if(seatedAtts.length === 0) continue;

                        seatedAtts.sort((a,b) => {
                            const idA = attendees.value.findIndex(x=>x.id===a.id);
                            const idB = attendees.value.findIndex(x=>x.id===b.id);
                            return idA - idB;
                        });

                        for (let i = 0; i < occupiedIndices.length; i++) {
                            const key = seq[occupiedIndices[i]];
                            seats[key].attendee = seatedAtts[i];
                        }
                    }
                };

                const applyRangeSelection = () => {
                    if (!assignRangeInput.value.trim()) {
                        assignSelectedAtts.value = [];
                        return;
                    }
                    const parts = assignRangeInput.value.split(',');
                    const toSelectIds = new Set();
                    parts.forEach(part => {
                        part = part.trim();
                        if (part.includes('-')) {
                            const [startStr, endStr] = part.split('-');
                            const start = parseInt(startStr);
                            const end = parseInt(endStr);
                            if (!isNaN(start) && !isNaN(end)) {
                                for (let i = Math.min(start, end); i <= Math.max(start, end); i++) {
                                    toSelectIds.add(i);
                                }
                            }
                        } else {
                            const num = parseInt(part);
                            if (!isNaN(num)) toSelectIds.add(num);
                        }
                    });

                    const nextSelection = [];
                    unassignedFilteredAttendees.value.forEach(att => {
                        const gIdx = getAttendeeOrder(att);
                        if (toSelectIds.has(gIdx)) {
                            nextSelection.push(att);
                        }
                    });
                    assignSelectedAtts.value = nextSelection;
                };

                const uploadFile = async (e) => {
                    const file = e.target.files[0];
                    if(!file) return;
                    const formData = new FormData(); formData.append('file', file);
                    try {
                        const res = await fetch(`/api/meetings/${currentMeetingId.value}/upload`, { method: 'POST', body: formData });
                        const data = await res.json();
                        alert(data.message || data.error);
                        fetchAttendees();
                    } catch(err) { alert('上传失败'); }
                    e.target.value = '';
                };

                const toggleSelectAllFiltered = (e) => {
                    if(e.target.checked) {
                        const ids = filteredAttendees.value.map(a => a.id);
                        selectedAttIds.value = [...new Set([...selectedAttIds.value, ...ids])];
                    } else {
                        const idsToRemove = filteredAttendees.value.map(a => a.id);
                        selectedAttIds.value = selectedAttIds.value.filter(id => !idsToRemove.includes(id));
                    }
                };

                const isAllSelected = computed(() => {
                    if(filteredAttendees.value.length === 0) return false;
                    return filteredAttendees.value.every(a => selectedAttIds.value.includes(a.id));
                });

                const bulkDeleteAttendees = async () => {
                    if(!confirm('确定删除选中的人员吗？若已排座，该座位将会空出并由后方人员补位！')) return;
                    saveHistory();

                    selectedAttIds.value.forEach(id => {
                        let foundKey = null;
                        for(let k in seats) { if(seats[k].attendee && seats[k].attendee.id === id) { foundKey = k; break; } }
                        if(foundKey) {
                            // 对于新布局类型，key 有 3 段以上；对于课桌式，key 是 zone-r-c
                            const parts = foundKey.split('-');
                            if (parts.length <= 3) {
                                removeAndShift(parts[0], parseInt(parts[1]), parseInt(parts[2]), false);
                            } else {
                                // 重建 zone 前缀和 pos
                                const zone = parts.slice(0, parts.length-1).join('-');
                                removeAndShift(zone, parseInt(parts[parts.length-1]), 0, false);
                            }
                        }
                    });

                    await fetch(`/api/meetings/${currentMeetingId.value}/attendees/bulk_delete`, {
                        method: 'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({ids: selectedAttIds.value})
                    });
                    selectedAttIds.value = [];
                    fetchAttendees();
                };

                const setPartyRoleForIds = async (ids, partyRole) => {
                    if(!ids.length) return;
                    const normalizedRole = normalizePartyRole(partyRole);
                    const idSet = new Set(ids);
                    const originalOrder = new Map(attendees.value.map((att, index) => [att.id, index]));
                    attendees.value.forEach(att => {
                        if(idSet.has(att.id)) att.party_role = normalizedRole;
                    });
                    Object.values(seats).forEach(seat => {
                        if(seat?.attendee && idSet.has(seat.attendee.id)) seat.attendee.party_role = normalizedRole;
                    });
                    await fetch(`/api/meetings/${currentMeetingId.value}/attendees/bulk_role`, {
                        method: 'POST', headers:{'Content-Type':'application/json'},
                        body: JSON.stringify({ids, party_role: normalizedRole})
                    });
                    const ordered = [...attendees.value].sort((a, b) => {
                        const roleDiff = partyRoleRank[normalizePartyRole(a.party_role)] - partyRoleRank[normalizePartyRole(b.party_role)];
                        if(roleDiff !== 0) return roleDiff;
                        return originalOrder.get(a.id) - originalOrder.get(b.id);
                    });
                    await persistAttendeeOrder(ordered);
                };
                const setSelectedPartyRole = async (partyRole) => {
                    await setPartyRoleForIds([...selectedAttIds.value], partyRole);
                    selectedAttIds.value = [];
                };
                const setAttendeePartyRole = async (attId, partyRole) => {
                    await setPartyRoleForIds([attId], partyRole);
                };

                const openAddModal = () => {
                    const extraFields = {};
                    extraKeys.value.forEach(key => extraFields[key] = '');
                    Object.assign(modalForm, {id: null, name: '', unit: '', category: '', phone: '', party_role: 'unassigned', extra: extraFields});
                    showModal.value = true;
                };
                const editAttendee = (att) => {
                    const attExtra = att.extra || {};
                    const extraFields = {};
                    extraKeys.value.forEach(key => extraFields[key] = attExtra[key] || '');
                    Object.assign(modalForm, {id: att.id, name: att.name, unit: att.unit, category: att.category, phone: att.phone, party_role: normalizePartyRole(att.party_role), extra: extraFields});
                    showModal.value = true;
                };
                const saveModalForm = async () => {
                    if(!modalForm.name) return alert('姓名必填');
                    const method = modalForm.id ? 'PUT' : 'POST';
                    const editedId = modalForm.id;
                    const formData = {
                        id: modalForm.id,
                        name: modalForm.name,
                        unit: modalForm.unit,
                        category: modalForm.category,
                        phone: modalForm.phone,
                        party_role: modalForm.party_role,
                        extra: modalForm.extra
                    };
                    await fetch(`/api/meetings/${currentMeetingId.value}/attendees`, { method, headers: {'Content-Type': 'application/json'}, body: JSON.stringify(formData) });
                    if(editedId) {
                        Object.values(seats).forEach(seat => {
                            if(seat?.attendee?.id === editedId) seat.attendee.party_role = normalizePartyRole(modalForm.party_role);
                        });
                    }
                    showModal.value = false;
                    await fetchAttendees();
                };

                const addRegionType = () => { layout.regionTypes.push({id: Date.now().toString(), name: '新区域', color: '#e5e7eb'}); };
                const removeRegionType = (id) => {
                    layout.regionTypes = layout.regionTypes.filter(r => r.id !== id);
                    if(activeRegionId.value === id && layout.regionTypes.length>0) activeRegionId.value = layout.regionTypes[0].id;
                    for(let key in seats) if(seats[key].regionId === id) seats[key].regionId = null;

                    const targetReg = `reg_${id}`;
                    for(let k in seats) {
                        if(seats[k].regionId === id) seats[k].regionId = null;
                    }
                };
                const getSeatRegionColor = (zone, r, c) => {
                    if(zone !== 'm') return null;
                    const s = seats[`m-${r}-${c}`];
                    if(s && s.regionId) {
                        const reg = layout.regionTypes.find(rg => rg.id === s.regionId);
                        return reg ? reg.color : null;
                    }
                    return null;
                };

                const makeSeatKey = (zone, r, c) => {
                    if (zone === 'p' || zone === 'm') return `${zone}-${r}-${c}`;
                    // 长会议桌桌外主位: lt-head-{left/right}-{pos}
                    if (zone.startsWith('lt-head-')) return `${zone}-${r}`;
                    // 长会议桌四边: lt-{side}-{row}-{col}
                    if (zone.startsWith('lt-')) return `${zone}-${r}-${c}`;
                    // 圆桌式: rt-{table}-{ring}-{pos}（zone 已含桌号和圈 id）
                    return `${zone}-${r}`;
                };

                const seatClass = (zone, r, c) => {
                    const s = seats[makeSeatKey(zone, r, c)];
                    if(!s) return '';
                    if(s.type === 'aisle') return 'seat-aisle';
                    if(s.type === 'disabled') return 'seat-disabled';
                    if(s.attendee) return 'seat-occupied';
                    return '';
                };

                const seatStyle = (zone, r, c) => {
                    const s = seats[makeSeatKey(zone, r, c)];
                    if(!s || s.type !== 'normal' || !s.attendee) return {};
                    if (zone === 'p') {
                        return { backgroundColor: layout.podium.color, borderColor: layout.podium.color, color: layout.podium.textColor };
                    } else if (zone === 'm') {
                        return { backgroundColor: layout.main.color, borderColor: layout.main.color, color: layout.main.textColor };
                    }
                    // 新布局类型：使用默认颜色
                    return { backgroundColor: '#6366f1', borderColor: '#6366f1', color: '#ffffff' };
                };

                const seatStyleLongtableRow = (side, row, col) => {
                    const key = `lt-${side}-${row}-${col}`;
                    const s = seats[key];
                    if(!s || s.type !== 'normal' || !s.attendee) return {};
                    return { backgroundColor: '#6366f1', borderColor: '#4f46e5', color: '#ffffff' };
                };
                const seatStyleLongtable = (sideId, pos) => {
                    const key = `lt-${sideId}-${pos}`;
                    const s = seats[key];
                    if(!s || s.type !== 'normal' || !s.attendee) return {};
                    return { backgroundColor: '#6366f1', borderColor: '#4f46e5', color: '#ffffff' };
                };

                const roundtableSeatStyle = (tableIdx, ring, pos) => {
                    const sceneW = Math.max(...layout.roundtableConfig.rings.map(r => r.radius)) * 2 + 120;
                    const cx = sceneW / 2, cy = sceneW / 2;
                    const step = 2 * Math.PI / ring.seatCount;
                    // 双主位时让 1、2 号座关于 12 点方向对称；座位键不变，因此增减人数不会打乱原座次。
                    const principalOffset = layout.roundtableConfig.dualPrincipal ? step / 2 : 0;
                    const angle = step * pos - Math.PI / 2 - principalOffset;
                    const left = cx + ring.radius * Math.cos(angle);
                    const top = cy + ring.radius * Math.sin(angle);
                    const baseStyle = {
                        position: 'absolute',
                        left: left + 'px',
                        top: top + 'px',
                        transform: 'translate(-50%, -50%)'
                    };
                    const key = `rt-${tableIdx}-${ring.id}-${pos}`;
                    const s = seats[key];
                    if(!s || s.type !== 'normal' || !s.attendee) return baseStyle;
                    return Object.assign({}, baseStyle, {
                        backgroundColor: '#6366f1',
                        borderColor: '#4f46e5',
                        color: '#ffffff'
                    });
                };

                const closeContextMenus = () => { seatMenu.show = false; axisMenu.show = false; };

                const openSeatDisplayEdit = (zone, r, c) => {
                    const key = makeSeatKey(zone, r, c);
                    const existing = seats[key]?.displayConfig;
                    seatDisplayEdit.zone = zone;
                    seatDisplayEdit.r = r;
                    seatDisplayEdit.c = c;
                    seatDisplayEdit.mainField = existing ? existing.mainField : layout.seatDisplayConfig.mainField;
                    seatDisplayEdit.subField = existing ? existing.subField : layout.seatDisplayConfig.subField;
                    seatDisplayEdit.show = true;
                };

                const saveSeatDisplayConfig = () => {
                    saveHistory();
                    const key = makeSeatKey(seatDisplayEdit.zone, seatDisplayEdit.r, seatDisplayEdit.c);
                    if (!seats[key]) seats[key] = { type: 'normal', attendee: null, regionId: null, customLabel: null, displayConfig: null };
                    seats[key].displayConfig = {
                        mainField: seatDisplayEdit.mainField,
                        subField: seatDisplayEdit.subField || null
                    };
                    seatDisplayEdit.show = false;
                    seatMenu.show = false;
                    autoSave();
                };

                const resetSeatDisplayConfig = (zone, r, c) => {
                    saveHistory();
                    const key = makeSeatKey(zone, r, c);
                    if (seats[key]) seats[key].displayConfig = null;
                    seatMenu.show = false;
                    autoSave();
                };
                const handleSeatRightClick = (e, zone, r, c) => {
                    const key = makeSeatKey(zone, r, c);
                    if(!seats[key] || seats[key].type !== 'normal') return;
                    seatMenuSearch.value = ""; seatMenu.zone = zone; seatMenu.r = r; seatMenu.c = c;
                    seatMenu.hasAttendee = !!seats[key].attendee;
                    seatMenu.x = Math.min(e.clientX, document.body.clientWidth - 220);
                    seatMenu.y = Math.min(e.clientY, document.body.clientHeight - 300);
                    seatMenu.show = true; axisMenu.show = false;
                };
                const handleAxisRightClick = (e, type, index, zone) => {
                    axisMenu.type = type; axisMenu.index = index; axisMenu.zone = zone;
                    axisMenu.x = Math.min(e.clientX, document.body.clientWidth - 200);
                    axisMenu.y = Math.min(e.clientY, document.body.clientHeight - 200);
                    axisMenu.show = true; seatMenu.show = false;
                };

                const forceAssign = (zone, r, c, att) => { saveHistory(); seats[makeSeatKey(zone, r, c)].attendee = att; closeContextMenus(); };
                const clearSeat = (zone, r, c) => { saveHistory(); seats[makeSeatKey(zone, r, c)].attendee = null; closeContextMenus(); };

                const removeAndShift = (zone, r, c, pushHistory=true) => {
                    const key = makeSeatKey(zone, r, c);
                    let foundBatch = null;
                    for(let bKey in regionAssignmentRecords) {
                        if(regionAssignmentRecords[bKey].includes(key)) { foundBatch = bKey; break; }
                    }
                    if(pushHistory) saveHistory();

                    if(!foundBatch) {
                        seats[key].attendee = null;
                    } else {
                        const seq = regionAssignmentRecords[foundBatch];
                        const dropIndex = seq.indexOf(key);
                        seats[key].attendee = null;
                        for(let i = dropIndex; i < seq.length - 1; i++) {
                            seats[seq[i]].attendee = seats[seq[i+1]].attendee;
                        }
                        seats[seq[seq.length - 1]].attendee = null;
                    }

                    if(pushHistory) {
                        closeContextMenus();
                        autoSave();
                    }
                };

                const handleSeatClick = (zone, r, c) => {
                    const key = makeSeatKey(zone, r, c);
                    if(!seats[key]) return;
                    if(clickMode.value === 'toggle_aisle') {
                        const wasFullMainAisleColumn = zone === 'm' && isFullMainAisleColumn(c);
                        saveHistory();
                        seats[key].type = seats[key].type==='aisle'?'normal':'aisle';
                        seats[key].attendee=null; seats[key].customLabel=null; seats[key].displayConfig=null;
                        if(zone === 'm' && (wasFullMainAisleColumn || isFullMainAisleColumn(c))) renumberMainColumnsAfterAisles();
                    }
                    else if(clickMode.value === 'toggle_disabled') { saveHistory(); seats[key].type = seats[key].type==='disabled'?'normal':'disabled'; seats[key].attendee=null; seats[key].customLabel=null; seats[key].displayConfig=null; }
                    else if(clickMode.value === 'draw_region') {
                        if(zone === 'p' || seats[key].type !== 'normal') return;
                        saveHistory(); seats[key].regionId = seats[key].regionId === activeRegionId.value ? null : activeRegionId.value;
                    }
                    else if(clickMode.value === 'draw_region_batch') {
                        if(zone === 'p' || seats[key].type !== 'normal') return;
                        if(!batchSelectState.hasStart) {
                            batchSelectState.hasStart = true;
                            batchSelectState.startZone = zone;
                            batchSelectState.startR = r;
                            batchSelectState.startC = c;
                        } else {
                            const startZone = batchSelectState.startZone;
                            const startR = batchSelectState.startR;
                            const startC = batchSelectState.startC;
                            if(startZone !== zone) {
                                batchSelectState.hasStart = false;
                                batchSelectState.startZone = '';
                                batchSelectState.startR = -1;
                                batchSelectState.startC = -1;
                                return;
                            }
                            saveHistory();
                            if(startR === r) {
                                const minC = Math.min(startC, c);
                                const maxC = Math.max(startC, c);
                                for(let col = minC; col <= maxC; col++) {
                                    const k = `${zone}-${r}-${col}`;
                                    if(seats[k] && seats[k].type === 'normal') {
                                        seats[k].regionId = activeRegionId.value;
                                    }
                                }
                            } else if(startC === c) {
                                const minR = Math.min(startR, r);
                                const maxR = Math.max(startR, r);
                                for(let row = minR; row <= maxR; row++) {
                                    const k = `${zone}-${row}-${c}`;
                                    if(seats[k] && seats[k].type === 'normal') {
                                        seats[k].regionId = activeRegionId.value;
                                    }
                                }
                            }
                            batchSelectState.hasStart = false;
                            batchSelectState.startZone = '';
                            batchSelectState.startR = -1;
                            batchSelectState.startC = -1;
                        }
                    }
                    else if(clickMode.value === 'define_sequence') {
                        if(zone === 'p' || seats[key].type !== 'normal') return;
                        const rId = seats[key].regionId;
                        if(!rId) return;
                        saveHistory();

                        if(activeRegionId.value !== rId) activeRegionId.value = rId;
                        if(!layout.customSequences) layout.customSequences = {};
                        if(!layout.customSequences[rId]) layout.customSequences[rId] = [];

                        const seq = layout.customSequences[rId];
                        const idx = seq.indexOf(key);
                        if(idx > -1) seq.splice(idx, 1);
                        else seq.push(key);
                    }
                };

                const clearTargetZone = () => {
                    if(!confirm('确定清除选中区域内所有人的座位吗？')) return;
                    saveHistory();
                    const z = assignConfig.targetZone;

                    let keysToClear = [];
                    if(layout.layoutType !== 'classroom') {
                        for(let k in seats) {
                            if(!seats[k].attendee) continue;
                            if(z === 'party_auto' && layout.layoutType === 'longtable') {
                                if(/^lt-(top|bottom|left|right)-/.test(k)) keysToClear.push(k);
                            }
                            else if(z === 'party_auto' && layout.layoutType === 'roundtable' && k.startsWith('rt-')) keysToClear.push(k);
                            else if(z === 'u_auto' && layout.layoutType === 'utable' && k.startsWith('u-')) keysToClear.push(k);
                            else if(z==='all') keysToClear.push(k);
                            else if(k.startsWith(z + '-')) keysToClear.push(k);
                        }
                    } else {
                        for(let k in seats) {
                            const p = k.split('-');
                            if(z==='podium' && p[0]==='p') keysToClear.push(k);
                            else if(z==='all_m' && p[0]==='m' && !seats[k].regionId) keysToClear.push(k);
                            else if(z.startsWith('reg_') && seats[k].regionId === z.split('_')[1]) keysToClear.push(k);
                        }
                    }

                    keysToClear.forEach(k => {
                        seats[k].attendee = null;
                        for(let bKey in regionAssignmentRecords) {
                            const idx = regionAssignmentRecords[bKey].indexOf(k);
                            if(idx > -1) regionAssignmentRecords[bKey].splice(idx, 1);
                        }
                    });
                };

                const sortLongtableSeatKeys = (keys, rule = 'center_out') => {
                    const rowGroups = {};
                    keys.forEach(key => {
                        const parts = key.split('-');
                        const rowPrefix = parts.slice(0, -1).join('-');
                        if(!rowGroups[rowPrefix]) rowGroups[rowPrefix] = [];
                        rowGroups[rowPrefix].push({key, pos: parseInt(parts[parts.length - 1])});
                    });
                    const ordered = [];
                    Object.keys(rowGroups).sort((a, b) => {
                        const ar = parseInt(a.split('-').pop()) || 0;
                        const br = parseInt(b.split('-').pop()) || 0;
                        return ar - br;
                    }).forEach(prefix => {
                        const items = rowGroups[prefix];
                        items.sort((a, b) => a.pos - b.pos);
                        const rowIdx = parseInt(prefix.split('-').pop()) || 0;
                        const physicalCenter = (items.length - 1) / 2;
                        if(rule === 'center_out') {
                            items.sort((a, b) => {
                                const distanceDiff = Math.abs(a.pos - physicalCenter) - Math.abs(b.pos - physicalCenter);
                                if(Math.abs(distanceDiff) > 0.01) return distanceDiff;
                                return rowIdx % 2 === 0 ? b.pos - a.pos : a.pos - b.pos;
                            });
                        } else if(rule === 'rtl') {
                            items.sort((a, b) => b.pos - a.pos);
                        } else if(rule === 'snake' && rowIdx % 2 === 1) {
                            items.sort((a, b) => b.pos - a.pos);
                        }
                        ordered.push(...items.map(item => item.key));
                    });
                    return ordered;
                };

                // 左右两侧是主宾共同的溢出区：上方人员从上端向下，下方人员从下端向上。
                // 同一距离优先左右交替，确保任一方都不是只占用其中一侧。
                const sortLongtableExtensionSeatKeys = (keys, primarySide) => {
                    const sideLists = {left: [], right: []};
                    keys.forEach(key => {
                        const match = key.match(/^lt-(left|right)-(\d+)-(\d+)$/);
                        if(!match) return;
                        sideLists[match[1]].push({key, row: Number(match[2]), pos: Number(match[3])});
                    });
                    ['left', 'right'].forEach(side => {
                        sideLists[side].sort((a, b) => {
                            if(a.row !== b.row) return a.row - b.row;
                            return primarySide === 'bottom' ? b.pos - a.pos : a.pos - b.pos;
                        });
                    });
                    const ordered = [];
                    const maxLength = Math.max(sideLists.left.length, sideLists.right.length);
                    for(let index = 0; index < maxLength; index++) {
                        if(sideLists.left[index]) ordered.push(sideLists.left[index].key);
                        if(sideLists.right[index]) ordered.push(sideLists.right[index].key);
                    }
                    return ordered;
                };

                const buildLongtablePartyPlan = (keys, hostCount, guestCount, rule, hostSide, guestSide) => {
                    const plan = {host: [], guest: []};
                    const primaryKeys = side => keys.filter(key => key.startsWith(`lt-${side}-`));
                    plan.host.push(...sortLongtableSeatKeys(primaryKeys(hostSide), rule).slice(0, hostCount));
                    plan.guest.push(...sortLongtableSeatKeys(primaryKeys(guestSide), rule).slice(0, guestCount));

                    const hostOverflowCount = Math.max(hostCount - plan.host.length, 0);
                    const guestOverflowCount = Math.max(guestCount - plan.guest.length, 0);
                    const extensionKeys = keys.filter(key => key.startsWith('lt-left-') || key.startsWith('lt-right-'));
                    const hostExtensionKeys = sortLongtableExtensionSeatKeys(extensionKeys, hostSide).slice(0, hostOverflowCount);
                    const hostUsed = new Set(hostExtensionKeys);
                    const guestExtensionKeys = sortLongtableExtensionSeatKeys(extensionKeys, guestSide)
                        .filter(key => !hostUsed.has(key))
                        .slice(0, guestOverflowCount);
                    plan.host.push(...hostExtensionKeys);
                    plan.guest.push(...guestExtensionKeys);
                    return plan;
                };

                const compareSeatPrefixes = (a, b) => {
                    const aParts = a.split('-');
                    const bParts = b.split('-');
                    for(let i = 0; i < Math.max(aParts.length, bParts.length); i++) {
                        const av = aParts[i] || '';
                        const bv = bParts[i] || '';
                        const an = parseInt(av);
                        const bn = parseInt(bv);
                        if(!isNaN(an) && !isNaN(bn) && an !== bn) return an - bn;
                        if((isNaN(an) || isNaN(bn)) && av !== bv) return av.localeCompare(bv);
                    }
                    return 0;
                };

                const buildRoundtablePartyPlan = (keys, hostCount, guestCount) => {
                    const groups = {};
                    keys.forEach(key => {
                        const parts = key.split('-');
                        const prefix = parts.slice(0, -1).join('-');
                        if(!groups[prefix]) groups[prefix] = [];
                        groups[prefix].push({key, pos: parseInt(parts[parts.length - 1])});
                    });
                    const plan = {host: [], guest: []};
                    let hostRemaining = hostCount;
                    let guestRemaining = guestCount;
                    Object.keys(groups).sort(compareSeatPrefixes).forEach(prefix => {
                        if(hostRemaining + guestRemaining <= 0) return;
                        const items = groups[prefix];
                        const capacity = items.length;
                        let hostTake = Math.min(hostRemaining, Math.ceil(capacity / 2));
                        let guestTake = Math.min(guestRemaining, capacity - hostTake);
                        let spare = capacity - hostTake - guestTake;
                        const hostExtra = Math.min(hostRemaining - hostTake, spare);
                        hostTake += hostExtra;
                        spare -= hostExtra;
                        guestTake += Math.min(guestRemaining - guestTake, spare);

                        const hostPreferences = [...items].sort((a, b) => {
                            if(a.pos === 0) return -1;
                            if(b.pos === 0) return 1;
                            return b.pos - a.pos;
                        });
                        const hostKeys = hostPreferences.slice(0, hostTake).map(item => item.key);
                        const used = new Set(hostKeys);
                        const guestPreferences = items.filter(item => !used.has(item.key)).sort((a, b) => {
                            if(a.pos === 0) return 1;
                            if(b.pos === 0) return -1;
                            return a.pos - b.pos;
                        });
                        const guestKeys = guestPreferences.slice(0, guestTake).map(item => item.key);
                        plan.host.push(...hostKeys);
                        plan.guest.push(...guestKeys);
                        hostRemaining -= hostKeys.length;
                        guestRemaining -= guestKeys.length;
                    });
                    return plan;
                };

                const assignPartyGroups = (seatPlan, attendeeGroups) => {
                    const batchStamp = Date.now();
                    ['host', 'guest'].forEach(role => {
                        const assignedKeys = [];
                        const keys = seatPlan[role] || [];
                        const people = attendeeGroups[role] || [];
                        for(let i = 0; i < Math.min(keys.length, people.length); i++) {
                            const key = keys[i];
                            seats[key].attendee = people[i];
                            assignedKeys.push(key);
                            for(let batchKey in regionAssignmentRecords) {
                                const oldIndex = regionAssignmentRecords[batchKey].indexOf(key);
                                if(oldIndex > -1) regionAssignmentRecords[batchKey].splice(oldIndex, 1);
                            }
                        }
                        if(assignedKeys.length) regionAssignmentRecords[`batch_${batchStamp}_${role}`] = assignedKeys;
                    });
                    assignSelectedAtts.value = [];
                };

                const executeAssignment = () => {
                    if(assignSelectedAtts.value.length === 0) return alert('未勾选要排的人员');

                    if(isPartyLayout.value && assignConfig.targetZone === 'party_auto') {
                        const unclassified = assignSelectedAtts.value.filter(att => normalizePartyRole(att.party_role) === 'unassigned');
                        if(unclassified.length) {
                            return alert(`有 ${unclassified.length} 位所选人员尚未区分主宾，请先到“人员管理”完成分类。`);
                        }
                        const attendeeGroups = {
                            host: assignSelectedAtts.value.filter(att => normalizePartyRole(att.party_role) === 'host').sort((a, b) => getAttendeeOrder(a) - getAttendeeOrder(b)),
                            guest: assignSelectedAtts.value.filter(att => normalizePartyRole(att.party_role) === 'guest').sort((a, b) => getAttendeeOrder(a) - getAttendeeOrder(b)),
                        };
                        let seatPlan = {host: [], guest: []};

                        if(layout.layoutType === 'longtable') {
                            layout.longtableConfig.partySeatRule = assignConfig.rule || 'center_out';
                            const hostSide = layout.longtableConfig.hostSide || 'bottom';
                            const guestSide = layout.longtableConfig.guestSide || oppositeLongtableSide(hostSide);
                            const availablePartyKeys = Object.keys(seats).filter(key => /^lt-(top|bottom|left|right)-/.test(key) && seats[key].type === 'normal' && !seats[key].attendee);
                            seatPlan = buildLongtablePartyPlan(
                                availablePartyKeys,
                                attendeeGroups.host.length,
                                attendeeGroups.guest.length,
                                assignConfig.rule,
                                hostSide,
                                guestSide
                            );
                            if(seatPlan.host.length < attendeeGroups.host.length) {
                                return alert(`主方的${longtableSideName(hostSide)}及左右两侧延伸席仅能安排 ${seatPlan.host.length} 人，无法安排所选的 ${attendeeGroups.host.length} 人。`);
                            }
                            if(seatPlan.guest.length < attendeeGroups.guest.length) {
                                return alert(`宾方的${longtableSideName(guestSide)}及左右两侧延伸席仅能安排 ${seatPlan.guest.length} 人，无法安排所选的 ${attendeeGroups.guest.length} 人。`);
                            }
                        } else {
                            const roundKeys = Object.keys(seats).filter(key => key.startsWith('rt-') && seats[key].type === 'normal' && !seats[key].attendee);
                            if(roundKeys.length < attendeeGroups.host.length + attendeeGroups.guest.length) {
                                return alert(`圆桌仅剩 ${roundKeys.length} 个空位，无法安排 ${attendeeGroups.host.length + attendeeGroups.guest.length} 人。`);
                            }
                            seatPlan = buildRoundtablePartyPlan(roundKeys, attendeeGroups.host.length, attendeeGroups.guest.length);
                        }

                        saveHistory();
                        assignPartyGroups(seatPlan, attendeeGroups);
                        return;
                    }

                    saveHistory();

                    let targetCoords = [];
                    if(layout.layoutType !== 'classroom') {
                        for(let key in seats) {
                            if(seats[key].type !== 'normal' || seats[key].attendee) continue;
                            if(assignConfig.targetZone === 'all') {
                                targetCoords.push(key);
                            } else if(layout.layoutType === 'utable' && assignConfig.targetZone === 'u_auto' && key.startsWith('u-')) {
                                targetCoords.push(key);
                            } else if(key.startsWith(assignConfig.targetZone + '-')) {
                                targetCoords.push(key);
                            }
                        }
                    } else {
                        for(let key in seats) {
                            const p = key.split('-'); const z = p[0]; const r = parseInt(p[1]); const c = parseInt(p[2]);
                            if(seats[key].type !== 'normal' || seats[key].attendee) continue;

                            if(assignConfig.targetZone === 'podium' && z === 'p') targetCoords.push({z, r, c});
                            else if(assignConfig.targetZone === 'all_m' && z === 'm' && !seats[key].regionId) targetCoords.push({z, r, c});
                            else if(assignConfig.targetZone.startsWith('reg_') && seats[key].regionId === assignConfig.targetZone.split('_')[1]) targetCoords.push({z, r, c});
                        }
                    }

                    if(targetCoords.length < assignSelectedAtts.value.length) {
                        if(!confirm(`空位仅剩 ${targetCoords.length} 个，人员有 ${assignSelectedAtts.value.length}，超出部分将抛弃，继续吗？`)) { history.value.pop(); return; }
                    }

                    let orderedCoords = [];
                    if(layout.layoutType === 'utable') {
                        orderedCoords = sortUtableSeatKeys(targetCoords);
                    } else if(layout.layoutType !== 'classroom') {
                        // 新布局类型排序
                        const rule = assignConfig.rule;

                        // 先按前缀分组（同一边/段/圈放在一起）
                        const groupMap = {};
                        targetCoords.forEach(k => {
                            const parts = k.split('-');
                            const posIdx = parseInt(parts[parts.length - 1]);
                            const prefix = parts.slice(0, parts.length - 1).join('-');
                            if (!groupMap[prefix]) groupMap[prefix] = [];
                            groupMap[prefix].push({key: k, pos: posIdx});
                        });

                        // 按前缀排序
                        const sortedGroups = Object.keys(groupMap).sort((a, b) => {
                            if (layout.layoutType === 'longtable') {
                                const squareOrder = ['lt-head-left', 'lt-head-right', 'lt-top', 'lt-right', 'lt-bottom', 'lt-left'];
                                const rank = prefix => {
                                    const index = squareOrder.findIndex(item => prefix.startsWith(item));
                                    return index === -1 ? squareOrder.length : index;
                                };
                                const rankDiff = rank(a) - rank(b);
                                if (rankDiff !== 0) return rankDiff;
                            }
                            const aParts = a.split('-'); const bParts = b.split('-');
                            for(let i=0; i<Math.max(aParts.length, bParts.length); i++) {
                                const va = aParts[i] || ''; const vb = bParts[i] || '';
                                const na = parseInt(va); const nb = parseInt(vb);
                                if(!isNaN(na) && !isNaN(nb)) { if(na !== nb) return na - nb; }
                                else { const cmp = va.localeCompare(vb); if(cmp !== 0) return cmp; }
                            }
                            return 0;
                        });

                        sortedGroups.forEach(prefix => {
                            const items = groupMap[prefix];
                            if (layout.layoutType === 'longtable') {
                                // 新算法：从中间向两边交替，区分单双排
                                const n = items.length;
                                items.sort((a, b) => a.pos - b.pos);
                                if (n === 0) return;

                                // prefix = "lt-{side}-{row}"；相邻排交替左右优先，视觉上更均衡。
                                const preParts = prefix.split('-');
                                const rowIdx = parseInt(preParts[2]) || 0; // 0-based row index
                                const isEvenRow = (rowIdx % 2 === 0);
                                const physicalCenter = (n - 1) / 2;

                                if (rule === 'center_out') {
                                    items.sort((a, b) => {
                                        const distanceDiff = Math.abs(a.pos - physicalCenter) - Math.abs(b.pos - physicalCenter);
                                        if (Math.abs(distanceDiff) > 0.01) return distanceDiff;
                                        return isEvenRow ? b.pos - a.pos : a.pos - b.pos;
                                    });
                                    items.forEach(item => orderedCoords.push(item.key));
                                } else if (rule === 'rtl') {
                                    items.sort((a, b) => b.pos - a.pos);
                                    items.forEach(item => orderedCoords.push(item.key));
                                } else if (rule === 'snake') {
                                    // 蛇形：偶数排左到右，奇数排右到左
                                    if (isEvenRow) {
                                        items.sort((a, b) => a.pos - b.pos);
                                    } else {
                                        items.sort((a, b) => b.pos - a.pos);
                                    }
                                    items.forEach(item => orderedCoords.push(item.key));
                                } else {
                                    items.sort((a, b) => a.pos - b.pos);
                                    items.forEach(item => orderedCoords.push(item.key));
                                }
                            } else if (layout.layoutType === 'roundtable') {
                                if (rule === 'banquet') {
                                    // 中式宴请排位算法
                                    // 12点钟方向(pos 0)为主位(主家第1人)
                                    // 顺时针方向(pos 1,2,3...)为宾客席
                                    // 逆时针方向(pos n-1,n-2...)为主家席
                                    // 排位顺序：主位 → 主宾右1 → 主家左1 → 宾客右2 → 主家左2 → ...
                                    const n = items.length;
                                    items.sort((a, b) => a.pos - b.pos);
                                    const result = [];
                                    // pos 0 = 主位 (host #1)
                                    if (n > 0) result.push(items[0]);
                                    // 交替填充：右侧（宾客，pos 1→2→3...）和左侧（主家，pos n-1→n-2→n-3...）
                                    let right = 1, left = n - 1;
                                    while (right <= left) {
                                        if (right <= left) { result.push(items[right]); right++; }
                                        if (right <= left) { result.push(items[left]); left--; }
                                    }
                                    result.forEach(item => orderedCoords.push(item.key));
                                } else if (rule === 'counterclockwise') {
                                    // 逆时针从12点: pos0, pos N-1, pos N-2, ..., pos1
                                    items.sort((a, b) => a.pos - b.pos);
                                    const result = [items[0]];
                                    for(let i = items.length - 1; i > 0; i--) result.push(items[i]);
                                    result.forEach(item => orderedCoords.push(item.key));
                                } else {
                                    // clockwise default: pos0(12点) → pos1(1点) → pos2(2点) → ...
                                    items.sort((a, b) => a.pos - b.pos);
                                    items.forEach(item => orderedCoords.push(item.key));
                                }
                            }
                        });
                    } else if(assignConfig.targetZone === 'podium') {
                        targetCoords.sort((a,b) => {
                            let valA_r = parseFloat(layout.podium.rowLabels[a.r]); if(isNaN(valA_r)) valA_r = a.r * 1000;
                            let valB_r = parseFloat(layout.podium.rowLabels[b.r]); if(isNaN(valB_r)) valB_r = b.r * 1000;
                            if(valA_r !== valB_r) return valA_r - valB_r;

                            let valA_c = parseFloat(layout.podium.colLabels[a.c]); if(isNaN(valA_c)) valA_c = a.c * 1000;
                            let valB_c = parseFloat(layout.podium.colLabels[b.c]); if(isNaN(valB_c)) valB_c = b.c * 1000;
                            return valA_c - valB_c;
                        });
                        orderedCoords = targetCoords;
                    } else if(assignConfig.rule === 'custom_seq') {
                        if(!assignConfig.targetZone.startsWith('reg_')) {
                            alert('自定义顺序排座仅支持具体的“特定色彩区”！请在上一步选择具体区域。');
                            history.value.pop(); return;
                        }
                        const regId = assignConfig.targetZone.split('_')[1];
                        const customSeq = layout.customSequences?.[regId] || [];
                        if (customSeq.length === 0) {
                            alert('该区域尚未定义自定义顺序连线！请前往“会场设置 -> 自定义区域排座次序”中进行点击定义。');
                            history.value.pop(); return;
                        }
                        let sequenced = [];
                        let unsequenced = [];
                        targetCoords.forEach(item => {
                            const k = `${item.z}-${item.r}-${item.c}`;
                            const sIdx = customSeq.indexOf(k);
                            if(sIdx > -1) { item._sIdx = sIdx; sequenced.push(item); }
                            else { unsequenced.push(item); }
                        });
                        sequenced.sort((a,b) => a._sIdx - b._sIdx);
                        unsequenced.sort((a,b) => { if(a.r !== b.r) return a.r - b.r; return a.c - b.c; });
                        orderedCoords = [...sequenced, ...unsequenced];
                    } else {
                        const rowGroups = {};
                        targetCoords.forEach(item => { if(!rowGroups[item.r]) rowGroups[item.r]=[]; rowGroups[item.r].push(item); });
                        const rKeys = Object.keys(rowGroups).sort((a,b)=>a-b);

                        let physicalCenter = (layout.main.cols - 1) / 2.0;
                        if (assignConfig.targetZone.startsWith('reg_')) {
                            const regId = assignConfig.targetZone.split('_')[1];
                            let minC = Infinity, maxC = -Infinity;
                            for(let r=0; r<layout.main.rows; r++) {
                                for(let c=0; c<layout.main.cols; c++) {
                                    if (seats[`m-${r}-${c}`]?.regionId === regId) {
                                        if (c < minC) minC = c;
                                        if (c > maxC) maxC = c;
                                    }
                                }
                            }
                            if (minC <= maxC) physicalCenter = (minC + maxC) / 2.0;
                        }

                        const isCenterOnSeat = (physicalCenter % 1 === 0);

                        rKeys.forEach((rk, idx) => {
                            const items = rowGroups[rk];
                            const rule = assignConfig.rule;
                            if (rule === 'ltr') {
                                items.sort((a,b)=>a.c-b.c);
                                orderedCoords.push(...items);
                            }
                            else if (rule === 'rtl') {
                                items.sort((a,b)=>b.c-a.c);
                                orderedCoords.push(...items);
                            }
                            else if (rule === 'snake') {
                                items.sort((a,b)=>a.c-b.c);
                                if(idx%2!==0) items.reverse();
                                orderedCoords.push(...items);
                            }
                            else if (rule === 'center_lr' || rule === 'center_rl') {
                                items.sort((a, b) => {
                                    const distA = Math.abs(a.c - physicalCenter);
                                    const distB = Math.abs(b.c - physicalCenter);
                                    if (Math.abs(distA - distB) > 0.01) return distA - distB;
                                    if (rule === 'center_lr') return isCenterOnSeat ? (b.c - a.c) : (a.c - b.c);
                                    else return isCenterOnSeat ? (a.c - b.c) : (b.c - a.c);
                                });
                                orderedCoords.push(...items);
                            }
                        });
                    }

                    assignSelectedAtts.value.sort((a,b) => {
                        return attendees.value.findIndex(x=>x.id===a.id) - attendees.value.findIndex(x=>x.id===b.id);
                    });

                    const seq = [];
                    for(let i=0; i<orderedCoords.length; i++) {
                        if(i >= assignSelectedAtts.value.length) break;
                        const pos = orderedCoords[i];
                        const k = typeof pos === 'string' ? pos : `${pos.z}-${pos.r}-${pos.c}`;
                        seats[k].attendee = assignSelectedAtts.value[i];
                        seq.push(k);
                    }

                    seq.forEach(k => {
                        for(let bKey in regionAssignmentRecords) {
                            const idx = regionAssignmentRecords[bKey].indexOf(k);
                            if(idx > -1) regionAssignmentRecords[bKey].splice(idx, 1);
                        }
                    });
                    regionAssignmentRecords['batch_'+Date.now()] = seq;
                    assignSelectedAtts.value = [];
                };

                const selectAllUnassigned = () => { assignSelectedAtts.value = [...unassignedFilteredAttendees.value]; };

                const saveHistory = () => { history.value.push(JSON.parse(JSON.stringify(seats))); if(history.value.length>15) history.value.shift(); };
                const undo = () => { if(history.value.length>0) Object.assign(seats, history.value.pop()); };

                const saveTemplate = async () => {
                    if(!layoutName.value) return alert('请输入模板名称');
                    const config = {
                        layoutType: layout.layoutType,
                        podium: layout.podium,
                        main: layout.main,
                        longtableConfig: layout.longtableConfig,
                        roundtableConfig: layout.roundtableConfig,
                        utableConfig: layout.utableConfig,
                        regionTypes: layout.regionTypes,
                        seatDisplayConfig: layout.seatDisplayConfig,
                        typeMap: {},
                        customSequences: layout.customSequences
                    };
                    for(let k in seats) {
                        if(seats[k].type!=='normal' || seats[k].regionId || seats[k].displayConfig || seats[k].customLabel) {
                            config.typeMap[k] = {t:seats[k].type, r:seats[k].regionId, d:seats[k].displayConfig, l:seats[k].customLabel};
                        }
                    }
                    await fetch('/api/templates', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({name:layoutName.value, config})});
                    loadGlobalData(); alert('模板保存成功');
                };
                const loadTemplate = (id) => {
                    if(!id) return;
                    const t = templates.value.find(x=>x.id==id);
                    if(!confirm('应用模板将重置本场会议的排位格局，是否继续？')) return;
                    const c = t.config;
                    layout.layoutType = c.layoutType || 'classroom';
                    if(c.podium) Object.assign(layout.podium, c.podium);
                    if(c.main) Object.assign(layout.main, c.main);
                    if(c.longtableConfig) Object.assign(layout.longtableConfig, c.longtableConfig);
                    if(c.roundtableConfig) Object.assign(layout.roundtableConfig, c.roundtableConfig);
                    if(c.utableConfig) Object.assign(layout.utableConfig, c.utableConfig);
                    layout.regionTypes = c.regionTypes || [];
                    if(c.seatDisplayConfig) layout.seatDisplayConfig = c.seatDisplayConfig;
                    if(c.customSequences) layout.customSequences = c.customSequences;
                    else layout.customSequences = {};
                    history.value = [];
                    for(let k in seats) delete seats[k];
                    for(let k in regionAssignmentRecords) delete regionAssignmentRecords[k];
                    initSeats();
                    for(let k in (c.typeMap || {})) {
                        if(seats[k]) {
                            seats[k].type = c.typeMap[k].t || 'normal';
                            seats[k].regionId = c.typeMap[k].r || null;
                            seats[k].displayConfig = c.typeMap[k].d || null;
                            seats[k].customLabel = c.typeMap[k].l || null;
                        }
                    }
                    resetAssignmentForLayout();
                    autoSave();
                };
                const deleteTemplate = async (id) => {
                    if(!id) return alert('请先选择要删除的模板');
                    if(!confirm('确定删除该模板吗？')) return;
                    await fetch(`/api/templates/${id}`, { method: 'DELETE' });
                    selectedTemplateId.value = '';
                    loadGlobalData();
                };

                const insertVar = (v) => {
                    const ta = document.getElementById('smsTemplate');
                    const start = ta.selectionStart;
                    layout.sms_template = layout.sms_template.substring(0, start) + v + layout.sms_template.substring(ta.selectionEnd);
                    setTimeout(()=> { ta.focus(); ta.selectionStart = ta.selectionEnd = start + v.length; }, 0);
                    autoSave();
                };

                const insertExportVar = (v) => {
                    const ta = document.getElementById('exportTemplate');
                    const start = ta.selectionStart;
                    layout.export_template = layout.export_template.substring(0, start) + v + layout.export_template.substring(ta.selectionEnd);
                    setTimeout(()=> { ta.focus(); ta.selectionStart = ta.selectionEnd = start + v.length; }, 0);
                    autoSave();
                };

                const exportExcel = async () => {
                    const exportSeats = {};
                    for(let k in seats) {
                        exportSeats[k] = {
                            type: seats[k].type, attendee: seats[k].attendee, customLabel: seats[k].customLabel,
                            region_name: seats[k].regionId ? layout.regionTypes.find(r=>r.id===seats[k].regionId)?.name : '未分区'
                        };
                    }
                    const payload = {
                        conf_name: meetingNameInput.value,
                        layoutType: layout.layoutType,
                        podium: layout.podium, main: layout.main,
                        longtableConfig: layout.longtableConfig, roundtableConfig: layout.roundtableConfig, utableConfig: layout.utableConfig,
                        seats_data: exportSeats,
                        sms_template: layout.sms_template, export_template: layout.export_template.replace(/\\n/g, '\n'),
                        region_types: layout.regionTypes
                    };
                    const res = await fetch('/api/export', { method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(payload) });
                    if(!res.ok) {
                        const errorData = await res.json().catch(() => ({}));
                        return alert(errorData.error || '导出失败，请稍后重试');
                    }
                    const blob = await res.blob();
                    const downloadUrl = URL.createObjectURL(blob);
                    const a = document.createElement('a'); a.href = downloadUrl; a.download = `${payload.conf_name}排座导出.xlsx`; a.click();
                    URL.revokeObjectURL(downloadUrl);
                };

                const filteredAttendees = computed(() => {
                    let result = partySortedAttendees.value;
                    if(isPartyLayout.value && attendeeRoleFilter.value !== 'all') {
                        result = result.filter(att => normalizePartyRole(att.party_role) === attendeeRoleFilter.value);
                    }
                    if(!attendeeSearch.value) return result;
                    const q = attendeeSearch.value.toLowerCase();
                    return result.filter(a => a.name.toLowerCase().includes(q) || (a.unit && a.unit.toLowerCase().includes(q)));
                });
                const unassignedAttendees = computed(() => {
                    const assignedIds = Object.values(seats).filter(s => s && s.attendee).map(s => s.attendee.id);
                    return partySortedAttendees.value.filter(a => !assignedIds.includes(a.id));
                });
                const unassignedFilteredAttendees = computed(() => {
                    let result = unassignedAttendees.value;
                    if(isPartyLayout.value && assignRoleFilter.value !== 'all') {
                        result = result.filter(att => normalizePartyRole(att.party_role) === assignRoleFilter.value);
                    }
                    if(!assignSearch.value) return result;
                    const q = assignSearch.value.toLowerCase();
                    return result.filter(a => a.name.toLowerCase().includes(q) || (a.unit && a.unit.toLowerCase().includes(q)));
                });
                const seatMenuFilteredAtts = computed(() => unassignedAttendees.value.filter(a => a.name.includes(seatMenuSearch.value)).slice(0, 40));

                const seatMenuLabel = computed(() => {
                    if(seatMenu.zone === '') return '';
                    if (seatMenu.zone === 'p' || seatMenu.zone === 'm') {
                        const obj = seatMenu.zone === 'p' ? layout.podium : layout.main;
                        const key = `${seatMenu.zone}-${seatMenu.r}-${seatMenu.c}`;
                        const customLabel = seats[key]?.customLabel;
                        const finalLabel = (customLabel !== undefined && customLabel !== null) ? customLabel : formatColLabel(obj.colLabels[seatMenu.c]);
                        return `${obj.rowLabels[seatMenu.r]}排 ${finalLabel}`;
                    }
                    const key = makeSeatKey(seatMenu.zone, seatMenu.r, seatMenu.c);
                    const customLabel = seats[key]?.customLabel;
                    if (seatMenu.zone.startsWith('lt-head-')) {
                        const sideName = seatMenu.zone.endsWith('left') ? '左侧' : '右侧';
                        return `桌外主位-${sideName} ${customLabel || `第${seatMenu.r + 1}座`}`;
                    }
                    if (seatMenu.zone.startsWith('lt-')) {
                        const side = seatMenu.zone.replace('lt-', '');
                        const sideName = layout.longtableConfig.sides[side]?.label || side;
                        return `${sideName} 第${seatMenu.r + 1}排 ${customLabel || `第${seatMenu.c + 1}座`}`;
                    }
                    if (seatMenu.zone.startsWith('rt-')) {
                        const parts = seatMenu.zone.split('-');
                        const tableNo = Number(parts[1]) + 1;
                        const ringId = parts.slice(2).join('-');
                        const ringName = layout.roundtableConfig.rings.find(r => r.id === ringId)?.label || ringId;
                        return `第${tableNo}桌-${ringName} ${customLabel || `第${seatMenu.r + 1}座`}`;
                    }
                    if (seatMenu.zone === 'u-leader') return `领导席 ${customLabel || `第${getUtableLeaderOrderNumber(seatMenu.r)}位`}`;
                    if (seatMenu.zone === 'u-left' || seatMenu.zone.startsWith('u-left-')) {
                        const column = seatMenu.zone === 'u-left' ? 0 : Number.parseInt(seatMenu.zone.split('-').pop(), 10);
                        return `左侧参会席 第${column + 1}列 ${customLabel || `距领导第${seatMenu.r + 1}位`}`;
                    }
                    if (seatMenu.zone === 'u-right' || seatMenu.zone.startsWith('u-right-')) {
                        const column = seatMenu.zone === 'u-right' ? 0 : Number.parseInt(seatMenu.zone.split('-').pop(), 10);
                        return `右侧参会席 第${column + 1}列 ${customLabel || `距领导第${seatMenu.r + 1}位`}`;
                    }
                    return customLabel || `第${seatMenu.r + 1}座`;
                });

                const seatMenuHasDisplayConfig = computed(() => {
                    if (!seatMenu.show) return false;
                    const key = makeSeatKey(seatMenu.zone, seatMenu.r, seatMenu.c);
                    return !!(seats[key]?.displayConfig);
                });

                // ===== 长会议桌式计算属性（内部保留 longtable 命名兼容已有数据） =====
                const longtableHasHead = computed(() => layout.longtableConfig.headSeats?.enabled || false);
                const longtableHasLeft = computed(() => (layout.longtableConfig.sides.left?.rows || 0) > 0);
                const longtableHasRight = computed(() => (layout.longtableConfig.sides.right?.rows || 0) > 0);
                const oppositeLongtableSide = (side) => ({top: 'bottom', bottom: 'top', left: 'right', right: 'left'})[side] || 'top';
                const longtableSideName = (side) => ({top: '上侧', bottom: '下侧', left: '左侧', right: '右侧'})[side] || side;
                const setLongtablePartySide = (role, side) => {
                    if(!['top', 'bottom'].includes(side)) return;
                    const oppositeSide = oppositeLongtableSide(side);
                    if(role === 'host') {
                        layout.longtableConfig.hostSide = side;
                        layout.longtableConfig.guestSide = oppositeSide;
                    } else {
                        layout.longtableConfig.guestSide = side;
                        layout.longtableConfig.hostSide = oppositeSide;
                    }
                };
                const longtableGuestSide = computed(() => layout.longtableConfig.guestSide || oppositeLongtableSide(layout.longtableConfig.hostSide || 'bottom'));
                const longtableHostSideLabel = computed(() => longtableSideName(layout.longtableConfig.hostSide || 'bottom'));
                const longtableGuestSideLabel = computed(() => longtableSideName(longtableGuestSide.value));

                // 横向与纵向分别按对应两侧的座位跨度计算，保证相对桌面中轴居中对称。
                const longtableSideSpan = (side) => {
                    const cfg = layout.longtableConfig.sides[side] || {};
                    const seatCount = Math.max(Number(cfg.seatsPerRow) || 0, 0);
                    const seatGap = Math.max(Number(cfg.seatGap) || 0, 0);
                    if(seatCount === 0) return 0;
                    return seatCount * 74 + Math.max(seatCount - 1, 0) * seatGap;
                };
                const longtableMinimumSpan = 3 * 74 + 2 * 8;
                const longtableTableWidth = computed(() => {
                    return Math.max(longtableSideSpan('top'), longtableSideSpan('bottom'), longtableMinimumSpan) + 28;
                });
                const longtableTableHeight = computed(() => {
                    return Math.max(longtableSideSpan('left'), longtableSideSpan('right'), longtableMinimumSpan) + 28;
                });

                const longtableHasSides = computed(() => longtableHasLeft.value || longtableHasRight.value);
                const longtableGridCols = computed(() => longtableHasSides.value ? '3' : '1');
                const longtableTopRow = computed(() => '1');
                const longtableTableRow = computed(() => '2');
                const longtableBottomRow = computed(() => '3');
                const longtableTableGridCol = computed(() => longtableHasSides.value ? '2' : '1');

                const longtableGridStyle = computed(() => {
                    const hasSides = longtableHasSides.value;
                    const tw = longtableTableWidth.value;
                    return {
                        gridTemplateColumns: hasSides ? `auto ${tw}px auto` : `${tw}px`,
                        gridTemplateRows: 'repeat(3, auto)',
                    };
                });

                const utableSeatSpan = (count, gap) => {
                    const safeCount = Math.max(Number(count) || 0, 0);
                    if(safeCount === 0) return 0;
                    return safeCount * 74 + Math.max(safeCount - 1, 0) * Math.max(Number(gap) || 0, 0);
                };
                const utableGridStyle = computed(() => {
                    const leaderWidth = Math.max(utableSeatSpan(layout.utableConfig.leaderSeatCount, layout.utableConfig.leaderGap), 3 * 74 + 16);
                    const sideColumnWidth = columnCount => utableSeatSpan(columnCount, layout.utableConfig.sideGap);
                    const leftWidth = sideColumnWidth(layout.utableConfig.leftColumnCount);
                    const rightWidth = sideColumnWidth(layout.utableConfig.rightColumnCount);
                    const sideHeight = Math.max(
                        utableSeatSpan(layout.utableConfig.leftSeatCount, layout.utableConfig.sideGap),
                        utableSeatSpan(layout.utableConfig.rightSeatCount, layout.utableConfig.sideGap),
                        3 * 74 + 16
                    );
                    return {
                        gridTemplateColumns: `${leftWidth}px ${leaderWidth}px ${rightWidth}px`,
                        gridTemplateRows: `auto ${sideHeight}px auto`,
                        '--u-accent': layout.utableConfig.accentColor || '#475569'
                    };
                });

                const extraKeys = computed(() => {
                    const keys = new Set();
                    attendees.value.forEach(a => {
                        if(a.extra) {
                            try {
                                const ex = typeof a.extra === 'string' ? JSON.parse(a.extra) : a.extra;
                                if(ex && typeof ex === 'object') {
                                    Object.keys(ex).forEach(k => {
                                        if(!['序号', '姓名', '单位', '电话', '类别'].includes(k)) {
                                            keys.add(k);
                                        }
                                    });
                                }
                            } catch(e) {}
                        }
                    });
                    return Array.from(keys);
                });

                return {
                    tabs, activeTab, meetings, currentMeetingId, meetingNameInput, currentMeetingObj, templates, selectedTemplateId, saveStatus, layout, seats, attendees, history,
                    clickMode, activeRegionId, layoutName, attendeeSearch, attendeeRoleFilter, selectedAttIds, dragTargetIndex,
                    assignSearch, assignRoleFilter, assignRangeInput, assignConfig, assignSelectedAtts, showModal, modalForm,
                    seatMenu, seatMenuSearch, seatMenuFilteredAtts, seatMenuLabel, seatMenuHasDisplayConfig, axisMenu, seatDisplayEdit, extraKeys, batchSelectState,
                    longtableHasHead, longtableHasLeft, longtableHasRight, longtableHostSideLabel, longtableGuestSideLabel, longtableTableWidth, longtableTableHeight, longtableHasSides, longtableGridCols, longtableTopRow, longtableTableRow, longtableBottomRow, longtableTableGridCol, longtableGridStyle, utableGridStyle,
                    isPartyLayout, partyCounts, normalizePartyRole, partyRoleLabel, partyRoleClass, getAttendeeOrder, isFirstInAttendeeScope, isLastInAttendeeScope,
                    filteredAttendees, unassignedFilteredAttendees, isAllSelected, toggleSelectAllFiltered,
                    loadMeetingState, createMeeting, deleteCurrentMeeting, renameMeeting, initSeats, scheduleUtableSeatSync, handleLabelChange, syncPodiumLabels, formatColLabel, getSeatColLabelUI, setCustomLabel, isDuplicateSeat, isSeatBgDark, clearCustomSequence, handleSeatClick,
                    addRegionType, removeRegionType, getSeatRegionColor, seatClass, seatStyle, modifyAxis,
                    onDragStart, onDragOver, onDragLeave, onDrop, moveAttendee, setAttendeePosition, applyRangeSelection, getGlobalIndex, getAttendeeFieldValue,
                    uploadFile, bulkDeleteAttendees, setSelectedPartyRole, setAttendeePartyRole, openAddModal, editAttendee, saveModalForm,
                    handleSeatRightClick, handleAxisRightClick, closeContextMenus, setAxisStatus, forceAssign, clearSeat, removeAndShift,
                    openSeatDisplayEdit, saveSeatDisplayConfig, resetSeatDisplayConfig, getSeatDisplayConfig,
                    seatStyleLongtable, seatStyleLongtableRow, roundtableSeatStyle, makeSeatKey, getUtableLeaderOrderNumber,
                    clearTargetZone, executeAssignment, selectAllUnassigned, undo,
                    saveTemplate, loadTemplate, deleteTemplate, insertVar, insertExportVar, exportExcel, autoSave,
                    onLayoutTypeChange, prevLayoutType, addRoundtableRing, removeRoundtableRing, setLongtablePartySide
                };
            }
        }).mount('#app');
    </script>
</body>
</html>
"""

@app.route('/')
def index():
    return HTML_TEMPLATE

if __name__ == '__main__':
    init_db()
    app.run(host='0.0.0.0', port=5000, debug=True)
