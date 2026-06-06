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
                  name TEXT, unit TEXT, category TEXT, phone TEXT, extra JSON)''')
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
            rows.append(d)
        conn.close()
        return jsonify(rows)
    elif request.method == 'POST':
        data = request.json
        c.execute("SELECT MAX(sort_order) FROM attendees WHERE meeting_id=?", (m_id,))
        max_sort = c.fetchone()[0] or 0
        c.execute("INSERT INTO attendees (meeting_id, sort_order, name, unit, category, phone, extra) VALUES (?, ?, ?, ?, ?, ?, ?)",
                  (m_id, max_sort + 1, data.get('name',''), data.get('unit',''), data.get('category',''), data.get('phone',''), json.dumps(data.get('extra', {}))))
        conn.commit()
        conn.close()
        return jsonify({"status": "success"})
    elif request.method == 'PUT':
        data = request.json
        c.execute("UPDATE attendees SET name=?, unit=?, category=?, phone=?, extra=? WHERE id=? AND meeting_id=?",
                  (data.get('name',''), data.get('unit',''), data.get('category',''), data.get('phone',''), json.dumps(data.get('extra', {})), data.get('id'), m_id))
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
        
        if idx_col:
            df[idx_col] = pd.to_numeric(df[idx_col], errors='coerce').fillna(9999)
            df = df.sort_values(by=idx_col)

        c.execute("SELECT MAX(sort_order) FROM attendees WHERE meeting_id=?", (m_id,))
        current_sort = c.fetchone()[0] or 0

        for _, row in df.iterrows():
            if not name_col: break
            name = str(row[name_col]).strip()
            if not name: continue
            
            unit = str(row[unit_col]).strip() if unit_col else ''
            category = str(row[cat_col]).strip() if cat_col else ''
            phone = str(row[phone_col]).strip() if phone_col else ''
            
            extra = {col: str(row[col]) for col in cols if col not in ['序号', '姓名', '单位', '电话', '类别']}
            current_sort += 1
            
            c.execute("SELECT id FROM attendees WHERE name=? AND unit=? AND meeting_id=?", (name, unit, m_id))
            existing = c.fetchone()
            if existing:
                c.execute("UPDATE attendees SET category=?, phone=?, extra=?, sort_order=? WHERE id=?", 
                          (category, phone, json.dumps(extra), current_sort, existing[0]))
            else:
                c.execute("INSERT INTO attendees (meeting_id, sort_order, name, unit, category, phone, extra) VALUES (?, ?, ?, ?, ?, ?, ?)", 
                          (m_id, current_sort, name, unit, category, phone, json.dumps(extra)))
                
        conn.commit()
        conn.close()
        return jsonify({"status": "success", "message": "名单导入成功！已按文件顺序排序。"})
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
    if label_str.endswith('座') or label_str.endswith('号'): return label_str
    try:
        int(label_str)
        return label_str + '座'
    except ValueError:
        return label_str

def get_seat_col_label(s_data, conf, idx):
    custom = s_data.get('customLabel')
    if custom is not None and custom != '': return str(custom)
    labels = conf.get('colLabels', [])
    return format_col_label(labels[idx] if idx < len(labels) else str(idx + 1))

@app.route('/api/export', methods=['POST'])
def export_excel():
    try:
        data = request.json
        conf_name = data.get('conf_name', '未命名会议')
        podium = data.get('podium', {})
        main = data.get('main', {})
        seats_data = data.get('seats_data', {}) 
        sms_template = data.get('sms_template', '')
        export_template = data.get('export_template', '{姓名}\n{单位}')
        
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

        max_cols = max(podium.get('cols', 0), main.get('cols', 0))
        offset_p = max(0, (max_cols - podium.get('cols', 0)) // 2) + 2 if podium.get('cols', 0) > 0 else 2
        offset_m = max(0, (max_cols - main.get('cols', 0)) // 2) + 2 if main.get('cols', 0) > 0 else 2

        for c in range(max_cols + 2): ws1.column_dimensions[get_column_letter(c+1)].width = 15

        ws1.merge_cells(start_row=1, start_column=1, end_row=2, end_column=max_cols+2)
        cell_title = ws1.cell(row=1, column=1, value=conf_name)
        cell_title.font = Font(size=20, bold=True)
        cell_title.alignment = align_center
        current_row = 4

        if podium.get('rows', 0) > 0:
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

        if main.get('rows', 0) > 0:
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
                        cell.fill = fill_seat; cell.font = Font(color=main_text_color)
                        if att: cell.value = format_custom_text(att, export_template)
                        else: cell.value = f"{r_label}排\n{c_label}"
                current_row += 1

        # --- Sheet 2: 签到表 ---
        ws2 = wb.create_sheet(title="参会人员签到表")
        headers = ["大区", "块区", "排号", "座/列号", "姓名", "单位", "类别", "电话", "签到"]
        ws2.append(headers)
        for col in range(1, len(headers)+1):
            ws2.cell(row=1, column=col).font = Font(bold=True)
            ws2.column_dimensions[get_column_letter(col)].width = 15
            
        for key, s_data in seats_data.items():
            att = s_data.get('attendee')
            if att:
                parts = key.split('-')
                zone, r, c = parts[0], int(parts[1]), int(parts[2])
                if zone == 'p':
                    r_label = get_label(podium, 'row', r); c_label = get_seat_col_label(s_data, podium, c); zone_name = "主席台"
                else:
                    r_label = get_label(main, 'row', r); c_label = get_seat_col_label(s_data, main, c); zone_name = "主会场"
                ws2.append([zone_name, s_data.get('region_name', '未分区'), r_label, c_label, 
                            att.get('name',''), att.get('unit',''), att.get('category',''), att.get('phone',''), ""])

        # --- Sheet 3: 短信清单 ---
        ws3 = wb.create_sheet(title="短信群发清单")
        ws3.append(["姓名", "单位", "电话", "短信内容"])
        for col in range(1, 5): ws3.cell(row=1, column=col).font = Font(bold=True)
        ws3.column_dimensions['A'].width = 15; ws3.column_dimensions['B'].width = 25
        ws3.column_dimensions['C'].width = 15; ws3.column_dimensions['D'].width = 80

        for key, s_data in seats_data.items():
            att = s_data.get('attendee')
            if att:
                parts = key.split('-')
                zone, r, c = parts[0], int(parts[1]), int(parts[2])
                r_label = get_label(podium, 'row', r) if zone=='p' else get_label(main, 'row', r)
                c_label = get_seat_col_label(s_data, podium, c) if zone=='p' else get_seat_col_label(s_data, main, c)
                
                msg = sms_template
                msg = msg.replace('{姓名}', att.get('name','')).replace('{单位}', att.get('unit',''))\
                         .replace('{类别}', att.get('category','')).replace('{电话}', att.get('phone',''))\
                         .replace('{排号}', r_label).replace('{座号}', c_label)
                extra = att.get('extra', {})
                if isinstance(extra, str): 
                    try: extra = json.loads(extra)
                    except: extra = {}
                for k, v in extra.items():
                    msg = msg.replace(f'{{{k}}}', str(v))
                
                ws3.append([att.get('name',''), att.get('unit',''), att.get('phone',''), msg])

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

                    <h2 class="text-sm font-bold border-l-4 border-indigo-500 pl-2 mb-2 mt-4 bg-indigo-50 py-1">🎨 画布交互模式</h2>
                    <select v-model="clickMode" class="w-full border rounded p-2 text-sm bg-gray-50 mb-1 font-bold text-indigo-700 focus:ring-2 focus:ring-indigo-500 outline-none">
                        <option value="none">鼠标仅查看信息 / 右键座位管理</option>
                        <option value="toggle_aisle">🟩 点击挖空设为走廊</option>
                        <option value="toggle_disabled">🚫 点击设为禁用废座</option>
                        <option value="draw_region">🖌️ 框选色彩区域 (用于分组排座)</option>
                        <option value="define_sequence">🔢 自定义区域排座次序</option>
                    </select>
                    <p class="text-[9px] text-red-500 mb-3">🔥 技巧：鼠标<b>右键点击</b>外圈行/列号，可插入、删除、设为走廊。</p>

                    <div v-if="clickMode === 'draw_region'" class="mb-4 border p-3 rounded-lg bg-gray-50 shadow-inner">
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
                        <p class="text-[9px] text-gray-400 mt-2">包含标题行: <span class="font-bold">序号, 姓名, 单位, 类别, 电话, 备注等</span>。</p>
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
                        <span class="text-[10px] text-gray-500 font-bold flex-1" v-if="!attendeeSearch">全选 / 拖拽项可改变排序同步至该排座批次</span>
                        <span class="text-[10px] text-blue-600 font-bold flex-1" v-else>全选当前搜索出的结果</span>
                    </div>

                    <div class="flex-1 overflow-y-auto border rounded bg-white shadow-inner">
                        <div v-for="(att, index) in filteredAttendees" :key="att.id" 
                             class="p-2 border-b text-sm flex items-center hover:bg-indigo-50 group transition"
                             :draggable="!attendeeSearch" 
                             @dragstart="onDragStart($event, index)" 
                             @dragover.prevent="onDragOver($event, index)" 
                             @dragleave="onDragLeave($event)" 
                             @drop="onDrop($event, index)"
                             :class="{'drag-over': dragTargetIndex === index}">
                            
                            <input type="checkbox" v-model="selectedAttIds" :value="att.id" class="mr-2 w-3.5 h-3.5 text-indigo-600 cursor-pointer">
                            <div class="w-6 text-[10px] text-gray-400 font-bold text-center mr-1">{{index + 1}}</div>
                            
                            <div class="flex-1 truncate">
                                <div class="font-bold text-gray-800">{{att.name}} <span class="text-[9px] bg-gray-100 border px-1 rounded text-gray-500 ml-1 font-normal">{{att.category}}</span></div>
                                <div class="text-[10px] text-gray-500 mt-0.5 truncate">{{att.unit}}</div>
                            </div>
                            <div class="opacity-0 group-hover:opacity-100 flex flex-col space-y-1 ml-1">
                                <button @click="editAttendee(att)" class="text-blue-500 hover:bg-blue-100 rounded px-1"><i class="fa-solid fa-pen"></i></button>
                            </div>
                        </div>
                        <div v-if="filteredAttendees.length===0" class="p-4 text-center text-xs text-gray-400">暂无人员</div>
                    </div>
                </div>

                <!-- Tab 3: 排座管理 -->
                <div v-show="activeTab === 'assign'" class="p-5 flex flex-col h-full overflow-y-auto">
                    
                    <div class="bg-blue-50 border border-blue-200 p-3 rounded-lg mb-4 shadow-sm">
                        <p class="text-[10px] text-blue-800 font-bold mb-1"><i class="fa-solid fa-lightbulb"></i> 标准排座四步曲：</p>
                        <ol class="text-[9px] text-blue-700 list-decimal pl-4 space-y-0.5">
                            <li>先选<b>主席台</b>区块填入核心领导；</li>
                            <li>鼠标<b>右键图上座位</b>强制绑定零星大咖；</li>
                            <li>选定<b>色彩区域</b>，填入所属分类人员；</li>
                            <li>选定<b>主会场</b>，一键填补其余所有人。</li>
                        </ol>
                    </div>

                    <div class="mb-4">
                        <label class="text-xs font-bold text-gray-700 uppercase block mb-1">选目标区块</label>
                        <select v-model="assignConfig.targetZone" class="w-full border rounded p-2 bg-white text-sm font-bold text-indigo-700 shadow-sm focus:ring-1">
                            <option value="podium">⭐ 主席台 (严格按图上排/列号顺排)</option>
                            <optgroup label="主会场特定区:">
                                <option v-for="reg in layout.regionTypes" :key="reg.id" :value="'reg_' + reg.id">区域: {{reg.name}}</option>
                            </optgroup>
                            <option value="all_m">主会场全局 (跳过已被排好的功能区)</option>
                        </select>
                    </div>

                    <div class="mb-4" v-if="assignConfig.targetZone !== 'podium'">
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

                    <div class="mb-3 flex-1 flex flex-col min-h-[220px]">
                        <div class="flex justify-between items-center mb-1">
                            <label class="text-xs font-bold text-gray-700 uppercase">选人入座 (选了 {{assignSelectedAtts.length}} 人)</label>
                            <button @click="selectAllUnassigned" class="text-[10px] bg-gray-200 px-2 rounded hover:bg-gray-300">全选未排</button>
                        </div>
                        <div class="flex flex-col space-y-1 mb-1">
                            <input type="text" v-model="assignSearch" placeholder="搜索姓名/单位..." class="w-full border rounded px-2 py-1 text-sm outline-none focus:ring-1 focus:ring-indigo-500">
                            <div class="flex space-x-1">
                                <input type="text" v-model="assignRangeInput" @keyup.enter="applyRangeSelection" placeholder="序号范围(如1-3,5,9-12)" class="flex-1 border rounded px-2 py-1 text-[11px] outline-none focus:ring-1 focus:ring-indigo-500" title="按回车或点击按钮勾选">
                                <button @click="applyRangeSelection" class="bg-indigo-100 text-indigo-700 px-2 rounded text-xs font-bold hover:bg-indigo-200">选取</button>
                            </div>
                        </div>
                        <div class="flex-1 overflow-y-auto border rounded bg-white shadow-inner p-1">
                            <label v-for="att in unassignedFilteredAttendees" :key="att.id" class="flex items-center p-1 hover:bg-indigo-50 rounded cursor-pointer border-b border-gray-50 last:border-0 transition">
                                <input type="checkbox" :value="att" v-model="assignSelectedAtts" class="mr-2 w-3.5 h-3.5 text-indigo-600">
                                <span class="text-[9px] text-gray-400 w-4 text-right mr-1 font-bold">{{ getGlobalIndex(att) }}</span>
                                <div class="text-[11px] truncate flex-1">
                                    <span class="font-bold text-gray-800">{{att.name}}</span>
                                    <span class="text-[9px] text-gray-500 ml-1">{{att.unit}}</span>
                                </div>
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
                                        <div class="font-extrabold text-[13px] truncate leading-tight">{{ seats[`p-${r-1}-${c-1}`].attendee.name }}</div>
                                        <div class="text-[8px] opacity-90 truncate leading-tight mt-0.5" :style="{ backgroundColor: isSeatBgDark('p', r-1, c-1) ? 'rgba(0,0,0,0.1)' : 'rgba(255,255,255,0.3)', padding: '0 2px', borderRadius: '2px', display: 'inline-block', maxWidth: '100%' }">{{ seats[`p-${r-1}-${c-1}`].attendee.unit }}</div>
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
                                        <div class="font-bold text-[13px] truncate leading-tight">{{ seats[`m-${r-1}-${c-1}`].attendee.name }}</div>
                                        <div class="text-[8px] opacity-90 truncate leading-tight mt-0.5" :style="{ backgroundColor: isSeatBgDark('m', r-1, c-1) ? 'rgba(0,0,0,0.1)' : 'rgba(255,255,255,0.3)', padding: '0 2px', borderRadius: '2px', display: 'inline-block', maxWidth: '100%' }">{{ seats[`m-${r-1}-${c-1}`].attendee.unit }}</div>
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
                <div class="bg-white p-6 rounded-xl shadow-2xl w-96 transform transition-all">
                    <h3 class="font-extrabold text-xl mb-4 text-gray-800">{{ modalForm.id ? '✏️ 编辑人员' : '✨ 增加人员' }}</h3>
                    <div class="space-y-3">
                        <div><label class="text-[10px] font-bold text-gray-500">姓名 *</label><input type="text" v-model="modalForm.name" class="w-full border-b-2 border-gray-300 focus:border-indigo-500 bg-gray-50 p-2 text-sm outline-none"></div>
                        <div><label class="text-[10px] font-bold text-gray-500">单位</label><input type="text" v-model="modalForm.unit" class="w-full border-b-2 border-gray-300 focus:border-indigo-500 bg-gray-50 p-2 text-sm outline-none"></div>
                        <div class="grid grid-cols-2 gap-2">
                            <div><label class="text-[10px] font-bold text-gray-500">类别</label><input type="text" v-model="modalForm.category" class="w-full border-b-2 border-gray-300 focus:border-indigo-500 bg-gray-50 p-2 text-sm outline-none"></div>
                            <div><label class="text-[10px] font-bold text-gray-500">电话</label><input type="text" v-model="modalForm.phone" class="w-full border-b-2 border-gray-300 focus:border-indigo-500 bg-gray-50 p-2 text-sm outline-none"></div>
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
                    podium: { rows: 1, cols: 8, rowLabels: [], colLabels: [], color: '#f97316', textColor: '#ffffff' },
                    main: { rows: 10, cols: 20, rowLabels: [], colLabels: [], color: '#22c55e', textColor: '#ffffff' },
                    regionTypes: [
                        { id: '1', name: '核心内场', color: '#fca5a5' },
                        { id: '2', name: '嘉宾区', color: '#fef08a' }
                    ],
                    customSequences: {},
                    sms_template: '(参会通知){姓名}您好，您的座位在 {排号}排 {座号}，请准时入场。',
                    export_template: '{姓名}\\n({单位})'
                });
                const seats = reactive({}); 
                const regionAssignmentRecords = reactive({});
                const attendees = ref([]);

                const clickMode = ref('none');
                const activeRegionId = ref('1');
                const layoutName = ref('');
                const meetingNameInput = ref('');
                const currentMeetingObj = computed(() => meetings.value.find(m => m.id === currentMeetingId.value) || {});
                
                const attendeeSearch = ref('');
                const selectedAttIds = ref([]);
                const dragTargetIndex = ref(-1);

                const assignSearch = ref('');
                const assignConfig = reactive({ rule: 'center_lr', targetZone: 'all_m' });
                const assignSelectedAtts = ref([]);
                const assignRangeInput = ref('');

                const showModal = ref(false);
                const modalForm = reactive({id: null, name: '', unit: '', category: '', phone: ''});

                const seatMenu = reactive({ show: false, x: 0, y: 0, zone: '', r: -1, c: -1, hasAttendee: false });
                const seatMenuSearch = ref('');
                const axisMenu = reactive({ show: false, x: 0, y: 0, zone: '', type: '', index: -1 });

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
                    const s = seats[`${zone}-${r}-${c}`];
                    if(!s) return false;
                    let hex = '#f9fafb';
                    if(s.attendee) {
                        hex = zone === 'p' ? layout.podium.color : layout.main.color;
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
                    
                    if(!layout.podium.color) layout.podium.color = '#f97316';
                    if(!layout.podium.textColor) layout.podium.textColor = '#ffffff';
                    if(!layout.main.color) layout.main.color = '#22c55e';
                    if(!layout.main.textColor) layout.main.textColor = '#ffffff';
                    if(!layout.export_template) layout.export_template = '{姓名}\\n({单位})';
                    if(!layout.customSequences) layout.customSequences = {};

                    for(let k in seats) delete seats[k];
                    const st = data.seats ? JSON.parse(data.seats) : null;
                    if(st) Object.assign(seats, st);
                    
                    for(let k in regionAssignmentRecords) delete regionAssignmentRecords[k];
                    const rec = data.records ? JSON.parse(data.records) : null;
                    if(rec) Object.assign(regionAssignmentRecords, rec);
                    
                    history.value = [];
                    initSeats(false); 
                    fetchAttendees();
                };

                const fetchAttendees = async () => {
                    if(!currentMeetingId.value) return;
                    const res = await fetch(`/api/meetings/${currentMeetingId.value}/attendees`);
                    attendees.value = await res.json();
                };

                const getGlobalIndex = (att) => {
                    return attendees.value.findIndex(a => a.id === att.id) + 1;
                };

                onMounted(() => loadGlobalData());

                const initSeats = (overwriteLabels = true) => {
                    while(layout.main.rowLabels.length < layout.main.rows) layout.main.rowLabels.push(String(layout.main.rowLabels.length + 1));
                    while(layout.main.colLabels.length < layout.main.cols) layout.main.colLabels.push(String(layout.main.colLabels.length + 1));
                    layout.main.rowLabels.length = layout.main.rows; layout.main.colLabels.length = layout.main.cols;

                    if (overwriteLabels) {
                        let pRows = [];
                        for(let i=layout.podium.rows; i>0; i--) pRows.push(String(i));
                        layout.podium.rowLabels = pRows;
                    }
                    while(layout.podium.colLabels.length < layout.podium.cols) layout.podium.colLabels.push(String(layout.podium.colLabels.length + 1));
                    layout.podium.colLabels.length = layout.podium.cols;

                    for(let k in seats) {
                        let p = k.split('-'); let z = p[0], r = parseInt(p[1]), c = parseInt(p[2]);
                        if(z==='p' && (r >= layout.podium.rows || c >= layout.podium.cols)) delete seats[k];
                        if(z==='m' && (r >= layout.main.rows || c >= layout.main.cols)) delete seats[k];
                    }

                    for(let r=0; r<layout.podium.rows; r++) {
                        for(let c=0; c<layout.podium.cols; c++) {
                            if(!seats[`p-${r}-${c}`]) seats[`p-${r}-${c}`] = { type: 'normal', attendee: null, regionId: null, customLabel: null };
                        }
                    }
                    for(let r=0; r<layout.main.rows; r++) {
                        for(let c=0; c<layout.main.cols; c++) {
                            if(!seats[`m-${r}-${c}`]) seats[`m-${r}-${c}`] = { type: 'normal', attendee: null, regionId: null, customLabel: null };
                        }
                    }
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
                    const key = `${zone}-${r}-${c}`;
                    const s = seats[key];
                    if(s && s.customLabel !== undefined && s.customLabel !== null) return s.customLabel;
                    const obj = zone === 'p' ? layout.podium : layout.main;
                    return formatColLabel(obj.colLabels[c]);
                };

                const setCustomLabel = (key, val) => {
                    saveHistory();
                    if(!val || val.trim() === '') seats[key].customLabel = null;
                    else seats[key].customLabel = val;
                    autoSave();
                };

                const isDuplicateSeat = (zone, r, c) => {
                    const s = seats[`${zone}-${r}-${c}`];
                    if (!s || s.type !== 'normal') return false;
                    const myLabel = getSeatColLabelUI(zone, r, c);
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
                                for (let c = 0; c < obj.cols; c++) seats[`${z}-${r}-${c}`] = seats[`${z}-${r-1}-${c}`] || {type:'normal', attendee:null, regionId:null};
                            }
                            for (let c = 0; c < obj.cols; c++) seats[`${z}-${newIdx}-${c}`] = {type:'normal', attendee:null, regionId:null};
                        } else {
                            obj.cols++; obj.colLabels.splice(newIdx, 0, '新列');
                            for (let c = obj.cols - 1; c > newIdx; c--) {
                                for (let r = 0; r < obj.rows; r++) seats[`${z}-${r}-${c}`] = seats[`${z}-${r}-${c-1}`] || {type:'normal', attendee:null, regionId:null};
                            }
                            for (let r = 0; r < obj.rows; r++) seats[`${z}-${r}-${newIdx}`] = {type:'normal', attendee:null, regionId:null};
                        }
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
                            if(seats[k]) { seats[k].type = status; seats[k].attendee = null; seats[k].regionId = null; seats[k].customLabel = null; }
                        }
                    } else {
                        for(let r=0; r<obj.rows; r++) {
                            const k = `${axisMenu.zone}-${r}-${axisMenu.index}`;
                            if(seats[k]) { seats[k].type = status; seats[k].attendee = null; seats[k].regionId = null; seats[k].customLabel = null; }
                        }
                    }
                    closeContextMenus();
                };

                const onDragStart = (e, idx) => { e.dataTransfer.setData('sourceIdx', idx); };
                const onDragOver = (e, idx) => { dragTargetIndex.value = idx; };
                const onDragLeave = (e) => { dragTargetIndex.value = -1; };
                const onDrop = async (e, targetIdx) => {
                    dragTargetIndex.value = -1;
                    const sourceIdx = parseInt(e.dataTransfer.getData('sourceIdx'));
                    if(isNaN(sourceIdx) || sourceIdx === targetIdx) return;
                    
                    saveHistory();
                    const item = attendees.value.splice(sourceIdx, 1)[0];
                    attendees.value.splice(targetIdx, 0, item);
                    
                    const ids = attendees.value.map(a=>a.id);
                    await fetch(`/api/meetings/${currentMeetingId.value}/attendees/reorder`, {
                        method: 'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({ids})
                    });
                    syncSeatingOrder(); 
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
                        const gIdx = getGlobalIndex(att);
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
                            const parts = foundKey.split('-');
                            removeAndShift(parts[0], parseInt(parts[1]), parseInt(parts[2]), false); 
                        }
                    });
                    
                    await fetch(`/api/meetings/${currentMeetingId.value}/attendees/bulk_delete`, {
                        method: 'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({ids: selectedAttIds.value})
                    });
                    selectedAttIds.value = [];
                    fetchAttendees();
                };
                
                const openAddModal = () => { Object.assign(modalForm, {id: null, name: '', unit: '', category: '', phone: ''}); showModal.value = true; };
                const editAttendee = (att) => { Object.assign(modalForm, att); showModal.value = true; };
                const saveModalForm = async () => {
                    if(!modalForm.name) return alert('姓名必填');
                    const method = modalForm.id ? 'PUT' : 'POST';
                    await fetch(`/api/meetings/${currentMeetingId.value}/attendees`, { method, headers: {'Content-Type': 'application/json'}, body: JSON.stringify(modalForm) });
                    showModal.value = false;
                    fetchAttendees();
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
                
                const seatClass = (zone, r, c) => {
                    const s = seats[`${zone}-${r}-${c}`];
                    if(!s) return '';
                    if(s.type === 'aisle') return 'seat-aisle';
                    if(s.type === 'disabled') return 'seat-disabled';
                    if(s.attendee) return 'seat-occupied';
                    return '';
                };

                const seatStyle = (zone, r, c) => {
                    const s = seats[`${zone}-${r}-${c}`];
                    if(!s || s.type !== 'normal' || !s.attendee) return {};
                    const conf = zone === 'p' ? layout.podium : layout.main;
                    return { backgroundColor: conf.color, borderColor: conf.color, color: conf.textColor };
                };

                const closeContextMenus = () => { seatMenu.show = false; axisMenu.show = false; };
                const handleSeatRightClick = (e, zone, r, c) => {
                    const key = `${zone}-${r}-${c}`;
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

                const forceAssign = (zone, r, c, att) => { saveHistory(); seats[`${zone}-${r}-${c}`].attendee = att; closeContextMenus(); };
                const clearSeat = (zone, r, c) => { saveHistory(); seats[`${zone}-${r}-${c}`].attendee = null; closeContextMenus(); };
                
                const removeAndShift = (zone, r, c, pushHistory=true) => {
                    const key = `${zone}-${r}-${c}`;
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
                    const key = `${zone}-${r}-${c}`;
                    if(clickMode.value === 'toggle_aisle') { saveHistory(); seats[key].type = seats[key].type==='aisle'?'normal':'aisle'; seats[key].attendee=null; seats[key].customLabel=null; }
                    else if(clickMode.value === 'toggle_disabled') { saveHistory(); seats[key].type = seats[key].type==='disabled'?'normal':'disabled'; seats[key].attendee=null; seats[key].customLabel=null; }
                    else if(clickMode.value === 'draw_region') {
                        if(zone === 'p' || seats[key].type !== 'normal') return;
                        saveHistory(); seats[key].regionId = seats[key].regionId === activeRegionId.value ? null : activeRegionId.value;
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
                    for(let k in seats) {
                        const p = k.split('-');
                        if(z==='podium' && p[0]==='p') keysToClear.push(k);
                        else if(z==='all_m' && p[0]==='m' && !seats[k].regionId) keysToClear.push(k);
                        else if(z.startsWith('reg_') && seats[k].regionId === z.split('_')[1]) keysToClear.push(k);
                    }
                    
                    keysToClear.forEach(k => {
                        seats[k].attendee = null;
                        for(let bKey in regionAssignmentRecords) {
                            const idx = regionAssignmentRecords[bKey].indexOf(k);
                            if(idx > -1) regionAssignmentRecords[bKey].splice(idx, 1);
                        }
                    });
                };

                const executeAssignment = () => {
                    if(assignSelectedAtts.value.length === 0) return alert('未勾选要排的人员');
                    saveHistory();
                    
                    let targetCoords = [];
                    for(let key in seats) {
                        const p = key.split('-'); const z = p[0]; const r = parseInt(p[1]); const c = parseInt(p[2]);
                        if(seats[key].type !== 'normal' || seats[key].attendee) continue;
                        
                        if(assignConfig.targetZone === 'podium' && z === 'p') targetCoords.push({z, r, c});
                        else if(assignConfig.targetZone === 'all_m' && z === 'm' && !seats[key].regionId) targetCoords.push({z, r, c});
                        else if(assignConfig.targetZone.startsWith('reg_') && seats[key].regionId === assignConfig.targetZone.split('_')[1]) targetCoords.push({z, r, c});
                    }

                    if(targetCoords.length < assignSelectedAtts.value.length) {
                        if(!confirm(`空位仅剩 ${targetCoords.length} 个，人员有 ${assignSelectedAtts.value.length}，超出部分将抛弃，继续吗？`)) { history.value.pop(); return; }
                    }

                    let orderedCoords = [];
                    if(assignConfig.targetZone === 'podium') {
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
                        const k = `${pos.z}-${pos.r}-${pos.c}`;
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
                    const config = { podium: layout.podium, main: layout.main, regionTypes: layout.regionTypes, typeMap: {}, customSequences: layout.customSequences };
                    for(let k in seats) { if(seats[k].type!=='normal' || seats[k].regionId) config.typeMap[k] = {t:seats[k].type, r:seats[k].regionId}; }
                    await fetch('/api/templates', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({name:layoutName.value, config})});
                    loadGlobalData(); alert('模板保存成功');
                };
                const loadTemplate = (id) => {
                    if(!id) return;
                    const t = templates.value.find(x=>x.id==id);
                    if(!confirm('应用模板将重置本场会议的排位格局，是否继续？')) return;
                    const c = t.config;
                    Object.assign(layout.podium, c.podium); Object.assign(layout.main, c.main); layout.regionTypes = c.regionTypes;
                    if(c.customSequences) layout.customSequences = c.customSequences;
                    else layout.customSequences = {};
                    history.value = [];
                    initSeats();
                    for(let k in c.typeMap) { if(seats[k]) { seats[k].type = c.typeMap[k].t; seats[k].regionId = c.typeMap[k].r; } }
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
                        podium: layout.podium, main: layout.main, seats_data: exportSeats, 
                        sms_template: layout.sms_template, export_template: layout.export_template.replace(/\\n/g, '\n')
                    };
                    const res = await fetch('/api/export', { method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(payload) });
                    const blob = await res.blob();
                    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `${payload.conf_name}排座导出.xlsx`; a.click();
                };

                const filteredAttendees = computed(() => {
                    if(!attendeeSearch.value) return attendees.value;
                    const q = attendeeSearch.value.toLowerCase();
                    return attendees.value.filter(a => a.name.includes(q) || (a.unit && a.unit.includes(q)));
                });
                const unassignedAttendees = computed(() => {
                    const assignedIds = Object.values(seats).filter(s => s && s.attendee).map(s => s.attendee.id);
                    return attendees.value.filter(a => !assignedIds.includes(a.id));
                });
                const unassignedFilteredAttendees = computed(() => {
                    if(!assignSearch.value) return unassignedAttendees.value;
                    return unassignedAttendees.value.filter(a => a.name.includes(assignSearch.value) || (a.unit && a.unit.includes(assignSearch.value)));
                });
                const seatMenuFilteredAtts = computed(() => unassignedAttendees.value.filter(a => a.name.includes(seatMenuSearch.value)).slice(0, 40));
                
                const seatMenuLabel = computed(() => {
                    if(seatMenu.zone === '') return '';
                    const obj = seatMenu.zone === 'p' ? layout.podium : layout.main;
                    const key = `${seatMenu.zone}-${seatMenu.r}-${seatMenu.c}`;
                    const customLabel = seats[key]?.customLabel;
                    const finalLabel = (customLabel !== undefined && customLabel !== null) ? customLabel : formatColLabel(obj.colLabels[seatMenu.c]);
                    return `${obj.rowLabels[seatMenu.r]}排 ${finalLabel}`;
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
                    clickMode, activeRegionId, layoutName, attendeeSearch, selectedAttIds, dragTargetIndex,
                    assignSearch, assignRangeInput, assignConfig, assignSelectedAtts, showModal, modalForm,
                    seatMenu, seatMenuSearch, seatMenuFilteredAtts, seatMenuLabel, axisMenu, extraKeys,
                    filteredAttendees, unassignedFilteredAttendees, isAllSelected, toggleSelectAllFiltered,
                    loadMeetingState, createMeeting, deleteCurrentMeeting, renameMeeting, initSeats, handleLabelChange, syncPodiumLabels, formatColLabel, getSeatColLabelUI, setCustomLabel, isDuplicateSeat, isSeatBgDark, clearCustomSequence, handleSeatClick, 
                    addRegionType, removeRegionType, getSeatRegionColor, seatClass, seatStyle, modifyAxis,
                    onDragStart, onDragOver, onDragLeave, onDrop, applyRangeSelection, getGlobalIndex,
                    uploadFile, bulkDeleteAttendees, openAddModal, editAttendee, saveModalForm,
                    handleSeatRightClick, handleAxisRightClick, closeContextMenus, setAxisStatus, forceAssign, clearSeat, removeAndShift,
                    clearTargetZone, executeAssignment, selectAllUnassigned, undo,
                    saveTemplate, loadTemplate, deleteTemplate, insertVar, insertExportVar, exportExcel, autoSave
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