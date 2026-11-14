import os
import zipfile
import re
import datetime
import tempfile
import shutil
from flask import Flask, request, send_file, render_template_string
import openpyxl
import exifread

app = Flask(__name__)

HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <title>📸 写真データ自動記録ツール</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; margin: 0; padding: 20px; background: #f0f2f5; color: #333; }
        .container { background: white; padding: 30px; border-radius: 12px; max-width: 600px; margin: 40px auto; box-shadow: 0 4px 12px rgba(0,0,0,0.05); }
        h1 { font-size: 1.5rem; margin-top: 0; text-align: center; color: #1a1a1a; }
        .description { font-size: 0.95rem; color: #666; margin-bottom: 25px; text-align: center; }
        .form-group { margin-bottom: 20px; padding: 15px; background: #f8f9fa; border-radius: 8px; border: 1px solid #e9ecef; }
        label { display: block; margin-bottom: 8px; font-weight: 600; font-size: 0.95rem; }
        input[type="file"] { width: 100%; }
        button { width: 100%; background: #0070f3; color: white; border: none; padding: 14px; font-size: 1rem; font-weight: bold; border-radius: 8px; cursor: pointer; transition: background 0.2s; margin-top: 10px; }
        button:hover { background: #005bb5; }
        .footer { text-align: center; margin-top: 20px; font-size: 0.8rem; color: #999; }
    </style>
</head>
<body>
    <div class="container">
        <h1>📸 写真データ自動記録</h1>
        <p class="description">エクセルテンプレートと写真を選択すると、<br>メタデータを自動入力したファイルがダウンロードされます。</p>
        <form action="/process" method="post" enctype="multipart/form-data">
            <div class="form-group">
                <label>1. エクセルテンプレート (.xlsx)</label>
                <input type="file" name="excel_file" accept=".xlsx" required>
            </div>
            <div class="form-group">
                <label>2. 写真データ (最大2枚選択)</label>
                <input type="file" name="photos" accept="image/*,.cr2,.cr3,.arw,.nef" multiple required>
            </div>
            <button type="submit">自動入力してダウンロード</button>
        </form>
    </div>
    <div class="footer">Powered by Vercel & Flask</div>
</body>
</html>
'''

def restore_excel_validations(orig_path, gen_path, target_sheet_name="データ票"):
    try:
        sheet_xml_path = None
        with zipfile.ZipFile(orig_path, 'r') as z:
            wb_xml = z.read('xl/workbook.xml').decode('utf-8')
            m = re.search(f'<sheet [^>]*name="{target_sheet_name}"[^>]*r:id="([^"]+)"', wb_xml)
            if not m: return
            rid = m.group(1)
            
            rels_xml = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
            m2 = re.search(f'<Relationship [^>]*Id="{rid}"[^>]*Target="([^"]+)"', rels_xml)
            if not m2: return
            sheet_xml_path = "xl/" + m2.group(1).lstrip("/")
            
            orig_sheet_xml = z.read(sheet_xml_path).decode('utf-8')
            
        dv_match = re.search(r'<dataValidations[\s>].*?</dataValidations>', orig_sheet_xml, re.DOTALL)
        extlst_match = re.search(r'<extLst>.*?</extLst>', orig_sheet_xml, re.DOTALL)
        
        if not dv_match and not extlst_match: return 
            
        temp_gen = str(gen_path) + ".tmp"
        with zipfile.ZipFile(gen_path, 'r') as zin, zipfile.ZipFile(temp_gen, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                content = zin.read(item.filename)
                if item.filename == sheet_xml_path:
                    gen_sheet_xml = content.decode('utf-8')
                    gen_sheet_xml = re.sub(r'<dataValidations[\s>].*?</dataValidations>', '', gen_sheet_xml, flags=re.DOTALL)
                    gen_sheet_xml = re.sub(r'<extLst>.*?</extLst>', '', gen_sheet_xml, flags=re.DOTALL)
                    insert_str = (dv_match.group(0) if dv_match else "") + (extlst_match.group(0) if extlst_match else "")
                    gen_sheet_xml = gen_sheet_xml.replace('</worksheet>', f'{insert_str}</worksheet>')
                    content = gen_sheet_xml.encode('utf-8')
                zout.writestr(item, content)
        os.replace(temp_gen, gen_path)
    except Exception as e:
        print("Validation restore error:", e)

class ExifExtractor:
    @classmethod
    def extract(cls, file_stream, filename):
        res = {"valid": False, "FileName": filename}
        try:
            tags = exifread.process_file(file_stream, details=False)
            if not tags: return res

            def get_str(key):
                return str(tags[key].values) if key in tags else None

            def get_ratio_float(key):
                if key in tags and tags[key].values:
                    val = tags[key].values[0]
                    if hasattr(val, 'num') and hasattr(val, 'den') and val.den != 0:
                        return val.num / val.den
                return None

            dt_str = get_str('EXIF DateTimeOriginal') or get_str('Image DateTime')
            if dt_str: res["DateTimeOriginal"] = str(dt_str)[:19]

            make = get_str('Image Make')
            model = get_str('Image Model')
            if model:
                name = str(model).replace("Canon ", "").replace("Canon", "").strip()
                name = name.replace(" ", "　").replace("MK-2", "MK-Ⅱ").replace("MK-3", "MK-Ⅲ").replace("MK-4", "MK-Ⅳ").replace("MK-II", "MK-Ⅱ")
                res["Camera"] = name

            res["LensModel"] = get_str('EXIF LensModel')
            
            fn_val = get_ratio_float('EXIF FNumber')
            if fn_val:
                if fn_val >= 10 and fn_val.is_integer(): res["FNumber"] = f"F {int(fn_val)}"
                else: res["FNumber"] = f"F {fn_val:.1f}".rstrip('0').rstrip('.')
            
            ss_val = get_ratio_float('EXIF ExposureTime')
            if ss_val:
                if ss_val < 1.0: res["ShutterSpeed"] = f"1/{int(round(1/ss_val))}"
                else: res["ShutterSpeed"] = str(ss_val)
                
            iso_val = tags.get('EXIF ISOSpeedRatings')
            res["ISO"] = int(iso_val.values[0]) if iso_val and iso_val.values else None
            
            fl_val = get_ratio_float('EXIF FocalLength')
            res["FocalLength_mm"] = int(fl_val) if fl_val else None
            res["valid"] = True
        except Exception as e:
            pass
        return res

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/process', methods=['POST'])
def process():
    if 'excel_file' not in request.files or 'photos' not in request.files:
        return "ファイルが不足しています", 400
        
    excel_file = request.files['excel_file']
    photos = request.files.getlist('photos')
    
    if not excel_file.filename.endswith('.xlsx'):
        return "エクセルファイル(.xlsx)を選択してください", 400
        
    if len(photos) > 2: photos = photos[:2]
        
    with tempfile.TemporaryDirectory() as temp_dir:
        orig_excel_path = os.path.join(temp_dir, "template.xlsx")
        excel_file.save(orig_excel_path)
        
        items_to_add = []
        for p in photos:
            if not p.filename: continue
            file_data = p.read()
            import io
            exif_info = ExifExtractor.extract(io.BytesIO(file_data), p.filename)
            if exif_info["valid"]:
                items_to_add.append(exif_info)
                
        if not items_to_add:
            return "有効な写真データ（EXIF）が見つかりませんでした", 400
            
        new_excel_path = os.path.join(temp_dir, "output.xlsx")
        shutil.copy2(orig_excel_path, new_excel_path)
        
        wb = openpyxl.load_workbook(new_excel_path)
        ws = wb["データ票"]
        
        BLOCK_SIZE = 20
        weekdays = ["月", "火", "水", "木", "金", "土", "日"]
        
        for idx, data in enumerate(items_to_add):
            start_row = (idx * BLOCK_SIZE) + 1  
            
            dto_str = data.get("DateTimeOriginal")
            if dto_str:
                try:
                    dt = datetime.datetime.strptime(dto_str, "%Y:%m:%d %H:%M:%S")
                    ws.cell(start_row + 5, 4, value=dt.year)
                    ws.cell(start_row + 5, 9, value=dt.month)
                    ws.cell(start_row + 5, 12, value=dt.day)
                    ws.cell(start_row + 5, 16, value=weekdays[dt.weekday()])
                    ws.cell(start_row + 5, 20, value=dt.hour)
                    ws.cell(start_row + 5, 23, value=f"{dt.minute:02d}")
                except: pass

            ws.cell(start_row + 6, 4, value=data.get("Camera"))
            ws.cell(start_row + 6, 17, value=data.get("LensModel"))
            ws.cell(start_row + 7, 21, value=data.get("FocalLength_mm"))
            ws.cell(start_row + 8, 4, value=data.get("FNumber"))
            ws.cell(start_row + 8, 21, value=data.get("ShutterSpeed"))
            ws.cell(start_row + 9, 4, value=data.get("ISO"))
            
        wb.save(new_excel_path)
        wb.close()
        
        restore_excel_validations(orig_excel_path, new_excel_path, "データ票")
        
        out_filename = f"自動入力_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return send_file(new_excel_path, as_attachment=True, download_name=out_filename)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8000)
