# ==========================================
# 【究極完成版】写真作品データ Excel自動記録システム
# (全メタデータリスト自動適合 ＋ プルダウン完全保護機能搭載)
# ==========================================

import os
import json
import hashlib
import shutil
import datetime
import unicodedata
import subprocess
import zipfile
import re
from pathlib import Path
import openpyxl

def restore_excel_validations(orig_path, gen_path, target_sheet_name):
    """
    openpyxlが保存時に破壊してしまうExcelのプルダウンリスト（データ入力規則）を、
    原本のExcelファイルの裏側（XML）から抽出し、新しいファイルへ完璧に復元するハック関数。
    """
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
        
        if not dv_match and not extlst_match:
            return 
            
        temp_gen = str(gen_path) + ".tmp"
        with zipfile.ZipFile(gen_path, 'r') as zin, zipfile.ZipFile(temp_gen, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                content = zin.read(item.filename)
                if item.filename == sheet_xml_path:
                    gen_sheet_xml = content.decode('utf-8')
                    gen_sheet_xml = re.sub(r'<dataValidations[\s>].*?</dataValidations>', '', gen_sheet_xml, flags=re.DOTALL)
                    gen_sheet_xml = re.sub(r'<extLst>.*?</extLst>', '', gen_sheet_xml, flags=re.DOTALL)
                    
                    insert_str = ""
                    if dv_match: insert_str += dv_match.group(0)
                    if extlst_match: insert_str += extlst_match.group(0)
                    
                    gen_sheet_xml = gen_sheet_xml.replace('</worksheet>', f'{insert_str}</worksheet>')
                    content = gen_sheet_xml.encode('utf-8')
                zout.writestr(item, content)
        
        os.replace(temp_gen, gen_path)
    except Exception as e:
        print(f"⚠️ プルダウン復元中にエラーが発生しましたが、データ入力は完了しています: {e}")

def main():
    # --- 0. 拡張ツール(ExifTool)の自動インストール ---
    print("⚙️ システムの初期設定を行っています（約10〜20秒）...")
    os.system('apt-get update -qq')
    os.system('apt-get install -y -qq libimage-exiftool-perl')

    # --- 1. Google Driveの接続 ---
    try:
        from google.colab import drive
        print("☁️ Googleドライブに接続しています...")
        drive.mount('/content/drive', force_remount=True)
    except ImportError:
        print("⚠️ Colab環境ではないため、ローカルで実行します。")

    drive_root = Path('/content/drive/MyDrive') if Path('/content/drive/MyDrive').exists() else Path('.')

    # --- 2. 作業フォルダとExcelファイルの特定 ---
    print("\n🔍 作業フォルダとExcelファイルを探査中...")
    target_dir = None
    if drive_root.name != '.':
        for item in drive_root.iterdir():
            item_norm = unicodedata.normalize('NFC', item.name)
            if '写真データ自動記録' in item_norm or ('写真' in item_norm and '自動記録' in item_norm):
                target_dir = item
                break
    if not target_dir: target_dir = drive_root
    
    BASE_DIR = target_dir
    PHOTO_DIR = BASE_DIR / 'photos'
    SIDECAR_FILE = BASE_DIR / 'processed_files.json'

    excel_file = None
    search_locations = [Path('/content'), BASE_DIR] if Path('/content').exists() else [BASE_DIR]
    
    for loc in search_locations:
        if loc.exists():
            for item in loc.iterdir():
                if item.is_file() and item.name.lower().endswith('.xlsx') and not item.name.startswith('~$') and not item.name.startswith('temp_'):
                    excel_file = item
                    break
        if excel_file: break

    if not excel_file:
        print("❌ Excelファイルが見つかりません。")
        return

    EXCEL_FILE = excel_file
    print(f"✅ 対象Excelファイル: {EXCEL_FILE.name}")

    # --- 3. EXIFデータ抽出＆Excelリスト完全適合エンジン ---
    class ExifExtractor:
        @classmethod
        def format_f_number(cls, f_val):
            try:
                f_float = float(f_val)
                if f_float >= 10 and f_float.is_integer(): return f"F {int(f_float)}"
                return f"F {f_float:.1f}".rstrip('0').rstrip('.')
            except: return str(f_val) if f_val else None

        @classmethod
        def format_shutter_speed(cls, ss_val):
            try:
                ss_float = float(ss_val)
                if ss_float < 1.0: return f"1/{int(round(1/ss_float))}"
                return str(ss_val)
            except: return str(ss_val) if ss_val else None

        @classmethod
        def format_camera_name(cls, make, model):
            if not model: return None
            name = str(model).replace("Canon ", "").replace("Canon", "").strip()
            name = name.replace(" ", "　")
            name = name.replace("MK-2", "MK-Ⅱ").replace("MK-3", "MK-Ⅲ").replace("MK-4", "MK-Ⅳ").replace("MK-II", "MK-Ⅱ")
            return name

        @classmethod
        def format_exposure_program(cls, prog):
            mapping = {
                "Aperture-priority AE": "Av：絞り優先AE",
                "Aperture Priority": "Av：絞り優先AE",
                "Shutter speed priority AE": "Tv：シャッター優先AE",
                "Shutter Priority": "Tv：シャッター優先AE",
                "Manual": "M：マニュアル",
                "Program AE": "P：プログラムAE",
                "Bulb": "バルブ撮影"
            }
            return mapping.get(str(prog), str(prog) if prog else None)

        @classmethod
        def format_exposure_comp(cls, ec):
            if ec is None or ec == "" or ec == 0: return "0EV"
            try:
                val = float(ec)
                if val == 0: return "0EV"
                sign = "+" if val > 0 else "-"
                abs_val = abs(val)
                if abs_val in [0.333, 0.33, 0.34]: return f"{sign}1/3EV"
                if abs_val in [0.666, 0.66, 0.67]: return f"{sign}2/3EV"
                if abs_val == 0.5: return f"{sign}1/2EV"
                if abs_val.is_integer(): return f"{sign}{int(abs_val)}EV"
                return f"{sign}{abs_val:.1f}EV"
            except: return str(ec)

        @classmethod
        def format_white_balance(cls, wb):
            wb_str = str(wb).lower()
            if "auto" in wb_str: return "AWB（ﾎﾜｲﾄ優先）"
            if "daylight" in wb_str or "sun" in wb_str: return "太陽光"
            if "shade" in wb_str: return "日陰"
            if "cloud" in wb_str: return "曇り"
            if "tungsten" in wb_str or "incandescent" in wb_str: return "白熱電球"
            if "fluorescent" in wb_str: return "蛍光灯"
            if "flash" in wb_str: return "ストロボ"
            return "AWB（ﾎﾜｲﾄ優先）"

        @classmethod
        def extract(cls, file_path: str) -> dict:
            path = Path(file_path)
            res = {"valid": False, "FileName": path.name}
            try:
                result = subprocess.run(['exiftool', '-json', str(path)], capture_output=True, text=True)
                if result.returncode != 0: raise ValueError("ExifToolエラー")
                
                data = json.loads(result.stdout)[0]
                dt_str = data.get("DateTimeOriginal") or data.get("CreateDate")
                if dt_str: res["DateTimeOriginal"] = str(dt_str)[:19] 

                res["Camera"] = cls.format_camera_name(data.get("Make"), data.get("Model"))
                res["LensModel"] = str(data.get("LensModel") or data.get("Lens", "")).strip() or None
                res["ExposureProgram"] = cls.format_exposure_program(data.get("ExposureProgram"))
                res["ISO"] = int(data.get("ISO")) if data.get("ISO") else None
                res["FNumber"] = cls.format_f_number(data.get("FNumber"))
                res["ShutterSpeed"] = cls.format_shutter_speed(data.get("ExposureTime") or data.get("ShutterSpeed"))
                res["ExposureComp"] = cls.format_exposure_comp(data.get("ExposureCompensation"))
                res["WhiteBalance"] = cls.format_white_balance(data.get("WhiteBalance"))

                fl_val = data.get("FocalLength")
                if fl_val:
                    try: res["FocalLength_mm"] = int(float(str(fl_val).replace('mm', '').strip()))
                    except: res["FocalLength_mm"] = str(fl_val)

                res["valid"] = True
            except:
                pass
            return res

    # --- 4. メイン処理（2枚読み込み＆上下の表へ入力） ---
    def calculate_file_hash(file_path):
        hasher = hashlib.sha256()
        with open(file_path, 'rb') as f:
            while chunk := f.read(8192): hasher.update(chunk)
        return hasher.hexdigest()

    processed_hashes = {}
    if SIDECAR_FILE.exists():
        with open(SIDECAR_FILE, 'r', encoding='utf-8') as f:
            processed_hashes = json.load(f)

    valid_exts = ['.jpg', '.jpeg', '.cr3', '.cr2', '.nef', '.arw', '.dng', '.heic']
    photo_files = []
    for d in [Path('/content'), PHOTO_DIR, BASE_DIR]:
        if d.exists():
            for item in d.iterdir():
                if item.is_file() and item.suffix.lower() in valid_exts:
                    photo_files.append(item)

    items_to_add = []
    for p in sorted(photo_files, key=os.path.getmtime):
        f_hash = calculate_file_hash(str(p))
        if f_hash in processed_hashes: continue
        
        exif_info = ExifExtractor.extract(str(p))
        if exif_info["valid"]:
            items_to_add.append((exif_info, f_hash))
            if len(items_to_add) == 2: break

    if not items_to_add:
        print("\n✅ 新しい写真が見つかりません。（すべて登録済みか、写真がありません）")
        return

    print(f"\n=== Excelへの自動入力処理 ({len(items_to_add)}枚) ===")
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    
    out_dir = BASE_DIR if BASE_DIR.exists() else Path('/content')
    new_excel_path = out_dir / f"{EXCEL_FILE.stem}_自動入力_{timestamp}.xlsx"
    shutil.copy2(EXCEL_FILE, new_excel_path)

    wb = openpyxl.load_workbook(new_excel_path)
    ws = wb["データ票"]

    BLOCK_SIZE = 20
    weekdays = ["月", "火", "水", "木", "金", "土", "日"]

    for idx, (data, f_hash) in enumerate(items_to_add):
        start_row = (idx * BLOCK_SIZE) + 1  

        # 撮影日時の書き込み
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

        # カメラ・レンズ・撮影データの書き込み
        ws.cell(start_row + 6, 4, value=data.get("Camera"))              # カメラ名
        ws.cell(start_row + 6, 17, value=data.get("LensModel"))          # レンズ名
        ws.cell(start_row + 7, 4, value=data.get("ExposureProgram"))     # 撮影モード
        ws.cell(start_row + 7, 21, value=data.get("FocalLength_mm"))     # 焦点距離
        ws.cell(start_row + 8, 4, value=data.get("FNumber"))             # 絞り値
        ws.cell(start_row + 8, 21, value=data.get("ShutterSpeed"))       # シャッタースピード
        ws.cell(start_row + 9, 4, value=data.get("ISO"))                 # ISO感度
        ws.cell(start_row + 9, 21, value=data.get("ExposureComp"))       # 露出補正
        ws.cell(start_row + 10, 4, value=data.get("WhiteBalance"))       # ホワイトバランス

        processed_hashes[f_hash] = {"file_name": data["FileName"], "added_at": datetime.datetime.now().isoformat()}
        print(f"✅ 入力完了: {data['FileName']} -> {'上の表 (ブロック1)' if idx==0 else '下の表 (ブロック2)'}")

    wb.save(new_excel_path)
    wb.close()

    # --- 5. プルダウンリスト（データ入力規則）の完全復元 ---
    print("🔄 Excelのプルダウンメニューを保護・復元中...")
    restore_excel_validations(EXCEL_FILE, new_excel_path, "データ票")

    with open(SIDECAR_FILE if BASE_DIR.exists() else Path('/content/processed_files.json'), 'w', encoding='utf-8') as f:
        json.dump(processed_hashes, f, indent=2, ensure_ascii=False)

    print(f"\n🎉 処理が完了しました！")
    print(f"📁 作成ファイル: 【{new_excel_path.name}】")

if __name__ == '__main__':
    main()
