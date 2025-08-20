# eines_automation_v6_1.py
import argparse, os, re, time, json, pandas as pd, xml.etree.ElementTree as ET
from pathlib import Path
from pywinauto.application import Application
from pywinauto import Desktop
import pyautogui, pytesseract, cv2, numpy as np, psutil

import pytesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0.03

DECIMAL_COMMA = True
STATION_ORDER = ["RIGHTSIDE1","RIGHTSIDE2","LEFTSIDE1","LEFTSIDE2","ROOF","ROOFHOOD","HOOD","BRIDGE"]

def to_float(s):
    if s is None or (isinstance(s, float) and pd.isna(s)): return None
    ss = str(s).strip()
    if DECIMAL_COMMA:
        ss = ss.replace('.', '').replace(',', '.') if ss.count(',')==1 and ss.count('.')>1 else ss.replace(',', '.')
    try: return float(ss)
    except: return None

def load_excel(path):
    xls = pd.ExcelFile(path)
    sheet = "Hoja1" if "Hoja1" in xls.sheet_names else xls.sheet_names[0]
    df_raw = pd.read_excel(path, sheet_name=sheet, header=None, dtype=str)
    header_row = None
    for i in range(min(30, len(df_raw))):
        row_vals = [str(v).strip() for v in df_raw.iloc[i].tolist()]
        if "X_FRAME" in row_vals and "Y_FRAME" in row_vals and "FRAME" in row_vals: header_row = i; break
        if "X_PAG" in row_vals and "X_EINES" in row_vals: header_row = i
    if header_row is None: header_row = 0
    df = pd.read_excel(path, sheet_name=sheet, header=header_row)
    df.columns = [str(c).strip().upper().replace(" ", "_") for c in df.columns]
    for c in ["X_NEW","Y_NEW","Z_NEW","TIME_MS"]:
        if c not in df.columns: df[c] = ""
    return df, sheet, header_row

def load_stations_map(xml_path):
    tree = ET.parse(xml_path); root = tree.getroot()
    station_by_id = {int(st.attrib["id"]): st.findtext("name") for st in root.findall("station")}
    cams_by_station_id = {}
    for st in root.findall("station"):
        sid = int(st.attrib["id"]); cams=[]
        for cam in st.find("cameras").findall("camera"):
            cams.append({"id": int(cam.attrib["id"]), "target": int(cam.findtext("target")), "subset": int(cam.findtext("subset"))})
        cams_by_station_id[sid] = cams
    return station_by_id, cams_by_station_id

def _find_window_by_process(exe_basename, title_hint, timeout=60):
    t0=time.time()
    while time.time()-t0<timeout:
        for w in Desktop(backend="uia").windows():
            try:
                title = w.window_text()
                pid = w.process_id()
                p = psutil.Process(pid)
                if exe_basename.lower() in os.path.basename(p.exe()).lower():
                    if not title_hint or title_hint.lower() in title.lower():
                        return w
            except Exception:
                continue
        time.sleep(0.5)
    return None

def attach_or_launch(app_path, title_hint, work_dir=None, attach_only=False, extra_load_wait=3.0):
    try:
        dlg = Desktop(backend="uia").window(title_re=".*"+re.escape(title_hint)+".*", found_index=0)
        if dlg.exists(): dlg.set_focus(); time.sleep(extra_load_wait); return dlg
    except Exception:
        pass
    byproc = _find_window_by_process(os.path.basename(app_path), title_hint, timeout=5)
    if byproc is not None: byproc.set_focus(); time.sleep(extra_load_wait); return byproc
    if attach_only: raise RuntimeError("Window not found to attach.")
    if work_dir is None: work_dir = os.path.dirname(app_path)
    app = Application(backend="uia").start(cmd_line=f'"{app_path}"', work_dir=work_dir)
    time.sleep(3.0)
    t0=time.time(); dlg=None
    while time.time()-t0<60:
        try:
            dlg = Desktop(backend="uia").window(title_re=".*"+re.escape(title_hint)+".*")
            if dlg.exists(): dlg.wait("visible", timeout=5); break
        except Exception:
            pass
        dlg = _find_window_by_process(os.path.basename(app_path), title_hint, timeout=1)
        if dlg: break
        time.sleep(1.0)
    if dlg is None: raise RuntimeError("Timed out waiting for window")
    dlg.set_focus(); time.sleep(extra_load_wait)
    return dlg

def jiggle():
    x,y = pyautogui.position()
    pyautogui.moveTo(x+2, y+2, duration=0.02)
    pyautogui.moveTo(x, y, duration=0.02)

def click_xy(x,y, double=False):
    jiggle()
    pyautogui.click(x,y)
    if double:
        time.sleep(0.05); pyautogui.click(x,y)

def open_combo_at(x,y):
    click_xy(x,y); time.sleep(0.15)

def ocr_dropdown_region_from_base(x, y, width=340, height=260):
    # capture region below (x,y)
    img = pyautogui.screenshot(region=(x-60, y+15, width, height))
    bgr = cv2.cvtColor(np.array(img), cv2.COLOR_RGBA2BGR)
    gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    thr = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY)[1]
    text = pytesseract.image_to_string(thr, lang="eng", config="--psm 6")
    proj = np.sum(255-thr, axis=1)
    rows=[]; inb=False; start=0
    for i,v in enumerate(proj):
        if v>0 and not inb: inb=True; start=i
        elif v==0 and inb: inb=False; rows.append((start,i))
    bands=[(a+b)//2 for a,b in rows if (b-a)>6]
    return (x-60, y+15, width, height), [ln.strip() for ln in text.splitlines() if ln.strip()], bands

def select_from_dropdown_at(x,y, target_text):
    open_combo_at(x,y)
    region, lines, bands = ocr_dropdown_region_from_base(x,y)
    # match target by exact / contains
    target_text = str(target_text).upper()
    idx=None
    for i,txt in enumerate(lines):
        if txt.upper() == target_text: idx=i; break
    if idx is None:
        for i,txt in enumerate(lines):
            if target_text in txt.upper(): idx=i; break
    if idx is not None and idx < len(bands):
        cx = region[0] + 120
        cy = region[1] + bands[idx]
        click_xy(cx, cy); time.sleep(0.05); pyautogui.press("enter")
        return True
    return False

def put_text_xy(x,y,value):
    s=str(value).strip()
    s = s.replace('.', ',') if DECIMAL_COMMA else s
    click_xy(x,y); pyautogui.hotkey("ctrl","a"); pyautogui.typewrite(s)

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument("--config", default="config_v6_1.json")
    ap.add_argument("--run", action="store_true")
    ap.add_argument("--dry", action="store_true")
    args=ap.parse_args()

    cfg=json.load(open(args.config,"r",encoding="utf-8"))
    excel=cfg["excel_path"]; xmlp=cfg["stations_xml_path"]; app=cfg["viewer_exe_path"]
    title=cfg.get("window_title_contains","EINES System 3D for ESPQi")
    mesh_name=cfg.get("mesh_name","TAYCAN_NACHFOLGER")
    work_dir=cfg.get("work_dir", os.path.dirname(app))
    attach_only=bool(cfg.get("attach_only", False))
    extra_wait=float(cfg.get("extra_load_wait_s", 3.0))

    df,sheet,_=load_excel(excel)
    tree = ET.parse(xmlp); root = tree.getroot()
    # map station id->name
    station_by_id = {int(st.attrib["id"]): st.findtext("name") for st in root.findall("station")}

    if args.dry:
        print(f"Excel '{excel}' sheet '{sheet}' rows={len(df)}")
        print(f"Stations: {station_by_id}")
        return

    dlg = attach_or_launch(app, title, work_dir=work_dir, attach_only=attach_only, extra_load_wait=extra_wait)
    r = dlg.rectangle()

    # ---- Screen coords heuristics (relative to window) ----
    mesh_x, mesh_y     = r.left+150, r.bottom-92
    station_x, station_y = r.left+150, r.bottom-62
    camera_x, camera_y   = r.left+150, r.bottom-32
    frame_x, frame_y     = r.left+515, r.bottom-92
    posx_x, posx_y       = r.left+515, r.bottom-62
    posy_x, posy_y       = r.left+515, r.bottom-32
    show_x, show_y       = r.left+655, r.bottom-70

    # 1) Select Mesh by OCR on dropdown
    if not select_from_dropdown_at(mesh_x, mesh_y, mesh_name):
        # fallback: click first row then Enter
        open_combo_at(mesh_x, mesh_y); pyautogui.click(mesh_x+100, mesh_y+35); pyautogui.press("enter")
    time.sleep(0.5)

    # Iterate rows
    nr_col = next((c for c in df.columns if c.startswith("NR")), df.columns[0])

    for idx,row in df.iterrows():
        try:
            st_id=int(to_float(row["STATION"]) or row["STATION"])
            cam_id=int(to_float(row["CAMERA"]) or row["CAMERA"])
            st_name = station_by_id.get(st_id, None)
            if not st_name: print(f"[WARN] Row {idx}: unknown station id {st_id}"); continue

            # 2) Station via OCR dropdown
            if not select_from_dropdown_at(station_x, station_y, st_name):
                # index fallback
                open_combo_at(station_x, station_y); pyautogui.moveTo(station_x+110, station_y+30+STATION_ORDER.index(st_name)*22); pyautogui.click(); pyautogui.press("enter")
            time.sleep(0.1)
            # 3) Camera via OCR dropdown
            if not select_from_dropdown_at(camera_x, camera_y, str(cam_id)):
                open_combo_at(camera_x, camera_y); pyautogui.moveTo(camera_x+90, camera_y+30+(cam_id-1)*20); pyautogui.click(); pyautogui.press("enter")
            time.sleep(0.1)

            # 4) Inputs
            put_text_xy(frame_x, frame_y, row["FRAME"])
            put_text_xy(posx_x, posx_y, row["X_FRAME"])
            put_text_xy(posy_x, posy_y, row["Y_FRAME"])

            # 5) Show
            click_xy(show_x, show_y, double=True)
            time.sleep(0.3 + float(cfg.get("post_show_delay_s", 0.7)))

            # 6) OCR result
            # For simplicity, capture a mid-right region where the text lives (adjust if needed)
            res_img = pyautogui.screenshot(region=(r.left+880, r.bottom-120, 520, 110))
            bgr = cv2.cvtColor(np.array(res_img), cv2.COLOR_RGBA2BGR)
            text = pytesseract.image_to_string(bgr, lang="eng", config="--psm 6")
            m = re.search(r"Intersect position:\s*([-+]?\d+,\d+)\s+([-+]?\d+,\d+)\s+([-+]?\d+,\d+)", text)
            X=Y=Z=T=None
            if m: X,Y,Z = m.group(1), m.group(2), m.group(3)
            m2 = re.search(r"([\d,]+)\s*ms", text); 
            if m2: T = m2.group(1)
            df.at[idx,"X_NEW"]=X or ""; df.at[idx,"Y_NEW"]=Y or ""; df.at[idx,"Z_NEW"]=Z or ""; df.at[idx,"TIME_MS"]=T or ""
            print(f"[OK] {row.get(nr_col, idx+1)} -> ({X},{Y},{Z}) {T} ms")
            time.sleep(float(cfg.get("per_row_pause_s", 0.2)))
        except Exception as e:
            print(f"[ERR] Row {idx}: {e}")

    out=cfg.get("output_excel_path") or excel
    try:
        with pd.ExcelWriter(out, engine="openpyxl", mode="w") as w: df.to_excel(w, sheet_name=sheet, index=False)
        print(f"[DONE] Saved to {out}")
    except PermissionError:
        alt = str(Path(out).with_name(Path(out).stem + "_out.xlsx"))
        with pd.ExcelWriter(alt, engine="openpyxl", mode="w") as w: df.to_excel(w, sheet_name=sheet, index=False)
        print(f"[WARN] Excel locked; saved to {alt}")
    
if __name__=="__main__": main()
