# eines_automation_v6.py
# - Uses OCR to pick items in owner-drawn dropdowns (Station/Camera)
# - No typing fallback (your viewer ignores typed text)
# - Mouse "jiggle" before clicks; double-click + Enter for reliability
# - Keeps index-based fallback if OCR fails
import argparse, os, re, time, json, pandas as pd, xml.etree.ElementTree as ET
from pathlib import Path

# UI / OCR
from pywinauto.application import Application
from pywinauto import Desktop
import psutil
import pyautogui
import pytesseract
import cv2
import numpy as np

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
    # Try attach
    try:
        dlg = Desktop(backend="uia").window(title_re=".*"+re.escape(title_hint)+".*", found_index=0)
        if dlg.exists(): dlg.set_focus(); time.sleep(extra_load_wait); return dlg
    except Exception:
        pass
    byproc = _find_window_by_process(os.path.basename(app_path), title_hint, timeout=5)
    if byproc is not None: byproc.set_focus(); time.sleep(extra_load_wait); return byproc
    if attach_only: raise RuntimeError("Window not found to attach.")
    # Launch
    if work_dir is None: work_dir = os.path.dirname(app_path)
    app = Application(backend="uia").start(cmd_line=f'"{app_path}"', work_dir=work_dir)
    time.sleep(3.0)
    # Wait
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

def find_controls(dlg):
    ctrls = {}
    try: ctrls["mesh"]    = dlg.child_window(title_re="^Mesh$", control_type="ComboBox")
    except: pass
    try: ctrls["station"] = dlg.child_window(title_re="^Station$", control_type="ComboBox")
    except: pass
    try: ctrls["camera"]  = dlg.child_window(title_re="^Camera$", control_type="ComboBox")
    except: pass
    try: ctrls["frame"]   = dlg.child_window(title_re="^Frame$", control_type="Edit")
    except: pass
    try: ctrls["posx"]    = dlg.child_window(title_re="^Position X$", control_type="Edit")
    except: pass
    try: ctrls["posy"]    = dlg.child_window(title_re="^Position Y$", control_type="Edit")
    except: pass
    try: ctrls["show"]    = dlg.child_window(title_re="^Show$", control_type="Button")
    except: pass
    try: ctrls["intersect_text"] = dlg.child_window(title_re="Intersect position.*", control_type="Text")
    except: pass
    return ctrls

def jiggle():
    x,y = pyautogui.position()
    pyautogui.moveTo(x+2, y+2, duration=0.02)
    pyautogui.moveTo(x, y, duration=0.02)

def click_center(rect, double=False):
    x = rect.left + rect.width()//2
    y = rect.top + rect.height()//2
    jiggle()
    pyautogui.click(x, y)
    if double:
        time.sleep(0.05)
        pyautogui.click(x, y)

def open_combo(ctrl):
    try:
        ctrl.click_input()
    except Exception:
        rect = ctrl.rectangle()
        click_center(rect)
    time.sleep(0.15)

def ocr_dropdown_region(ctrl):
    """Capture area under combo to detect dropdown list text lines and their y positions."""
    rect = ctrl.rectangle()
    # capture generous area below the combo
    region = (rect.left, rect.bottom, max(280, rect.width()+200), 260)
    img = pyautogui.screenshot(region=region)
    bgr = cv2.cvtColor(np.array(img), cv2.COLOR_RGBA2BGR)
    gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    # light threshold to improve OCR
    thr = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY)[1]
    text = pytesseract.image_to_string(thr, lang="eng", config="--psm 6")
    # split lines and estimate y for each line using horizontal projections
    proj = np.sum(255-thr, axis=1)
    # find peaks (text rows)
    rows = []
    in_band=False; start=0
    for i,v in enumerate(proj):
        if v>0 and not in_band:
            in_band=True; start=i
        elif v==0 and in_band:
            in_band=False; rows.append((start,i))
    lines = [text.strip() for text in text.splitlines() if text.strip()]
    bands = [( (a+b)//2 ) for a,b in rows if (b-a)>6]
    return region, lines, bands

def select_from_dropdown_by_text(ctrl, target_text):
    open_combo(ctrl)
    region, lines, bands = ocr_dropdown_region(ctrl)
    # try exact match
    idx = None
    for i,txt in enumerate(lines):
        if txt.strip().upper() == str(target_text).upper():
            idx = i; break
    if idx is None:
        # try contains
        for i,txt in enumerate(lines):
            if str(target_text).upper() in txt.strip().upper():
                idx = i; break
    if idx is not None and idx < len(bands):
        y_local = bands[idx]
        x = region[0] + 100
        y = region[1] + y_local
        jiggle(); pyautogui.click(x, y); time.sleep(0.05); pyautogui.press("enter")
        return True
    return False

def select_station(ctrls, station_name):
    ctrl = ctrls["station"]
    if select_from_dropdown_by_text(ctrl, station_name):
        return
    # fallback by index (order known)
    try:
        idx = STATION_ORDER.index(station_name)
    except ValueError:
        idx = 0
    open_combo(ctrl)
    # assume ~22 px per item
    rect = ctrl.rectangle()
    base_x = rect.left + 120
    base_y = rect.bottom + 12 + idx*22
    jiggle(); pyautogui.click(base_x, base_y); time.sleep(0.05); pyautogui.press("enter")

def select_camera(ctrls, cam_id):
    ctrl = ctrls["camera"]
    if select_from_dropdown_by_text(ctrl, str(cam_id)):
        return
    # fallback by index: cameras 1..N
    idx = max(int(cam_id)-1, 0)
    open_combo(ctrl)
    rect = ctrl.rectangle()
    base_x = rect.left + 80
    base_y = rect.bottom + 12 + idx*20
    jiggle(); pyautogui.click(base_x, base_y); time.sleep(0.05); pyautogui.press("enter")

def put_text(ctrl, value):
    s=str(value).strip(); s=s.replace('.', ',') if DECIMAL_COMMA else s
    try: ctrl.set_edit_text(s)
    except Exception:
        rect = ctrl.rectangle(); click_center(rect)
        pyautogui.hotkey("ctrl","a"); pyautogui.typewrite(s)

def click_show(ctrl):
    try: ctrl.click_input(); return
    except Exception: pass
    rect = ctrl.rectangle(); click_center(rect, double=True)

def ocr_intersect(dlg, intersect_ctrl=None):
    try:
        if intersect_ctrl:
            rect = intersect_ctrl.rectangle()
            img = pyautogui.screenshot(region=(rect.left, rect.top, rect.width(), rect.height()+6))
        else:
            r = dlg.rectangle()
            img = pyautogui.screenshot(region=(r.left, r.top, r.width(), r.height()))
        frame = cv2.cvtColor(np.array(img), cv2.COLOR_RGBA2BGR)
    except Exception:
        r = dlg.rectangle()
        img = pyautogui.screenshot(region=(r.left, r.top, r.width(), r.height()))
        frame = cv2.cvtColor(np.array(img), cv2.COLOR_RGBA2BGR)
    return pytesseract.image_to_string(frame, lang="eng", config="--psm 6")

def parse_intersect(text):
    m = re.search(r"Intersect position:\s*([-+]?\d+,\d+)\s+([-+]?\d+,\d+)\s+([-+]?\d+,\d+)", text)
    X=Y=Z=T=None
    if m: X,Y,Z = m.group(1), m.group(2), m.group(3)
    m2 = re.search(r"([\d,]+)\s*ms", text)
    if m2: T = m2.group(1)
    return X,Y,Z,T,text

def safe_save_excel(df, sheet, out_path):
    try:
        with pd.ExcelWriter(out_path, engine="openpyxl", mode="w") as w:
            df.to_excel(w, sheet_name=sheet, index=False)
        print(f"[DONE] Saved to {out_path}")
    except PermissionError:
        alt = str(Path(out_path).with_name(Path(out_path).stem + "_out.xlsx"))
        with pd.ExcelWriter(alt, engine="openpyxl", mode="w") as w:
            df.to_excel(w, sheet_name=sheet, index=False)
        print(f"[WARN] Excel locked; saved to {alt}")

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument("--config", default="config_v6.json")
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
    st_by_id,_=load_stations_map(xmlp)

    if args.dry:
        print(f"Excel '{excel}' sheet '{sheet}' rows={len(df)}")
        print(f"Stations: {st_by_id}")
        return

    dlg = attach_or_launch(app, title, work_dir=work_dir, attach_only=attach_only, extra_load_wait=extra_wait)
    # Focus & jiggle
    r=dlg.rectangle(); pyautogui.click(r.left+150, r.bottom-120); jiggle(); time.sleep(0.2)

    ctrls = find_controls(dlg)
    # Mesh first
    if "mesh" in ctrls:
        open_combo(ctrls["mesh"])
        # try OCR choose mesh by text
        if not select_from_dropdown_by_text(ctrls["mesh"], mesh_name):
            # fall back: click first item then type Enter repeatedly until label changes
            pyautogui.click(ctrls["mesh"].rectangle().left+100, ctrls["mesh"].rectangle().bottom+20)
            pyautogui.press("enter")
    time.sleep(0.6)
    ctrls = find_controls(dlg)

    nr_col = next((c for c in df.columns if c.startswith("NR")), df.columns[0])

    for idx,row in df.iterrows():
        try:
            st_id=int(to_float(row["STATION"]) or row["STATION"]); cam_id=int(to_float(row["CAMERA"]) or row["CAMERA"])
            st_name = st_by_id.get(st_id, None)
            if not st_name: print(f"[WARN] Row {idx}: unknown station id {st_id}"); continue

            print(f"[STEP] Row {idx}: Station={st_name}  Camera={cam_id}")
            select_station(ctrls, st_name)
            time.sleep(0.15)
            select_camera(ctrls, cam_id)
            time.sleep(0.15)

            if "frame" in ctrls: put_text(ctrls["frame"], row["FRAME"])
            if "posx" in ctrls:  put_text(ctrls["posx"], row["X_FRAME"])
            if "posy" in ctrls:  put_text(ctrls["posy"], row["Y_FRAME"])

            if "show" in ctrls: click_show(ctrls["show"])
            else: pyautogui.press("enter")

            time.sleep(0.3 + float(cfg.get("post_show_delay_s", 0.7)))

            text = ocr_intersect(dlg, ctrls.get("intersect_text"))
            X,Y,Z,T,_ = parse_intersect(text)
            df.at[idx,"X_NEW"]=X or ""; df.at[idx,"Y_NEW"]=Y or ""; df.at[idx,"Z_NEW"]=Z or ""; df.at[idx,"TIME_MS"]=T or ""
            print(f"[OK] {row.get(nr_col, idx+1)} -> ({X},{Y},{Z}) {T} ms")
            time.sleep(float(cfg.get("per_row_pause_s", 0.2)))
        except Exception as e:
            print(f"[ERR] Row {idx}: {e}")

    out=cfg.get("output_excel_path") or excel
    safe_save_excel(df, sheet, out)

if __name__=="__main__": main()
