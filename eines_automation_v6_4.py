# eines_automation_v6_4.py
import argparse, os, re, time, json, pandas as pd, xml.etree.ElementTree as ET
from pathlib import Path
from pywinauto.application import Application
from pywinauto import Desktop
import pyautogui, pytesseract, cv2, numpy as np, psutil

pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0.03

DECIMAL_COMMA = True
STATION_ORDER = ["RIGHTSIDE1","RIGHTSIDE2","LEFTSIDE1","LEFTSIDE2","ROOF","ROOFHOOD","HOOD","BRIDGE"]
TARGET_LABELS = {
    "MESH": "Mesh", "STATION": "Station", "CAMERA": "Camera",
    "FRAME": "Frame", "POSX": "Position X", "POSY": "Position Y", "SHOW": "Show",
}

def to_float(s):
    if s is None or (isinstance(s, float) and pd.isna(s)): return None
    ss = str(s).strip()
    if DECIMAL_COMMA:
        if ss.count(',')==1 and ss.count('.')>1:
            ss = ss.replace('.', '').replace(',', '.')
        else:
            ss = ss.replace(',', '.')
    try: return float(ss)
    except: return None

def load_excel(path):
    xls = pd.ExcelFile(path)
    sheet = "Hoja1" if "Hoja1" in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(path, sheet_name=sheet, header=0)
    df.columns = [str(c).strip().upper().replace(" ", "_") for c in df.columns]
    for c in ["X_NEW","Y_NEW","Z_NEW","TIME_MS"]:
        if c not in df.columns: df[c] = ""
        df[c] = df[c].astype("object")
    return df, sheet, 0

def load_stations_map(xml_path):
    tree = ET.parse(xml_path); root = tree.getroot()
    station_by_id = {}
    for st in root.findall("station"):
        try: station_by_id[int(st.attrib["id"])] = st.findtext("name")
        except: pass
    return station_by_id

def _find_window_by_process(exe_basename, title_hint, timeout=60):
    t0=time.time()
    while time.time()-t0<timeout:
        for w in Desktop(backend="uia").windows():
            try:
                pid = w.process_id(); p = psutil.Process(pid)
                if exe_basename.lower() in os.path.basename(p.exe()).lower():
                    if (not title_hint) or (title_hint.lower() in w.window_text().lower()):
                        return w
            except Exception: continue
        time.sleep(0.5)
    return None

def attach_or_launch(app_path, title_hint, work_dir=None, attach_only=False, extra_load_wait=3.0):
    try:
        dlg = Desktop(backend="uia").window(title_re=".*"+re.escape(title_hint)+".*", found_index=0)
        if dlg.exists(): dlg.set_focus(); time.sleep(extra_load_wait); dlg.maximize(); return dlg
    except Exception: pass
    byproc = _find_window_by_process(os.path.basename(app_path), title_hint, timeout=5)
    if byproc is not None: byproc.set_focus(); time.sleep(extra_load_wait); byproc.maximize(); return byproc
    if attach_only: raise RuntimeError("Window not found to attach.")
    if work_dir is None: work_dir = os.path.dirname(app_path)
    app = Application(backend="uia").start(cmd_line=f'"{app_path}"', work_dir=work_dir)
    time.sleep(3.0)
    t0=time.time(); dlg=None
    while time.time()-t0<60:
        try:
            dlg = Desktop(backend="uia").window(title_re=".*"+re.escape(title_hint)+".*")
            if dlg.exists(): dlg.wait("visible", timeout=5); break
        except Exception: pass
        dlg = _find_window_by_process(os.path.basename(app_path), title_hint, timeout=1)
        if dlg: break
        time.sleep(1.0)
    if dlg is None: raise RuntimeError("Timed out waiting for window")
    dlg.set_focus(); time.sleep(extra_load_wait)
    try: dlg.maximize()
    except: pass
    return dlg

def screenshot_region(rect):
    x,y,w,h = rect
    img = pyautogui.screenshot(region=(x,y,w,h))
    return cv2.cvtColor(np.array(img), cv2.COLOR_RGBA2BGR)

def ocr_data(img_bgr, psm=6, timeout=2.0):
    try:
        from pytesseract import image_to_data, Output
        return pytesseract.image_to_data(img_bgr, lang="eng", config=f"--psm {psm}", output_type=Output.DICT, timeout=timeout)
    except Exception:
        return None

def find_footer_labels_dynamic(dlg):
    """Desplaza una ventana OCR desde bottom-200 hasta bottom-60 y escoge donde detecta más etiquetas."""
    r = dlg.rectangle()
    left, top, right, bottom = r.left, r.top, r.right, r.bottom
    best = (0, {}, (left, bottom-140, right-left, 140))
    for offset in range(200, 60, -10):
        rect = (left, bottom-offset, right-left, 140)
        bgr = screenshot_region(rect)
        gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
        thr = cv2.adaptiveThreshold(gray,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY,31,5)
        data = ocr_data(thr, psm=6, timeout=1.6)
        found = {}
        if data:
            for i,txt in enumerate(data["text"]):
                t = (txt or "").strip()
                if not t: continue
                for key,wanted in TARGET_LABELS.items():
                    if t.upper() == wanted.upper():
                        x = rect[0] + data["left"][i]; y = rect[1] + data["top"][i]
                        w = data["width"][i]; h = data["height"][i]
                        found[key]=(x,y,w,h)
        score = len(found)
        if score > best[0]:
            best = (score, found, rect)
            if score >= 5: break
    # fallback aprox si faltan
    labels = best[1]
    rect = best[2]
    if len(labels) < 5:
        fx,fy,fw,fh = rect
        approx = {
            "MESH":   (fx+40, fy+15, 40, 18),
            "STATION":(fx+40, fy+40, 60, 18),
            "CAMERA": (fx+40, fy+65, 55, 18),
            "FRAME":  (fx+360, fy+15, 45, 18),
            "POSX":   (fx+360, fy+40, 80, 18),
            "POSY":   (fx+360, fy+65, 80, 18),
            "SHOW":   (fx+540, fy+28, 45, 30),
        }
        for k,v in approx.items():
            labels.setdefault(k, v)
    return labels, rect

def click_right_of(label_rect, dx=120):
    x = label_rect[0] + label_rect[2] + dx
    y = label_rect[1] + label_rect[3]//2
    pyautogui.moveTo(x, y, duration=0.05); pyautogui.click()
    return x,y

def open_dropdown_at(x,y):
    pyautogui.moveTo(x,y, duration=0.05); pyautogui.click(); time.sleep(0.15)

def select_from_dropdown_by_text_near(x,y, target_text, debug_prefix=None):
    open_dropdown_at(x,y)
    region = (x-90, y+12, 440, 340)
    bgr = screenshot_region(region)
    if debug_prefix: cv2.imwrite(f"{debug_prefix}_dropdown.png", bgr)
    gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    thr = cv2.adaptiveThreshold(gray,255,cv2.ADAPTIVE_THRESH_MEAN_C,cv2.THRESH_BINARY,31,7)
    data = ocr_data(thr, psm=6, timeout=1.6)
    if data:
        target_up = str(target_text).upper()
        best=None
        for i,txt in enumerate(data["text"]):
            t=(txt or "").strip()
            if not t: continue
            if (t.upper()==target_up) or (target_up in t.upper()):
                best=(data["left"][i], data["top"][i], data["width"][i], data["height"][i]); break
        if best:
            cx = region[0] + best[0] + best[2]//2
            cy = region[1] + best[1] + best[3]//2
            pyautogui.moveTo(cx, cy, duration=0.05); pyautogui.click(); pyautogui.press("enter")
            return True
    # fallback: click primera fila
    pyautogui.moveTo(region[0]+150, region[1]+30, duration=0.05); pyautogui.click(); pyautogui.press("enter")
    return False

def put_text_at(x,y,value):
    s = str(value).strip(); s = s.replace('.', ',') if DECIMAL_COMMA else s
    pyautogui.moveTo(x,y, duration=0.05); pyautogui.click(); pyautogui.hotkey("ctrl","a"); pyautogui.typewrite(s)

def capture_result_block(dlg):
    r = dlg.rectangle()
    # eleva 20px por el tema de taskbar / margen
    rect = (r.left + 700, r.bottom - 130, 640, 125)
    return screenshot_region(rect)

def parse_intersect_text(text):
    m = re.search(r"Intersect position:\s*([-+]?\d+,\d+)\s+([-+]?\d+,\d+)\s+([-+]?\d+,\d+)", text)
    X=Y=Z=T=None
    if m: X,Y,Z = m.group(1), m.group(2), m.group(3)
    m2 = re.search(r"([\d,]+)\s*ms", text)
    if m2: T = m2.group(1)
    return X,Y,Z,T

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument("--config", default="config_v6_4.json")
    ap.add_argument("--run", action="store_true")
    args=ap.parse_args()

    cfg=json.load(open(args.config,"r",encoding="utf-8"))
    if cfg.get("tesseract_cmd"):
        pytesseract.pytesseract.tesseract_cmd = cfg["tesseract_cmd"]

    df, sheet, _ = load_excel(cfg["excel_path"])
    stations = load_stations_map(cfg["stations_xml_path"])

    dlg = attach_or_launch(cfg["viewer_exe_path"], cfg.get("window_title_contains","EINES System 3D for ESPQi"),
                           work_dir=cfg.get("work_dir"), attach_only=bool(cfg.get("attach_only", False)),
                           extra_load_wait=float(cfg.get("extra_load_wait_s",3.0)))
    # detectar footer dinámico
    labels, footer_rect = find_footer_labels_dynamic(dlg)

    # Mesh (con 2 reintentos subiendo 15px por si cayó bajo)
    mx,my = click_right_of(labels["MESH"], dx=120)
    ok = select_from_dropdown_by_text_near(mx, my, cfg.get("mesh_name","TAYCAN_NACHFOLGER"), debug_prefix="debug_mesh")
    if not ok:
        for dy in (-15, -30):
            ok = select_from_dropdown_by_text_near(mx, my+dy, cfg.get("mesh_name","TAYCAN_NACHFOLGER"), debug_prefix=f"debug_mesh_retry{abs(dy)}")
            if ok: break
    time.sleep(0.4)

    nr_col = next((c for c in df.columns if c.startswith("NR")), df.columns[0])

    for idx,row in df.iterrows():
        try:
            st_id=int(to_float(row["STATION"]) or row["STATION"]); cam_id=int(to_float(row["CAMERA"]) or row["CAMERA"])
            st_name = stations.get(st_id, None)
            if not st_name: print(f"[WARN] Row {idx}: unknown station {st_id}"); continue

            # Station
            sx,sy = click_right_of(labels["STATION"], dx=120)
            ok_st = select_from_dropdown_by_text_near(sx, sy, st_name, debug_prefix=f"debug_row{idx}_station")
            if not ok_st and st_name in STATION_ORDER:
                open_dropdown_at(sx, sy); pyautogui.moveTo(sx+110, sy+30+STATION_ORDER.index(st_name)*22, duration=0.05); pyautogui.click(); pyautogui.press("enter")

            # Camera
            cx,cy = click_right_of(labels["CAMERA"], dx=120)
            ok_cam = select_from_dropdown_by_text_near(cx, cy, str(cam_id), debug_prefix=f"debug_row{idx}_camera")
            if not ok_cam:
                open_dropdown_at(cx, cy); pyautogui.moveTo(cx+100, cy+30+(max(cam_id-1,0))*20, duration=0.05); pyautogui.click(); pyautogui.press("enter")

            # Inputs
            fx,fy = click_right_of(labels["FRAME"], dx=140); put_text_at(fx, fy, row["FRAME"])
            px,py_ = click_right_of(labels["POSX"], dx=140); put_text_at(px, py_, row["X_FRAME"])
            pyx,pyy = click_right_of(labels["POSY"], dx=140); put_text_at(pyx, pyy, row["Y_FRAME"])

            # Show (doble click)
            sh = labels["SHOW"]; sx2 = sh[0]+sh[2]//2+60; sy2 = sh[1]+sh[3]//2
            pyautogui.moveTo(sx2, sy2, duration=0.05); pyautogui.click(); time.sleep(0.05); pyautogui.click()

            # Read result (3 intentos)
            coords=None; T=None
            for t in range(3):
                time.sleep(0.3 + float(cfg.get("post_show_delay_s",0.7)))
                img = capture_result_block(dlg)
                gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
                thr = cv2.adaptiveThreshold(gray,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY,31,5)
                text = pytesseract.image_to_string(thr, lang="eng", config="--psm 6", timeout=1.6)
                X,Y,Z,T = parse_intersect_text(text)
                if X and Y and Z:
                    coords=(X,Y,Z); break
                if idx<5: cv2.imwrite(f"debug_result_row_{idx}_try{t}.png", img)
            X=Y=Z=None
            if coords: X,Y,Z = coords
            df.at[idx,"X_NEW"]=X or ""
            df.at[idx,"Y_NEW"]=Y or ""
            df.at[idx,"Z_NEW"]=Z or ""
            df.at[idx,"TIME_MS"]=T or ""
            print(f"[OK] {row.get(nr_col, idx+1)} -> ({X},{Y},{Z}) {T} ms")
        except Exception as e:
            print(f"[ERR] Row {idx}: {e}")

    out = cfg.get("output_excel_path") or cfg["excel_path"]
    try:
        with pd.ExcelWriter(out, engine="openpyxl", mode="w") as w: df.to_excel(w, sheet_name=sheet, index=False)
        print(f"[DONE] Saved to {out}")
    except PermissionError:
        alt = str(Path(out).with_name(Path(out).stem + "_out.xlsx"))
        with pd.ExcelWriter(alt, engine="openpyxl", mode="w") as w: df.to_excel(w, sheet_name=sheet, index=False)
        print(f"[WARN] Excel locked; saved to {alt}")
if __name__=="__main__": main()
