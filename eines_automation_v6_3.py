# eines_automation_v6_3.py
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
    "MESH": "Mesh",
    "STATION": "Station",
    "CAMERA": "Camera",
    "FRAME": "Frame",
    "POSX": "Position X",
    "POSY": "Position Y",
    "SHOW": "Show",
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
    df_raw = pd.read_excel(path, sheet_name=sheet, header=None, dtype=str)
    header_row = None
    for i in range(min(30, len(df_raw))):
        row_vals = [str(v).strip() for v in df_raw.iloc[i].tolist()]
        if "X_FRAME" in row_vals and "Y_FRAME" in row_vals and "FRAME" in row_vals:
            header_row = i; break
        if "X_PAG" in row_vals and "X_EINES" in row_vals:
            header_row = i
    if header_row is None: header_row = 0
    df = pd.read_excel(path, sheet_name=sheet, header=header_row)
    df.columns = [str(c).strip().upper().replace(" ", "_") for c in df.columns]
    for c in ["X_NEW","Y_NEW","Z_NEW","TIME_MS"]:
        if c not in df.columns: df[c] = ""
        df[c] = df[c].astype("object")
    # try sort by first NR-ish column
    for nr_name in ["NR","NR.","Nº","N°","NO","NUM","NUMERO"]:
        if nr_name in df.columns:
            try:
                df[nr_name+"_SORT"] = pd.to_numeric(df[nr_name], errors="coerce")
                df = df.sort_values(by=[nr_name+"_SORT"]).drop(columns=[nr_name+"_SORT"])
            except Exception: pass
            break
    return df, sheet, header_row

def load_stations_map(xml_path):
    tree = ET.parse(xml_path); root = tree.getroot()
    station_by_id = {}
    for st in root.findall("station"):
        try:
            station_by_id[int(st.attrib["id"])] = st.findtext("name")
        except Exception:
            pass
    return station_by_id

def _find_window_by_process(exe_basename, title_hint, timeout=60):
    t0=time.time()
    while time.time()-t0<timeout:
        for w in Desktop(backend="uia").windows():
            try:
                title = w.window_text()
                pid = w.process_id()
                p = psutil.Process(pid)
                if exe_basename.lower() in os.path.basename(p.exe()).lower():
                    if (not title_hint) or (title_hint.lower() in title.lower()):
                        return w
            except Exception:
                continue
        time.sleep(0.5)
    return None

def attach_or_launch(app_path, title_hint, work_dir=None, attach_only=False, extra_load_wait=3.0):
    # try attach by title
    try:
        dlg = Desktop(backend="uia").window(title_re=".*"+re.escape(title_hint)+".*", found_index=0)
        if dlg.exists():
            dlg.set_focus(); time.sleep(extra_load_wait); return dlg
    except Exception:
        pass
    # try attach by process
    byproc = _find_window_by_process(os.path.basename(app_path), title_hint, timeout=5)
    if byproc is not None:
        byproc.set_focus(); time.sleep(extra_load_wait); return byproc
    if attach_only:
        raise RuntimeError("Window not found to attach. Open it manually.")
    # launch
    if work_dir is None: work_dir = os.path.dirname(app_path)
    app = Application(backend="uia").start(cmd_line=f'"{app_path}"', work_dir=work_dir)
    time.sleep(3.0)
    # wait for window
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
    try: dlg.maximize()
    except Exception: pass
    return dlg

# ---------- OCR helpers ----------
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

def ocr_text(img_bgr, psm=6, timeout=2.0):
    try:
        return pytesseract.image_to_string(img_bgr, lang="eng", config=f"--psm {psm}", timeout=timeout)
    except Exception:
        return ""

def find_labels_in_footer_rect(window_rect):
    """OCR the bottom strip to locate labels; returns dict of label rects and footer rect."""
    # window_rect is tuple (left, top, right, bottom) as from dlg.rectangle()
    left, top, right, bottom = window_rect.left, window_rect.top, window_rect.right, window_rect.bottom
    footer_h = 110
    rect = (left, bottom - footer_h, right-left, footer_h)
    bgr = screenshot_region(rect)
    data = ocr_data(bgr, psm=6, timeout=2.0)
    labels = {}
    if data:
        n = len(data["text"])
        for i in range(n):
            txt = (data["text"][i] or "").strip()
            if not txt: continue
            up = txt.upper()
            for key, wanted in TARGET_LABELS.items():
                if up == wanted.upper():
                    x = rect[0] + data["left"][i]
                    y = rect[1] + data["top"][i]
                    w = data["width"][i]
                    h = data["height"][i]
                    labels[key] = (x,y,w,h)
    return labels, rect

def click_right_of(label_rect, dx=120, dy=0):
    x = label_rect[0] + label_rect[2] + dx
    y = label_rect[1] + label_rect[3]//2 + dy
    pyautogui.moveTo(x, y, duration=0.05)
    pyautogui.click(x, y)
    return x,y

def open_dropdown_at(x,y):
    pyautogui.moveTo(x, y, duration=0.05); pyautogui.click(x, y); time.sleep(0.15)

def select_from_dropdown_by_text_near(x,y, target_text, debug_prefix=None):
    """Open dropdown at (x,y), OCR the list below, click item by text; return True/False."""
    open_dropdown_at(x,y)
    region = (x-80, y+10, 420, 320)
    bgr = screenshot_region(region)
    if debug_prefix: cv2.imwrite(f"{debug_prefix}_dropdown.png", bgr)
    gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    # adaptive threshold helps on gradients
    thr = cv2.adaptiveThreshold(gray,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY,31,5)
    data = ocr_data(thr, psm=6, timeout=2.0)
    if data:
        items = []
        for i,txt in enumerate(data["text"]):
            if txt and txt.strip():
                items.append((txt.strip(), data["left"][i], data["top"][i], data["width"][i], data["height"][i]))
        target_up = str(target_text).upper()
        # exact first, then contains
        best=None
        for (txt,l,t,w,h) in items:
            if txt.upper()==target_up: best=(l,t,w,h); break
        if not best:
            for (txt,l,t,w,h) in items:
                if target_up in txt.upper(): best=(l,t,w,h); break
        if best:
            cx = region[0] + best[0] + best[2]//2
            cy = region[1] + best[1] + best[3]//2
            pyautogui.moveTo(cx, cy, duration=0.05); pyautogui.click(); pyautogui.press("enter")
            return True
    # fallback: first option
    pyautogui.moveTo(region[0]+140, region[1]+28, duration=0.05); pyautogui.click(); pyautogui.press("enter")
    return False

def put_text_at(x,y,value):
    s = str(value).strip()
    s = s.replace('.', ',') if DECIMAL_COMMA else s
    pyautogui.moveTo(x,y, duration=0.05); pyautogui.click(); pyautogui.hotkey("ctrl","a"); pyautogui.typewrite(s)

def find_intersect_text(window_rect):
    """OCR right footer block and parse Intersect position line."""
    left, top, right, bottom = window_rect.left, window_rect.top, window_rect.right, window_rect.bottom
    res_rect = (left + 700, bottom - 110, 620, 105)  # wide capture
    bgr = screenshot_region(res_rect)
    # save debug optionally outside
    return bgr

def parse_intersect_from_image(bgr):
    # convert and OCR
    gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    thr = cv2.adaptiveThreshold(gray,255,cv2.ADAPTIVE_THRESH_MEAN_C,cv2.THRESH_BINARY,31,7)
    text = ocr_text(thr, psm=6, timeout=2.0)
    # try parse
    m = re.search(r"Intersect position:\s*([-+]?\d+,\d+)\s+([-+]?\d+,\d+)\s+([-+]?\d+,\d+)", text)
    X=Y=Z=T=None
    if m: X,Y,Z = m.group(1), m.group(2), m.group(3)
    m2 = re.search(r"([\d,]+)\s*ms", text)
    if m2: T = m2.group(1)
    return X,Y,Z,T,text

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument("--config", default="config_v6_3.json")
    ap.add_argument("--run", action="store_true")
    ap.add_argument("--dry", action="store_true")
    args=ap.parse_args()

    cfg=json.load(open(args.config,"r",encoding="utf-8"))
    if cfg.get("tesseract_cmd"):
        pytesseract.pytesseract.tesseract_cmd = cfg["tesseract_cmd"]

    excel=cfg["excel_path"]; xmlp=cfg["stations_xml_path"]; app=cfg["viewer_exe_path"]
    title=cfg.get("window_title_contains","EINES System 3D for ESPQi")
    mesh_name=cfg.get("mesh_name","TAYCAN_NACHFOLGER")
    work_dir=cfg.get("work_dir", os.path.dirname(app))
    attach_only=bool(cfg.get("attach_only", False))
    extra_wait=float(cfg.get("extra_load_wait_s", 3.0))

    df,sheet,_=load_excel(excel)
    st_by_id=load_stations_map(xmlp)

    if args.dry:
        print(f"Excel '{excel}' sheet '{sheet}' rows={len(df)}")
        print(f"Stations: {st_by_id}")
        return

    dlg = attach_or_launch(app, title, work_dir=work_dir, attach_only=attach_only, extra_load_wait=extra_wait)
    # cache window rect once to avoid re-query UIA repeatedly
    wrect = dlg.rectangle()

    # find footer labels
    labels, footer_rect = find_labels_in_footer_rect(wrect)
    # fallback approx if missing
    for key in ["MESH","STATION","CAMERA","FRAME","POSX","POSY","SHOW"]:
        if key not in labels:
            fx,fy,fw,fh = footer_rect
            approx = {
                "MESH":   (fx+40, fy+15, 40, 18),
                "STATION":(fx+40, fy+40, 60, 18),
                "CAMERA": (fx+40, fy+65, 55, 18),
                "FRAME":  (fx+360, fy+15, 45, 18),
                "POSX":   (fx+360, fy+40, 80, 18),
                "POSY":   (fx+360, fy+65, 80, 18),
                "SHOW":   (fx+540, fy+28, 45, 30),
            }
            labels[key]=approx[key]

    # Mesh selection
    mesh_xy = click_right_of(labels["MESH"], dx=120)
    select_from_dropdown_by_text_near(*mesh_xy, target_text=mesh_name, debug_prefix="debug_mesh")
    time.sleep(0.4)

    nr_col = next((c for c in df.columns if c.startswith("NR")), df.columns[0])

    for idx,row in df.iterrows():
        try:
            st_id=int(to_float(row["STATION"]) or row["STATION"]); cam_id=int(to_float(row["CAMERA"]) or row["CAMERA"])
            st_name = st_by_id.get(st_id, None)
            if not st_name:
                print(f"[WARN] Row {idx}: unknown station id {st_id}"); continue

            # Station/Camera
            st_xy = click_right_of(labels["STATION"], dx=120)
            ok_st = select_from_dropdown_by_text_near(*st_xy, target_text=st_name, debug_prefix=f"debug_row{idx}_station")
            if not ok_st and st_name in STATION_ORDER:
                open_dropdown_at(*st_xy); pyautogui.moveTo(st_xy[0]+100, st_xy[1]+30 + STATION_ORDER.index(st_name)*22, duration=0.05); pyautogui.click(); pyautogui.press("enter")

            cam_xy = click_right_of(labels["CAMERA"], dx=120)
            ok_cam = select_from_dropdown_by_text_near(*cam_xy, target_text=str(cam_id), debug_prefix=f"debug_row{idx}_camera")
            if not ok_cam:
                open_dropdown_at(*cam_xy); pyautogui.moveTo(cam_xy[0]+90, cam_xy[1]+30 + (max(cam_id-1,0))*20, duration=0.05); pyautogui.click(); pyautogui.press("enter")

            # Inputs
            frame_xy = click_right_of(labels["FRAME"], dx=140); put_text_at(*frame_xy, row["FRAME"])
            posx_xy  = click_right_of(labels["POSX"],  dx=140); put_text_at(*posx_xy,  row["X_FRAME"])
            posy_xy  = click_right_of(labels["POSY"],  dx=140); put_text_at(*posy_xy,  row["Y_FRAME"])

            # Show
            sh = labels["SHOW"]; sx = sh[0]+sh[2]//2+60; sy = sh[1]+sh[3]//2
            pyautogui.moveTo(sx, sy, duration=0.05); pyautogui.click(); time.sleep(0.05); pyautogui.click()

            # Read result with retries
            X=Y=Z=T=None; last_text=""; tries=3
            for t in range(tries):
                time.sleep(0.3 + float(cfg.get("post_show_delay_s", 0.7)))
                img = find_intersect_text(wrect)
                if idx<5: cv2.imwrite(f"debug_result_row_{idx}_try{t}.png", img)
                X,Y,Z,last_text_T = parse_intersect_from_image(img)[:4]
                last_text = parse_intersect_from_image(img)[4] if len(parse_intersect_from_image(img))>4 else ""
                if X and Y and Z:
                    T = last_text_T
                    break

            df.at[idx,"X_NEW"]=X or ""
            df.at[idx,"Y_NEW"]=Y or ""
            df.at[idx,"Z_NEW"]=Z or ""
            df.at[idx,"TIME_MS"]=T or ""
            print(f"[OK] {row.get(nr_col, idx+1)} -> ({X},{Y},{Z}) {T} ms")
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
