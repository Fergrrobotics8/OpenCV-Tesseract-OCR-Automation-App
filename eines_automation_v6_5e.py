# eines_automation_v6_5e.py
# Paso-a-paso: etapa "mesh" con capturas claras
# - OCR tras maximizar
# - En <=1s abre el desplegable de Mesh
# - Guarda capturas: debug_footer_detect.png, debug_mesh_label_crop.png,
#   debug_mesh_click_overlay.png, debug_mesh_dropdown_open.png
#
# Opciones:
#   --run --stage mesh
#   --run (equivale a stage=full)
#   --dry  (solo imprime mapeo Excel)
#
# Basado en v6_5d (ancho de footer configurable y OCR robusto)

import argparse, os, re, time, json, pandas as pd, xml.etree.ElementTree as ET
from pathlib import Path
from pywinauto.application import Application
from pywinauto import Desktop
import pyautogui, pytesseract, cv2, numpy as np, psutil

pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0.03

DECIMAL_COMMA = True
STATION_ORDER = ["RIGHTSIDE1","RIGHTSIDE2","LEFTSIDE1","LEFTSIDE2","ROOF","ROOFHOOD","HOOD","BRIDGE"]
WIN_RECT = None  # (left, top, right, bottom) cache

# ---------- Utils ----------
def to_float(s):
    if s is None or (isinstance(s, float) and pd.isna(s)): return None
    ss = str(s).strip()
    if DECIMAL_COMMA:
        if ss.count(',')==1 and ss.count('.')>1:  # 1.234,56
            ss = ss.replace('.', '').replace(',', '.')
        else:
            ss = ss.replace(',', '.')
    try: return float(ss)
    except: return None

def norm_name(n): return re.sub(r'[^A-Z0-9]','', str(n).upper())

def pick_col(df, preferred):
    cand = {norm_name(x): x for x in df.columns}
    for p in preferred:
        if p in cand: return cand[p]
    for p in preferred:
        for k,orig in cand.items():
            if p in k: return orig
    return None

# ---------- Excel ----------
def detect_header_row(raw, tokens=None, max_scan=60):
    if tokens is None:
        tokens = ["FRAME","X_FRAME","Y_FRAME","STATION","CAMERA","X_PAG","Y_PAG","Z_PAG",
                  "X_EINES","Y_EINES","Z_EINES","NR","Nº","NUMERO"]
    best_row, best_score = 0, -1
    for i in range(min(max_scan, len(raw))):
        vals = [str(v).strip().upper() for v in raw.iloc[i].tolist()]
        score = sum(1 for t in tokens if t.upper() in vals)
        if score > best_score: best_score, best_row = score, i
    return best_row

def load_excel_any_header(path, sheet_name="Hoja1"):
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None, dtype=str)
    hdr = detect_header_row(raw)
    df  = pd.read_excel(path, sheet_name=sheet_name, header=hdr)
    df  = df.loc[:, ~df.columns.astype(str).str.match(r'Unnamed', case=False)]
    df  = df.dropna(axis=1, how="all")
    for c in ["X_NEW","Y_NEW","Z_NEW","TIME_MS"]:
        if c not in df.columns: df[c] = ""
        df[c] = df[c].astype("object")
    return df, sheet_name, hdr

def load_stations_map(xml_path):
    tree = ET.parse(xml_path); root = tree.getroot()
    m={}
    for st in root.findall("station"):
        try: m[int(st.attrib["id"])]=st.findtext("name")
        except: pass
    return m

# ---------- Window helpers ----------
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
        time.sleep(0.3)
    return None

def attach_or_launch(app_path, title_hint, work_dir=None, attach_only=False, extra_load_wait=1.0):
    """extra_load_wait por defecto 1s para cumplir: clicar mesh ~a 1s de maximizar."""
    global WIN_RECT
    try:
        dlg = Desktop(backend="uia").window(title_re=".*"+re.escape(title_hint)+".*", found_index=0)
        if dlg.exists():
            dlg.set_focus(); time.sleep(extra_load_wait); dlg.maximize()
            r = dlg.rectangle(); WIN_RECT = (r.left, r.top, r.right, r.bottom)
            return dlg
    except Exception: pass
    byproc = _find_window_by_process(os.path.basename(app_path), title_hint, timeout=5)
    if byproc is not None:
        byproc.set_focus(); time.sleep(extra_load_wait)
        try: byproc.maximize()
        except: pass
        r = byproc.rectangle(); WIN_RECT = (r.left, r.top, r.right, r.bottom)
        return byproc
    if attach_only: raise RuntimeError("Window not found to attach.")
    if work_dir is None: work_dir = os.path.dirname(app_path)
    app = Application(backend="uia").start(cmd_line=f'"{app_path}"', work_dir=work_dir)
    time.sleep(2.0)
    t0=time.time(); dlg=None
    while time.time()-t0<40:
        try:
            dlg = Desktop(backend="uia").window(title_re=".*"+re.escape(title_hint)+".*")
            if dlg.exists(): dlg.wait("visible", timeout=5); break
        except Exception: pass
        dlg = _find_window_by_process(os.path.basename(app_path), title_hint, timeout=1)
        if dlg: break
        time.sleep(0.6)
    if dlg is None: raise RuntimeError("Timed out waiting for window")
    dlg.set_focus(); time.sleep(extra_load_wait)
    try: dlg.maximize()
    except: pass
    r = dlg.rectangle(); WIN_RECT = (r.left, r.top, r.right, r.bottom)
    return dlg

def screenshot_region(rect):
    x,y,w,h=rect
    img=pyautogui.screenshot(region=(x,y,w,h))
    return cv2.cvtColor(np.array(img), cv2.COLOR_RGBA2BGR)

def ocr_data(img_bgr, psm=6, timeout=2.0):
    try:
        from pytesseract import image_to_data, Output
        return pytesseract.image_to_data(img_bgr, lang="eng", config=f"--psm {psm}", output_type=Output.DICT, timeout=timeout)
    except Exception:
        return None

def find_footer_labels(footer_scan_extra_up_px=120):
    """Escanea una banda alta para cubrir todos los labels; devuelve labels y el rectángulo analizado."""
    assert WIN_RECT is not None
    left, top, right, bottom = WIN_RECT
    best_labels={}; best_rect=(left, bottom-260-footer_scan_extra_up_px, right-left, 260+footer_scan_extra_up_px); best_score=0
    for h in (260+footer_scan_extra_up_px, 300+footer_scan_extra_up_px):
        for off in range(300+footer_scan_extra_up_px, 140, -12):
            rect=(left, bottom-off, right-left, h)
            bgr=screenshot_region(rect)
            gray=cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
            thr=cv2.adaptiveThreshold(gray,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY,31,7)
            data=ocr_data(thr, psm=6, timeout=1.6)
            labels={}
            if data:
                for i,txt in enumerate(data["text"]):
                    t=(txt or "").strip(); 
                    if not t: continue
                    up=t.upper()
                    for key,wanted in {"MESH":"MESH","STATION":"STATION","CAMERA":"CAMERA","FRAME":"FRAME","POSITIONX":"POSITION X","POSITIONY":"POSITION Y","SHOW":"SHOW"}.items():
                        if wanted in up:
                            x=rect[0]+data["left"][i]; y=rect[1]+data["top"][i]; w=data["width"][i]; h2=data["height"][i]
                            labels[key]=(x,y,w,h2)
            if len(labels)>best_score:
                best_labels, best_rect, best_score = labels, rect, len(labels)
                if best_score>=6: break
        if best_score>=6: break
    labels_final={}
    for k,v in best_labels.items():
        if k=="POSITIONX": labels_final["POSX"]=v
        elif k=="POSITIONY": labels_final["POSY"]=v
        else: labels_final[k]=v
    # fallback aproximado
    fx,fy,fw,fh=best_rect
    approx={"MESH":(fx+40,fy+15,40,18),"STATION":(fx+40,fy+45,60,18),"CAMERA":(fx+40,fy+75,55,18),
            "FRAME":(fx+360,fy+15,45,18),"POSX":(fx+360,fy+45,80,18),"POSY":(fx+360,fy+75,80,18),
            "SHOW":(fx+540,fy+35,45,30)}
    for k,v in approx.items(): labels_final.setdefault(k,v)

    # Debug: imagen del footer + cajas
    dbg=screenshot_region(best_rect)
    for k,(x,y,w,h2) in labels_final.items():
        cv2.rectangle(dbg,(x-best_rect[0],y-best_rect[1]),(x-best_rect[0]+w,y-best_rect[1]+h2),(255,0,0),1)
        cv2.putText(dbg,k,(x-best_rect[0],y-best_rect[1]-2),cv2.FONT_HERSHEY_SIMPLEX,0.45,(255,0,0),1,cv2.LINE_AA)
    cv2.imwrite("debug_footer_detect.png", dbg)

    # Captura recortada de la etiqueta MESH (debug)
    if "MESH" in labels_final:
        x,y,w,h2 = labels_final["MESH"]
        crop_rect = (max(best_rect[0], x-20), max(best_rect[1], y-10),
                     min(w+40, best_rect[2]), min(h2+30, best_rect[3]))
        crop = screenshot_region(crop_rect)
        cv2.imwrite("debug_mesh_label_crop.png", crop)

    return labels_final, best_rect

def click_right_of(label_rect, dx=130):
    x=label_rect[0]+label_rect[2]+dx; y=label_rect[1]+label_rect[3]//2
    pyautogui.moveTo(x,y,duration=0.05); pyautogui.click()
    # Overlay de la pulsación
    roi = (x-20, y-20, 40, 40)
    shot = screenshot_region(roi)
    cv2.circle(shot, (20,20), 10, (0,0,255), 2)
    cv2.imwrite("debug_mesh_click_overlay.png", shot)
    return x,y

def open_dropdown_at(x,y): pyautogui.moveTo(x,y,duration=0.05); pyautogui.click(); time.sleep(0.15)

def select_from_dropdown_by_text_near(x,y, target_text, debug_prefix=None):
    open_dropdown_at(x,y)
    region=(x-140,y+10,560,380)
    bgr=screenshot_region(region)
    if debug_prefix: cv2.imwrite(f"{debug_prefix}_dropdown.png", bgr)
    gray=cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    thr=cv2.adaptiveThreshold(gray,255,cv2.ADAPTIVE_THRESH_MEAN_C,cv2.THRESH_BINARY,31,7)
    data=ocr_data(thr, psm=6, timeout=1.6)
    if data:
        target_up=str(target_text).upper(); best=None
        for i,txt in enumerate(data["text"]):
            t=(txt or "").strip()
            if not t: continue
            if target_up in t.upper():
                best=(data["left"][i],data["top"][i],data["width"][i],data["height"][i]); break
        if best:
            cx=region[0]+best[0]+best[2]//2; cy=region[1]+best[1]+best[3]//2
            pyautogui.moveTo(cx,cy,duration=0.05); pyautogui.click(); pyautogui.press("enter"); return True
    pyautogui.moveTo(region[0]+160,region[1]+34,duration=0.05); pyautogui.click(); pyautogui.press("enter"); return False

# ---------- Main ----------
def main():
    ap=argparse.ArgumentParser()
    ap.add_argument("--config", default="config_v6_5e.json")
    ap.add_argument("--run", action="store_true")
    ap.add_argument("--dry", action="store_true")
    ap.add_argument("--stage", choices=["mesh","full"], default="full")
    args=ap.parse_args()

    cfg=json.load(open(args.config,"r",encoding="utf-8"))
    if cfg.get("tesseract_cmd"): pytesseract.pytesseract.tesseract_cmd = cfg["tesseract_cmd"]

    # Excel (solo para --dry o stage full; para stage mesh no es crítico)
    df, sheet, hdr = load_excel_any_header(cfg["excel_path"], sheet_name=cfg.get("excel_sheet","Hoja1"))
    col_station = pick_col(df, ["STATION","STATIONID","ESTACION"])
    col_camera  = pick_col(df, ["CAMERA","CAM","CAMERAID"])
    col_frame   = pick_col(df, ["FRAME"])
    col_xf      = pick_col(df, ["X_FRAME","XFRAME","XFRM"])
    col_yf      = pick_col(df, ["Y_FRAME","YFRAME","YFRM"])

    if args.dry:
        print("[DRY] Header row index:", hdr)
        print("[DRY] Columns:", list(df.columns))
        print("[DRY] Mapping -> STATION:", col_station, "CAMERA:", col_camera, "FRAME:", col_frame, "X_FRAME:", col_xf, "Y_FRAME:", col_yf)
        print("[DRY] Head:\n", df.head(3))
        return

    # Lanzar visor (espera corta para poder clicar Mesh rápido)
    dlg=attach_or_launch(cfg["viewer_exe_path"], cfg.get("window_title_contains","EINES System 3D for ESPQi"),
                         work_dir=cfg.get("work_dir"), attach_only=bool(cfg.get("attach_only",False)),
                         extra_load_wait=float(cfg.get("extra_load_wait_s",1.0)))
    labels, footer_rect = find_footer_labels(footer_scan_extra_up_px=int(cfg.get("footer_scan_extra_up_px", 140)))

    # --- STAGE: MESH (solo abrir dropdown Mesh y guardar capturas) ---
    if args.stage=="mesh":
        mx,my = click_right_of(labels["MESH"], dx=130)
        # Abrimos y guardamos captura del dropdown
        pyautogui.click(mx,my); time.sleep(0.15)
        dd_region=(mx-140,my+10,560,380)
        dd_img = screenshot_region(dd_region)
        cv2.imwrite("debug_mesh_dropdown_open.png", dd_img)
        # Intento de seleccionar el mesh objetivo (opcional)
        _ = select_from_dropdown_by_text_near(mx,my,cfg.get("mesh_name","TAYCAN_NACHFOLGER"), debug_prefix="debug_mesh")
        print("[STAGE mesh] He clicado el dropdown de Mesh y guardado capturas.")
        return

    # Si no es stage=mesh, aquí seguiría el flujo completo (como v6_5d)
    # Por brevedad, puedes reutilizar v6_5d para 'full'. Aquí te dejo un recordatorio:
    print("[INFO] Para el flujo completo usa v6_5d o ejecuta este archivo con --stage full cuando integre todo el pipeline.")
    return

if __name__=="__main__": main()
