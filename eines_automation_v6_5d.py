# eines_automation_v6_5d.py
# Mejoras vs v6_5c:
# - Búsqueda del footer más ALTA y más ALTA (región más grande) para no cortar "Mesh/Station/Camera"
# - Cachea el rectángulo de la ventana (WIN_RECT) y evita llamadas repetidas a dlg.rectangle()
# - Captura del bloque de resultados ANCLADA a "Intersect position:" (OCR) o relativa al botón SHOW si no se encuentra
# - Más robustez en dropdowns (más reintentos y margen)
# - Parámetros de ajuste extra en config (footer_scan_extra_up_px, result_scan_height)

import argparse, os, re, time, json, pandas as pd, xml.etree.ElementTree as ET
from pathlib import Path
from pywinauto.application import Application
from pywinauto import Desktop
import pyautogui, pytesseract, cv2, numpy as np, psutil, sys

pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0.03

DECIMAL_COMMA = True
STATION_ORDER = ["RIGHTSIDE1","RIGHTSIDE2","LEFTSIDE1","LEFTSIDE2","ROOF","ROOFHOOD","HOOD","BRIDGE"]

WIN_RECT = None  # (left, top, right, bottom) cache global

def to_float(s):
    if s is None or (isinstance(s, float) and pd.isna(s)): return None
    ss = str(s).strip()
    if DECIMAL_COMMA:
        if ss.count(',')==1 and ss.count('.')>1:  # 1.234,56 -> 1234.56
            ss = ss.replace('.', '').replace(',', '.')
        else:
            ss = ss.replace(',', '.')
    try: return float(ss)
    except: return None

def norm_name(n):
    return re.sub(r'[^A-Z0-9]','', str(n).upper())

def pick_col(df, preferred):
    cand = {norm_name(x): x for x in df.columns}
    for p in preferred:
        if p in cand: return cand[p]
    for p in preferred:
        for k,orig in cand.items():
            if p in k: return orig
    return None

# -------- Excel helpers --------
def detect_header_row(raw, tokens=None, max_scan=60):
    if tokens is None:
        tokens = ["FRAME","X_FRAME","Y_FRAME","STATION","CAMERA","X_PAG","Y_PAG","Z_PAG","X_EINES","Y_EINES","Z_EINES",
                  "HIST_X_MESH","HIST_Y_MESH","HIST_Z_MESH","PAG_X_MESH","PAG_Y_MESH","PAG_Z_MESH","NR","Nº","NUMERO"]
    best_row, best_score = 0, -1
    for i in range(min(max_scan, len(raw))):
        vals = [str(v).strip().upper() for v in raw.iloc[i].tolist()]
        score = sum(1 for t in tokens if t.upper() in vals)
        if score > best_score:
            best_score = score
            best_row = i
    return best_row

def load_excel_any_header(path, sheet_name="Hoja1"):
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None, dtype=str)
    hdr = detect_header_row(raw)
    df = pd.read_excel(path, sheet_name=sheet_name, header=hdr)
    # drop Unnamed/empty columns
    df = df.loc[:, ~df.columns.astype(str).str.match(r'Unnamed', case=False)]
    df = df.dropna(axis=1, how="all")
    # ensure output cols
    for c in ["X_NEW","Y_NEW","Z_NEW","TIME_MS"]:
        if c not in df.columns: df[c] = ""
        df[c] = df[c].astype("object")
    return df, sheet_name, hdr

# -------- Stations map --------
def load_stations_map(xml_path):
    tree = ET.parse(xml_path); root = tree.getroot()
    m={}
    for st in root.findall("station"):
        try: m[int(st.attrib["id"])]=st.findtext("name")
        except: pass
    return m

# -------- Window helpers --------
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
        time.sleep(0.4)
    return None

def attach_or_launch(app_path, title_hint, work_dir=None, attach_only=False, extra_load_wait=3.0):
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
    time.sleep(3.0)
    t0=time.time(); dlg=None
    while time.time()-t0<60:
        try:
            dlg = Desktop(backend="uia").window(title_re=".*"+re.escape(title_hint)+".*")
            if dlg.exists(): dlg.wait("visible", timeout=5); break
        except Exception: pass
        dlg = _find_window_by_process(os.path.basename(app_path), title_hint, timeout=1)
        if dlg: break
        time.sleep(0.8)
    if dlg is None: raise RuntimeError("Timed out waiting for window")
    dlg.set_focus(); time.sleep(extra_load_wait)
    try: dlg.maximize()
    except: pass
    r = dlg.rectangle(); WIN_RECT = (r.left, r.top, r.right, r.bottom)
    return dlg

def screenshot_region(rect):
    x, y, w, h = rect
    # Asegura que todos los valores sean enteros
    x, y, w, h = int(x), int(y), int(w), int(h)
    img = pyautogui.screenshot(region=(x, y, w, h))
    return cv2.cvtColor(np.array(img), cv2.COLOR_RGBA2BGR)

def ocr_data(img_bgr, psm=6, timeout=2.0):
    try:
        from pytesseract import image_to_data, Output
        return pytesseract.image_to_data(img_bgr, lang="eng", config=f"--psm {psm}", output_type=Output.DICT, timeout=timeout)
    except Exception:
        return None

def find_footer_labels(footer_scan_extra_up_px=60):
    """Escanea una banda más ALTA que antes para cubrir todos los labels."""
    assert WIN_RECT is not None
    left, top, right, bottom = WIN_RECT
    best_labels={}; best_rect=(left, bottom-220-footer_scan_extra_up_px, right-left, 220+footer_scan_extra_up_px); best_score=0
    # probamos varias alturas/desplazamientos
    for h in (240+footer_scan_extra_up_px, 260+footer_scan_extra_up_px, 200+footer_scan_extra_up_px):
        for off in range(260+footer_scan_extra_up_px, 140, -12):
            rect=(left, bottom-off, right-left, h)
            bgr=screenshot_region(rect)
            gray=cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
            thr=cv2.adaptiveThreshold(gray,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY,31,7)
            data=ocr_data(thr, psm=6, timeout=1.8)
            labels={}
            if data:
                for i,txt in enumerate(data["text"]):
                    t=(txt or "").strip()
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
    # convertir a final con aproximaciones
    labels_final={}
    for k,v in best_labels.items():
        if k=="POSITIONX": labels_final["POSX"]=v
        elif k=="POSITIONY": labels_final["POSY"]=v
        else: labels_final[k]=v
    fx,fy,fw,fh=best_rect
    approx={"MESH":(fx+40,fy+15,40,18),"STATION":(fx+40,fy+45,60,18),"CAMERA":(fx+40,fy+75,55,18),
            "FRAME":(fx+360,fy+15,45,18),"POSX":(fx+360,fy+45,80,18),"POSY":(fx+360,fy+75,80,18),
            "SHOW":(fx+540,fy+35,45,30)}
    for k,v in approx.items(): labels_final.setdefault(k,v)
    # debug
    dbg=screenshot_region(best_rect)
    for k,(x,y,w,h2) in labels_final.items():
        cv2.rectangle(dbg,(x-best_rect[0],y-best_rect[1]),(x-best_rect[0]+w,y-best_rect[1]+h2),(255,0,0),1)
        cv2.putText(dbg,k,(x-best_rect[0],y-best_rect[1]-2),cv2.FONT_HERSHEY_SIMPLEX,0.45,(255,0,0),1,cv2.LINE_AA)
    cv2.imwrite("debug_footer_detect.png", dbg)
    return labels_final, best_rect

def click_right_of(label_rect, dx=130):
    x=label_rect[0]+label_rect[2]+dx; y=label_rect[1]+label_rect[3]//2
    pyautogui.moveTo(x,y,duration=0.05); pyautogui.click()
    return x,y

def open_dropdown_at(x,y): pyautogui.moveTo(x,y,duration=0.05); pyautogui.click(); time.sleep(0.15)

def select_from_dropdown_by_text_near(x,y, target_text, debug_prefix=None):
    open_dropdown_at(x,y)
    region=(x-140,y+10,560,380)
    bgr=screenshot_region(region)
    if debug_prefix: cv2.imwrite(f"{debug_prefix}_dropdown.png", bgr)
    gray=cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    thr=cv2.adaptiveThreshold(gray,255,cv2.ADAPTIVE_THRESH_MEAN_C,cv2.THRESH_BINARY,31,7)
    data=ocr_data(thr, psm=6, timeout=1.8)
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

def put_text_at(x,y,val):
    s=str(val).strip(); s=s.replace('.', ',') if DECIMAL_COMMA else s
    pyautogui.moveTo(x,y,duration=0.05); pyautogui.click(); pyautogui.hotkey("ctrl","a"); pyautogui.typewrite(s)

def locate_result_block(labels, result_scan_height=170):
    """Intenta localizar 'Intersect position:' por OCR en la parte derecha inferior."""
    assert WIN_RECT is not None
    left, top, right, bottom = WIN_RECT
    search_rect=(int((left+right)*0.55), bottom-result_scan_height-60, right-int((right-left)*0.02), result_scan_height)
    bgr=screenshot_region(search_rect)
    gray=cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    thr=cv2.adaptiveThreshold(gray,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY,31,7)
    data=ocr_data(thr, psm=6, timeout=1.8)
    if data:
        for i,txt in enumerate(data["text"]):
            t=(txt or "").strip()
            if not t: continue
            if "INTERSECT POSITION" in t.upper():
                y = search_rect[1] + max(0, data["top"][i]-6)
                return (search_rect[0], y, search_rect[2]-search_rect[0], min(result_scan_height, search_rect[3]- (y-search_rect[1])))
    # Fallback: a la derecha del botón SHOW
    if "SHOW" in labels:
        sx, sy, w, h = labels["SHOW"]
        rect = (sx + w + 40, sy - 20, max(420, int((WIN_RECT[2] - WIN_RECT[0]) * 0.35)), result_scan_height)
    else:
        # Usa una posición por defecto segura en la parte derecha inferior
        left, top, right, bottom = WIN_RECT
        rect = (int((left + right) * 0.7), int(bottom - result_scan_height - 40), int((right - left) * 0.25), result_scan_height)
    return rect

def capture_result_block(rect):
    return screenshot_region(rect)

def parse_intersect(text):
    m=re.search(r"Intersect position:\s*([-+]?\d+,\d+)\s+([-+]?\d+,\d+)\s+([-+]?\d+,\d+)", text, re.IGNORECASE)
    X=Y=Z=T=None
    if m: X,Y,Z=m.group(1),m.group(2),m.group(3)
    m2=re.search(r"([\d,]+)\s*ms", text)
    if m2: T=m2.group(1)
    return X,Y,Z,T

def find_footer_by_color(win_rect):
    """Detecta la separación entre zona gris y blanca en la parte baja de la ventana."""
    left, top, right, bottom = win_rect
    img = screenshot_region((left, top, right-left, bottom-top))
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # Promedia columnas para detectar el salto de color
    mean_col = np.mean(gray, axis=1)
    # Busca el mayor gradiente (cambio brusco)
    diff = np.abs(np.diff(mean_col))
    sep_idx = np.argmax(diff[::-1])  # desde abajo hacia arriba
    sep_y = bottom - sep_idx
    # Debug visual
    dbg = img.copy()
    cv2.line(dbg, (0, sep_y-top), (right-left, sep_y-top), (0,0,255), 2)
    cv2.imwrite("debug_footer_separator.png", dbg)
    return sep_y

def find_labels_and_boxes(win_rect, sep_y):
    """Detecta los textos y las cajas a la derecha en la zona blanca."""
    left, top, right, bottom = win_rect
    # Asegura enteros en la región
    footer_img = screenshot_region((int(left), int(sep_y), int(right-left), int(bottom-sep_y)))
    gray = cv2.cvtColor(footer_img, cv2.COLOR_BGR2GRAY)
    thr = cv2.adaptiveThreshold(gray,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY,31,7)
    data = ocr_data(thr, psm=6, timeout=2.0)
    labels = {}
    if data:
        for i, txt in enumerate(data["text"]):
            t = (txt or "").strip()
            if not t: continue
            up = t.upper()
            for key in ["MESH","STATION","CAMERA","FRAME","POSITION X","POSITION Y","SHOW"]:
                if key in up:
                    x, y, w, h = int(data["left"][i]), int(data["top"][i]), int(data["width"][i]), int(data["height"][i])
                    labels[key] = (x, y, w, h)
    # Busca cajas a la derecha de cada label
    boxes = []
    contours, _ = cv2.findContours(255-thr, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        if w > 60 and h > 18:  # tamaño típico de caja
            boxes.append((int(x), int(y), int(w), int(h)))
    # Asocia cada label con la caja más cercana a su derecha
    label_boxes = {}
    for key, (lx, ly, lw, lh) in labels.items():
        candidates = [b for b in boxes if b[0] > lx+lw and abs(b[1]-ly)<20]
        if candidates:
            bx = min(candidates, key=lambda b: b[0])
            # Centro de la caja
            cx = bx[0] + bx[2]//2
            cy = bx[1] + bx[3]//2
            label_boxes[key] = (int(left)+cx, int(sep_y)+cy, bx)
    # Debug visual
    dbg = footer_img.copy()
    for key, (lx, ly, lw, lh) in labels.items():
        cv2.rectangle(dbg, (lx, ly), (lx+lw, ly+lh), (255,0,0), 1)
        cv2.putText(dbg, key, (lx, ly-2), cv2.FONT_HERSHEY_SIMPLEX, 0.45, (255,0,0), 1, cv2.LINE_AA)
    for (x, y, w, h) in boxes:
        cv2.rectangle(dbg, (x, y), (x+w, y+h), (0,255,0), 1)
    cv2.imwrite("debug_footer_labels_boxes.png", dbg)
    return label_boxes

def find_footer_separator_and_labels(win_rect):
    """Detecta la separación gris/blanco usando OCR de 'Mesh' y busca la línea de cambio de color encima."""
    left, top, right, bottom = win_rect
    img = screenshot_region((int(left), int(top), int(right-left), int(bottom-top)))
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # 1. OCR de toda la pantalla para encontrar 'Mesh'
    data = ocr_data(gray, psm=6, timeout=2.0)
    mesh_y = None
    if data:
        for i, txt in enumerate(data["text"]):
            t = (txt or "").strip().upper()
            if "MESH" in t:
                mesh_y = int(data["top"][i] + data["height"][i]//2)
                break
    if mesh_y is None:
        # Fallback: usa el último 30% de la pantalla
        mesh_y = int(gray.shape[0]*0.7)

    # 2. Desde mesh_y, sube y busca el cambio de blanco a gris
    search_top = max(mesh_y - 120, 0)
    col = np.mean(gray[search_top:mesh_y], axis=1)
    diff = np.abs(np.diff(col))
    if len(diff) > 0:
        sep_offset = np.argmax(diff[::-1])
        sep_y = mesh_y - sep_offset
    else:
        sep_y = mesh_y

    sep_y_abs = int(top + sep_y)
    dbg = img.copy()
    cv2.line(dbg, (0, sep_y), (int(right-left), sep_y), (0,0,255), 2)
    cv2.imwrite("debug_footer_separator.png", dbg)

    # 3. Zona blanca (footer)
    footer_img = screenshot_region((int(left), int(sep_y_abs), int(right-left), int(bottom-sep_y_abs)))
    cv2.imwrite("debug_footer_img_for_ocr.png", footer_img)

    # Preprocesado solo con gray_footer (sin contrasted)
    gray_footer = cv2.cvtColor(footer_img, cv2.COLOR_BGR2GRAY)
    scales = [1.0, 1.5, 2.0]
    psms = [6, 7, 11, 3]
    for scale in scales:
        if scale != 1.0:
            img = cv2.resize(gray_footer, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
        else:
            img = gray_footer
        for psm in psms:
            data = ocr_data(img, psm=psm, timeout=2.0)
            dbg_labels = cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
            found = False
            if data:
                for i, txt in enumerate(data["text"]):
                    t = (txt or "").strip()
                    if not t: continue
                    up = t.upper()
                    for key in ["MESH","STATION","CAMERA","FRAME","POSITION X","POSITION Y","SHOW"]:
                        if key in up:
                            found = True
                            x, y, w, h = int(data["left"][i]), int(data["top"][i]), int(data["width"][i]), int(data["height"][i])
                            cv2.rectangle(dbg_labels, (x, y), (x+w, y+h), (0, 0, 255), 2)
                            cv2.putText(dbg_labels, key, (x, y-5), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0,0,255), 1, cv2.LINE_AA)
            cv2.imwrite(f"debug_footer_labels_texts_psm{psm}_scale{scale}.png", dbg_labels)
            # Guarda el texto bruto detectado
            with open(f"debug_footer_ocr_raw_psm{psm}_scale{scale}.txt", "w", encoding="utf-8") as f:
                if data:
                    for txt in data["text"]:
                        f.write(str(txt) + "\n")
    # Busca cajas a la derecha de cada label
    boxes = []
    contours, _ = cv2.findContours(255-thr, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        if w > 60 and h > 18:
            boxes.append((int(x), int(y), int(w), int(h)))
    label_boxes = {}
    for key, (lx, ly, lw, lh) in labels.items():
        candidates = [b for b in boxes if b[0] > lx+lw and abs(b[1]-ly)<20]
        if candidates:
            bx = min(candidates, key=lambda b: b[0])
            cx = bx[0] + bx[2]//2
            cy = bx[1] + bx[3]//2
            label_boxes[key] = (int(left)+cx, int(sep_y_abs)+cy, bx)
    # Debug visual
    dbg2 = footer_img.copy()
    for key, (lx, ly, lw, lh) in labels.items():
        cv2.rectangle(dbg2, (lx, ly), (lx+lw, ly+lh), (255,0,0), 1)
        cv2.putText(dbg2, key, (lx, ly-2), cv2.FONT_HERSHEY_SIMPLEX, 0.45, (255,0,0), 1, cv2.LINE_AA)
    for (x, y, w, h) in boxes:
        cv2.rectangle(dbg2, (x, y), (x+w, y+h), (0,255,0), 1)
    cv2.imwrite("debug_footer_labels_boxes.png", dbg2)
    return label_boxes

def select_from_dropdown_by_text_robust(x, y, target_text, debug_prefix=None):
    """Abre el desplegable y busca el texto cerca de la caja clicada."""
    open_dropdown_at(x, y)
    time.sleep(0.5)
    # Región: debajo de la caja, suficientemente ancha y alta
    region = (int(x - 120), int(y + 10), 400, 350)
    bgr = screenshot_region(region)
    if debug_prefix: cv2.imwrite(f"{debug_prefix}_dropdown.png", bgr)
    gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    thr = cv2.adaptiveThreshold(gray,255,cv2.ADAPTIVE_THRESH_MEAN_C,cv2.THRESH_BINARY,31,7)
    data = ocr_data(thr, psm=6, timeout=2.0)
    if data:
        target_up = str(target_text).upper()
        for i, txt in enumerate(data["text"]):
            t = (txt or "").strip()
            if not t: continue
            if target_up in t.upper():
                bx, by, bw, bh = int(data["left"][i]), int(data["top"][i]), int(data["width"][i]), int(data["height"][i])
                cx = region[0] + bx + bw//2
                cy = region[1] + by + bh//2
                pyautogui.moveTo(cx, cy, duration=0.05)
                pyautogui.click()
                pyautogui.press("enter")
                return True
    # Si no encuentra, clic por defecto
    pyautogui.moveTo(region[0]+160, region[1]+34, duration=0.05)
    pyautogui.click()
    pyautogui.press("enter")
    return False

def load_stations_and_cameras(stations_xml, cameras_xml):
    """Carga el mapeo de estaciones y cámaras desde los XML."""
    stations = {}
    cameras = {}
    tree = ET.parse(stations_xml)
    for st in tree.getroot().findall("station"):
        try:
            st_id = int(st.attrib["id"])
            st_name = st.findtext("name")
            stations[st_id] = st_name
        except Exception:
            continue
    if os.path.exists(cameras_xml):
        tree = ET.parse(cameras_xml)
        for cam in tree.getroot().findall("camera"):
            try:
                cam_id = int(cam.attrib["id"])
                subset = int(cam.findtext("subset"))
                cameras.setdefault(subset, []).append(cam_id)
            except Exception:
                continue
    return stations, cameras

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", default="config_v6_5e.json")
    ap.add_argument("--run", action="store_true")
    ap.add_argument("--dry", action="store_true")
    # Si no hay argumentos, añade los que quieras por defecto
    if len(sys.argv) == 1:
        sys.argv += ["--run", "--config", "config_v6_5e.json"]
    args = ap.parse_args()

    cfg = json.load(open(args.config, "r", encoding="utf-8"))
    if cfg.get("tesseract_cmd"):
        pytesseract.pytesseract.tesseract_cmd = cfg["tesseract_cmd"]

    df, sheet, hdr = load_excel_any_header(cfg["excel_path"], sheet_name=cfg.get("excel_sheet", "Hoja1"))

    col_station = pick_col(df, ["STATION", "STATIONID", "ESTACION"])
    col_camera = pick_col(df, ["CAMERA", "CAM", "CAMERAID"])
    col_frame = pick_col(df, ["FRAME"])
    col_xf = pick_col(df, ["X_FRAME", "XFRAME", "XFRM"])
    col_yf = pick_col(df, ["Y_FRAME", "YFRAME", "YFRM"])

    if args.dry:
        print("[DRY] Header row index:", hdr)
        print("[DRY] Columns:", list(df.columns))
        print("[DRY] Mapping -> STATION:", col_station, "CAMERA:", col_camera, "FRAME:", col_frame, "X_FRAME:", col_xf, "Y_FRAME:", col_yf)
        print("[DRY] Head:\n", df.head(3))
        return

    if not all([col_station, col_camera, col_frame, col_xf, col_yf]):
        print("[ERR] No encuentro columnas. Tengo:", list(df.columns))
        return

    stations, cameras = load_stations_and_cameras(cfg["stations_xml_path"], cfg.get("cameras_xml_path", "cameras.xml"))

    # Launch viewer
    dlg = attach_or_launch(cfg["viewer_exe_path"], cfg.get("window_title_contains", "EINES System 3D for ESPQi"),
                           work_dir=cfg.get("work_dir"), attach_only=bool(cfg.get("attach_only", False)),
                           extra_load_wait=float(cfg.get("extra_load_wait_s", 3.0)))

    # Detección robusta de footer y labels/cajas
    label_boxes = find_footer_separator_and_labels(WIN_RECT)

    # Selecciona Mesh solo una vez
    if "MESH" in label_boxes:
        mx, my, _ = label_boxes["MESH"]
        ok = select_from_dropdown_by_text_robust(mx, my, cfg.get("mesh_name", "TAYCAN_NACHFOLGER"), debug_prefix="debug_mesh")
        if not ok:
            print("[WARN] No se pudo seleccionar Mesh.")
        time.sleep(0.3)

    # Cachea la zona de resultados una sola vez
    if "SHOW" in label_boxes:
        sh = label_boxes["SHOW"]
        sx2 = sh[0] + 70
        sy2 = sh[1]
        result_rect = locate_result_block({"SHOW": (sx2, sy2, 45, 30)}, result_scan_height=int(cfg.get("result_scan_height", 170)))
    else:
        result_rect = locate_result_block({}, result_scan_height=int(cfg.get("result_scan_height", 170)))

    for idx, row in df.iterrows():
        try:
            st_id = int(to_float(row[col_station]) or row[col_station])
            cam_id = int(to_float(row[col_camera]) or row[col_camera])
            st_name = stations.get(st_id, None)
            if not st_name:
                print(f"[WARN] Row {idx}: unknown station {st_id}")
                continue

            # Selecciona estación
            if "STATION" in label_boxes:
                sx, sy, _ = label_boxes["STATION"]
                ok_st = select_from_dropdown_by_text_robust(sx, sy, st_name, debug_prefix=f"debug_row{idx}_station")
                if not ok_st:
                    print(f"[WARN] Row {idx}: no se pudo seleccionar estación {st_name}")

            # Selecciona cámara
            if "CAMERA" in label_boxes:
                cx, cy, _ = label_boxes["CAMERA"]
                ok_cam = select_from_dropdown_by_text_robust(cx, cy, str(cam_id), debug_prefix=f"debug_row{idx}_camera")
                if not ok_cam:
                    print(f"[WARN] Row {idx}: no se pudo seleccionar cámara {cam_id}")

            # Frame y posiciones
            if "FRAME" in label_boxes:
                fx, fy, _ = label_boxes["FRAME"]
                put_text_at(fx, fy, row[col_frame])
            if "POSITION X" in label_boxes:
                px, py_, _ = label_boxes["POSITION X"]
                put_text_at(px, py_, row[col_xf])
            if "POSITION Y" in label_boxes:
                pyx, pyy, _ = label_boxes["POSITION Y"]
                put_text_at(pyx, pyy, row[col_yf])

            # Click SHOW
            if "SHOW" in label_boxes:
                shx, shy, _ = label_boxes["SHOW"]
                pyautogui.moveTo(shx, shy, duration=0.05)
                pyautogui.click()
                time.sleep(0.05)
                pyautogui.click()

            X = Y = Z = T = None
            for t in range(3):
                time.sleep(0.3 + float(cfg.get("post_show_delay_s", 0.7)))
                img = capture_result_block(result_rect)
                gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
                thr = cv2.adaptiveThreshold(gray,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY,31,7)
                text = pytesseract.image_to_string(thr, lang="eng", config="--psm 6", timeout=1.8)
                X, Y, Z, T = parse_intersect(text)
                if X and Y and Z:
                    break
                if idx < 5:
                    cv2.imwrite(f"debug_result_row_{idx}_try{t}.png", img)

            df.at[idx, "X_NEW"] = X or ""
            df.at[idx, "Y_NEW"] = Y or ""
            df.at[idx, "Z_NEW"] = Z or ""
            df.at[idx, "TIME_MS"] = T or ""
            print(f"[OK] row {idx+1} -> ({X},{Y},{Z}) {T} ms")
        except Exception as e:
            print(f"[ERR] Row {idx}: {e}")

    out = cfg.get("output_excel_path") or cfg["excel_path"]
    try:
        with pd.ExcelWriter(out, engine="openpyxl", mode="w") as w:
            df.to_excel(w, sheet_name=cfg.get("excel_sheet", "Hoja1"), index=False)
        print(f"[DONE] Saved to {out}")
    except PermissionError:
        alt = str(Path(out).with_name(Path(out).stem + "_out.xlsx"))
        with pd.ExcelWriter(alt, engine="openpyxl", mode="w") as w:
            df.to_excel(w, sheet_name=cfg.get("excel_sheet", "Hoja1"), index=False)
        print(f"[WARN] Excel locked; saved to {alt}")

if __name__ == "__main__":
    main()
