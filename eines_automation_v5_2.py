# eines_automation_v5_2.py
import argparse, os, re, time, json, pandas as pd, xml.etree.ElementTree as ET
from pathlib import Path
try:
    from pywinauto.application import Application
    from pywinauto import Desktop
    import psutil
    import pyautogui, pytesseract, cv2, numpy as np
except Exception as e:
    print("[WARN] Some packages not available here:", e)

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
        try:
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
        except Exception:
            pass
        time.sleep(0.5)
    return None

def attach_or_launch(app_path, title_hint, work_dir=None, attach_only=False, extra_load_wait=3.0):
    # Try attach
    try:
        dlg = Desktop(backend="uia").window(title_re=".*"+re.escape(title_hint)+".*", found_index=0)
        if dlg.exists(): dlg.set_focus(); return dlg
    except Exception:
        pass
    byproc = _find_window_by_process(os.path.basename(app_path), title_hint, timeout=5)
    if byproc is not None: byproc.set_focus(); return byproc
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
    # IDs by captions
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

def click_in_window(dlg):
    r = dlg.rectangle()
    pyautogui.click(r.left+150, r.bottom-120); time.sleep(0.2)

def combo_select_by_index(ctrl, idx):
    """If owner-drawn combo ignores text typing, open and click by index using a reasonable item height."""
    # open dropdown
    try: ctrl.click_input()
    except Exception:
        rect = ctrl.rectangle(); pyautogui.click(rect.mid_point().x, rect.mid_point().y)
    time.sleep(0.15)
    # compute a click under the box
    rect = ctrl.rectangle()
    # try several item heights to account for scaling
    for item_h in (18,20,22,24,26):
        x = rect.left + 80
        y = rect.bottom + 8 + idx*item_h + item_h//2
        pyautogui.click(x, y); time.sleep(0.1)
        pyautogui.press("enter"); time.sleep(0.1)
        # break; we assume selection took
        break

def select_mesh(ctrls, mesh_name):
    if "mesh" in ctrls:
        try:
            ctrls["mesh"].select(mesh_name)
            return
        except Exception:
            pass
        # fallback: typing
        try:
            ctrls["mesh"].click_input(); time.sleep(0.1); pyautogui.typewrite(mesh_name); pyautogui.press("enter"); return
        except Exception:
            pass
    # last resort: heuristic click
    r = ctrls.get("mesh", None).rectangle() if "mesh" in ctrls else Desktop(backend="uia").window(found_index=0).rectangle()
    pyautogui.click(r.left+120, r.bottom-90); time.sleep(0.1); pyautogui.typewrite(mesh_name); pyautogui.press("enter")

def select_station(ctrls, station_name):
    # try native
    try:
        ctrls["station"].select(station_name); return
    except Exception:
        pass
    # try typing
    try:
        ctrls["station"].click_input(); time.sleep(0.1); pyautogui.typewrite(station_name); pyautogui.press("enter"); return
    except Exception:
        pass
    # coordinates by index
    idx = STATION_ORDER.index(station_name) if station_name in STATION_ORDER else 0
    combo_select_by_index(ctrls["station"], idx)

def select_camera(ctrls, cam_id):
    text = str(cam_id)
    # try native
    try:
        ctrls["camera"].select(text); return
    except Exception:
        pass
    # try typing
    try:
        ctrls["camera"].click_input(); time.sleep(0.1); pyautogui.typewrite(text); pyautogui.press("enter"); return
    except Exception:
        pass
    # fallback by index (cameras appear sorted starting at 1)
    idx = max(int(cam_id)-1, 0)
    combo_select_by_index(ctrls["camera"], idx)

def put_text(ctrl, value):
    s=str(value).strip(); s=s.replace('.', ',') if DECIMAL_COMMA else s
    try: ctrl.set_edit_text(s)
    except Exception:
        rect = ctrl.rectangle(); pyautogui.click(rect.mid_point().x, rect.mid_point().y)
        pyautogui.hotkey("ctrl","a"); pyautogui.typewrite(s)

def click_show(ctrl):
    try: ctrl.click_input(); return
    except Exception: pass
    rect = ctrl.rectangle(); pyautogui.click(rect.mid_point().x, rect.mid_point().y)

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
    m = re.search(r"Intersect position:\\s*([-+]?\\d+,\\d+)\\s+([-+]?\\d+,\\d+)\\s+([-+]?\\d+,\\d+)", text)
    X=Y=Z=T=None
    if m: X,Y,Z = m.group(1), m.group(2), m.group(3)
    m2 = re.search(r"([\\d,]+)\\s*ms", text)
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
    ap.add_argument("--config", default="config_v5_2.json")
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
    # click to focus
    r=dlg.rectangle(); pyautogui.click(r.left+150, r.bottom-120); time.sleep(0.2)

    ctrls = find_controls(dlg)
    # Mesh
    select_mesh(ctrls, mesh_name)
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

            if "show" in ctrls: 
                try: ctrls["show"].click_input()
                except Exception: 
                    rect = ctrls["show"].rectangle(); pyautogui.click(rect.mid_point().x, rect.mid_point().y)
            else:
                pyautogui.press("enter")

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
