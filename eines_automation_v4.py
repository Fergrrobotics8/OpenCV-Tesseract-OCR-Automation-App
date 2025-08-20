# eines_automation_v4.py
import argparse, os, re, time, json, pandas as pd, xml.etree.ElementTree as ET, sys
from pathlib import Path
try:
    from pywinauto.application import Application
    from pywinauto import Desktop
    from pywinauto.controls.uiawrapper import UIAWrapper
    import pyautogui, pytesseract, cv2, numpy as np
except Exception as e:
    print("[WARN] Some automation libs not available in this environment:", e)

DECIMAL_COMMA = True

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

def attach_or_launch(app_path, title_hint, work_dir=None, attach_only=False):
    try:
        dlg = Desktop(backend="uia").window(title_re=".*"+re.escape(title_hint)+".*", found_index=0)
        if dlg.exists(): dlg.set_focus(); return dlg
    except Exception: pass
    if attach_only:
        raise RuntimeError("Viewer window not found to attach. Open it manually and try again.")
    if work_dir is None: work_dir = os.path.dirname(app_path)
    app = Application(backend="uia").start(cmd_line=f'"{app_path}"', work_dir=work_dir)
    time.sleep(3.0)
    dlg = Desktop(backend="uia").window(title_re=".*"+re.escape(title_hint)+".*")
    dlg.wait("visible", timeout=25)
    dlg.set_focus()
    return dlg

def dump_ui_tree(dlg, out_txt="ui_tree.txt", screenshot="ui_screenshot.png"):
    try:
        with open(out_txt,"w",encoding="utf-8") as f:
            f.write(dlg.print_control_identifiers())
    except Exception:
        pass
    try:
        r = dlg.rectangle()
        img = pyautogui.screenshot(region=(r.left, r.top, r.width(), r.height()))
        img.save(screenshot)
    except Exception:
        pass

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

def wait_enable(ctrl, timeout=20.0):
    t0=time.time()
    while time.time()-t0<timeout:
        try:
            if ctrl.is_enabled(): return True
        except: pass
        time.sleep(0.3)
    return False

def initial_focus_click(dlg):
    r = dlg.rectangle()
    pyautogui.click(r.left+200, r.bottom-120)  # sobre la franja inferior
    time.sleep(0.2)

def fallback_click_mesh(dlg, mesh_name):
    # Heurística: en la barra inferior izquierda, el primer combobox es Mesh.
    r = dlg.rectangle()
    x = r.left + 120
    y = r.bottom - 90
    pyautogui.click(x, y)    # abre el combo
    time.sleep(0.3)
    pyautogui.typewrite(mesh_name)
    pyautogui.press("enter")

def fallback_click_station(dlg, station_name):
    r = dlg.rectangle()
    x = r.left + 120
    y = r.bottom - 60
    pyautogui.click(x, y)
    time.sleep(0.2)
    pyautogui.typewrite(station_name)
    pyautogui.press("enter")

def fallback_click_camera(dlg, cam_id):
    r = dlg.rectangle()
    x = r.left + 120
    y = r.bottom - 30
    pyautogui.click(x, y)
    time.sleep(0.2)
    pyautogui.typewrite(str(cam_id))
    pyautogui.press("enter")

def type_into_edit(ctrl, value):
    s = str(value).strip()
    s = s.replace('.', ',') if DECIMAL_COMMA else s
    try:
        ctrl.set_edit_text(s); return
    except Exception:
        pass
    rect = ctrl.rectangle()
    pyautogui.click(rect.mid_point().x, rect.mid_point().y)
    pyautogui.hotkey("ctrl","a"); pyautogui.typewrite(s)

def click_show(ctrl):
    try: ctrl.click(); return
    except Exception: pass
    rect = ctrl.rectangle()
    pyautogui.click(rect.mid_point().x, rect.mid_point().y)

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

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument("--config", default="config_v4.json")
    ap.add_argument("--run", action="store_true")
    ap.add_argument("--dry", action="store_true")
    args=ap.parse_args()

    cfg=json.load(open(args.config,"r",encoding="utf-8"))
    excel=cfg["excel_path"]; xmlp=cfg["stations_xml_path"]; app=cfg["viewer_exe_path"]
    title=cfg.get("window_title_contains","eines.system3d")
    mesh_name=cfg.get("mesh_name","TAYCAN_NACHFOLGER")
    work_dir=cfg.get("work_dir", os.path.dirname(app))
    attach_only=bool(cfg.get("attach_only", False))

    df,sheet,_=load_excel(excel)
    st_by_id,_=load_stations_map(xmlp)

    if args.dry:
        print(f"Excel '{excel}' sheet '{sheet}' rows={len(df)}")
        print(f"Stations: {st_by_id}")
        return

    dlg = attach_or_launch(app, title, work_dir=work_dir, attach_only=attach_only)
    dump_ui_tree(dlg, out_txt="ui_tree.txt", screenshot="ui_screenshot.png")
    initial_focus_click(dlg)

    ctrls = find_controls(dlg)

    # Select MESH (UIA first, else fallback by coordinates)
    if "mesh" in ctrls:
        try:
            ctrls["mesh"].select(mesh_name)
        except Exception:
            fallback_click_mesh(dlg, mesh_name)
    else:
        fallback_click_mesh(dlg, mesh_name)
    time.sleep(0.6)

    # Ensure Station/Camera
    if "station" not in ctrls or "camera" not in ctrls:
        # maybe they appear only after choosing mesh; refresh search
        ctrls = find_controls(dlg)

    nr_col = next((c for c in df.columns if c.startswith("NR")), df.columns[0])

    for idx,row in df.iterrows():
        st_id=int(to_float(row["STATION"]) or row["STATION"]); cam_id=int(to_float(row["CAMERA"]) or row["CAMERA"])
        st_name = st_by_id.get(st_id, None)
        if not st_name:
            print(f"[WARN] Row {idx}: unknown station id {st_id}"); continue

        # Select Station
        if "station" in ctrls:
            try: ctrls["station"].select(st_name)
            except Exception: fallback_click_station(dlg, st_name)
        else:
            fallback_click_station(dlg, st_name)

        time.sleep(0.2)

        # Select Camera
        if "camera" in ctrls:
            try: ctrls["camera"].select(str(cam_id))
            except Exception: fallback_click_camera(dlg, cam_id)
        else:
            fallback_click_camera(dlg, cam_id)

        time.sleep(0.2)

        # Fill edits
        for key,val in [("frame",row["FRAME"]),("posx",row["X_FRAME"]),("posy",row["Y_FRAME"])]:
            if key in ctrls:
                type_into_edit(ctrls[key], val)
            else:
                print(f"[WARN] Edit '{key}' not found; skipping this row.")
        time.sleep(0.2)

        # Show
        if "show" in ctrls: click_show(ctrls["show"])
        else: print("[WARN] Show button not found; attempting Enter"); pyautogui.press("enter")

        time.sleep(0.3 + float(cfg.get("post_show_delay_s", 0.7)))

        text = ocr_intersect(dlg, ctrls.get("intersect_text"))
        X,Y,Z,T,_ = parse_intersect(text)
        df.at[idx,"X_NEW"]=X or ""; df.at[idx,"Y_NEW"]=Y or ""; df.at[idx,"Z_NEW"]=Z or ""; df.at[idx,"TIME_MS"]=T or ""
        print(f"[OK] {row.get(nr_col, idx+1)} -> ({X},{Y},{Z}) {T} ms")

        time.sleep(float(cfg.get("per_row_pause_s", 0.2)))

    out=cfg.get("output_excel_path") or excel
    with pd.ExcelWriter(out, engine="openpyxl", mode="w") as w: df.to_excel(w, sheet_name=sheet, index=False)
    print(f"[DONE] Saved to {out}")

if __name__=="__main__": main()
