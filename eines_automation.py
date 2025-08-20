# eines_automation.py
import argparse, os, re, time, json, pandas as pd, xml.etree.ElementTree as ET
from pathlib import Path
try:
    from pywinauto.application import Application
    from pywinauto import Desktop
    import pyautogui, pytesseract, cv2, numpy as np
except Exception:
    pass

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
        if "X_FRAME" in row_vals and "Y_FRAME" in row_vals and "FRAME" in row_vals:
            header_row = i; break
        if "X_PAG" in row_vals and "X_EINES" in row_vals:
            header_row = i
    if header_row is None: header_row = 0
    df = pd.read_excel(path, sheet_name=sheet, header=header_row)
    df.columns = [str(c).strip().upper().replace(" ", "_") for c in df.columns]
    return df, sheet, header_row

def ensure_output_columns(df):
    for c in ["X_NEW","Y_NEW","Z_NEW","TIME_MS"]:
        if c not in df.columns: df[c] = ""
    return df

def load_stations_map(xml_path):
    tree = ET.parse(xml_path); root = tree.getroot()
    station_by_id, cams_by_station_id = {}, {}
    for st in root.findall("station"):
        sid = int(st.attrib["id"]); name = st.findtext("name")
        station_by_id[sid] = name; cams = []
        for cam in st.find("cameras").findall("camera"):
            cams.append({"id": int(cam.attrib["id"]),
                        "target": int(cam.findtext("target")),
                        "subset": int(cam.findtext("subset"))})
        cams_by_station_id[sid] = cams
    return station_by_id, cams_by_station_id

def bring_app_to_front(app_path, window_title_contains="eines.system3d"):
    from pywinauto import Desktop
    from pywinauto.application import Application
    try:
        dlg = Desktop(backend="uia").window(title_re=".*"+re.escape(window_title_contains)+".*", found_index=0)
        if dlg.exists(): dlg.set_focus(); return dlg
    except Exception: pass
    app = Application(backend="uia").start(cmd_line=f'"{app_path}"')
    time.sleep(2.5)
    dlg = Desktop(backend="uia").window(title_re=".*"+re.escape(window_title_contains)+".*")
    dlg.wait("ready", timeout=15); dlg.set_focus(); return dlg

def find_controls(dlg):
    controls = {}
    try:
        controls["station"] = dlg.child_window(title_re="Station", control_type="ComboBox")
        controls["camera"]  = dlg.child_window(title_re="Camera", control_type="ComboBox")
        controls["frame"]   = dlg.child_window(title_re="Frame", control_type="Edit")
        controls["posx"]    = dlg.child_window(title_re="Position X", control_type="Edit")
        controls["posy"]    = dlg.child_window(title_re="Position Y", control_type="Edit")
        controls["show"]    = dlg.child_window(title_re="Show", control_type="Button")
    except Exception: pass
    return controls

def select_station_and_camera(controls, station_name, camera_id):
    import pyautogui, time
    try: controls["station"].select(station_name)
    except Exception:
        rect = controls["station"].rectangle(); pyautogui.click(rect.mid_point().x, rect.mid_point().y); time.sleep(0.2)
        pyautogui.typewrite(station_name[:3]); pyautogui.press("enter")
    time.sleep(0.2)
    try: controls["camera"].select(str(camera_id))
    except Exception:
        rect = controls["camera"].rectangle(); pyautogui.click(rect.mid_point().x, rect.mid_point().y); time.sleep(0.2)
        pyautogui.typewrite(str(camera_id)); pyautogui.press("enter")

def fill_fields_and_show(controls, frame, x_frame, y_frame):
    import pyautogui, time
    def norm(v): s=str(v).strip(); return s.replace('.', ',') if DECIMAL_COMMA else s
    for key,val in [("frame",frame),("posx",x_frame),("posy",y_frame)]:
        try: controls[key].set_edit_text(norm(val))
        except Exception:
            rect = controls[key].rectangle(); pyautogui.click(rect.mid_point().x, rect.mid_point().y)
            pyautogui.hotkey("ctrl","a"); pyautogui.typewrite(norm(val))
    time.sleep(0.1)
    try: controls["show"].click()
    except Exception:
        rect = controls["show"].rectangle(); pyautogui.click(rect.mid_point().x, rect.mid_point().y)

def ocr_intersect_line(dlg):
    import pyautogui, numpy as np, cv2, pytesseract
    try:
        txt = dlg.child_window(title_re="Intersect position.*", control_type="Text")
        rect = txt.rectangle()
        img = pyautogui.screenshot(region=(rect.left, rect.top, rect.width(), rect.height()+5))
        frame = cv2.cvtColor(np.array(img), cv2.COLOR_RGBA2BGR)
    except Exception:
        r = dlg.rectangle()
        img = pyautogui.screenshot(region=(r.left, r.top, r.width(), r.height()))
        frame = cv2.cvtColor(np.array(img), cv2.COLOR_RGBA2BGR)
    return pytesseract.image_to_string(frame, lang="eng", config="--psm 6")

def parse_intersect(text):
    m = re.search(r"Intersect position:\s*([-+]?\d+,\d+)\s+([-+]?\d+,\d+)\s+([-+]?\d+,\d+)", text)
    x=y=z=t=None
    if m: x,y,z = m.group(1), m.group(2), m.group(3)
    m2 = re.search(r"([\d,]+)\s*ms", text)
    if m2: t=m2.group(1)
    return x,y,z,t,text

def main():
    import argparse
    p=argparse.ArgumentParser(); p.add_argument("--config", default="config.json"); p.add_argument("--dry", action="store_true"); p.add_argument("--run", action="store_true")
    args=p.parse_args()
    cfg=json.load(open(args.config,"r",encoding="utf-8"))
    excel_path=cfg["excel_path"]; xml_path=cfg["stations_xml_path"]; app_path=cfg["viewer_exe_path"]; title_hint=cfg.get("window_title_contains","eines.system3d")
    df,sheet,header=load_excel(excel_path); df=ensure_output_columns(df)
    station_by_id,cams_by_station_id=load_stations_map(xml_path)
    if args.dry:
        print(f"Excel '{excel_path}' sheet '{sheet}' rows={len(df)}"); print("Stations:",station_by_id); return
    dlg=bring_app_to_front(app_path, window_title_contains=title_hint); ctrls=find_controls(dlg)
    nr_col=None
    for c in df.columns:
        if c.startswith("NR"): nr_col=c; break
    if nr_col is None: nr_col=df.columns[0]
    need=["FRAME","X_FRAME","Y_FRAME","STATION","CAMERA"]; miss=[c for c in need if c not in df.columns]
    if miss: raise RuntimeError(f"Missing columns: {miss}")
    for idx,row in df.iterrows():
        try:
            nr=row.get(nr_col, idx+1); frame=row["FRAME"]; x=row["X_FRAME"]; y=row["Y_FRAME"]
            st_id=int(to_float(row["STATION"]) or row["STATION"]); cam_id=int(to_float(row["CAMERA"]) or row["CAMERA"])
            st_name=station_by_id.get(st_id,None)
            if not st_name: print(f"[WARN] Row {idx} Nr{nr}: unknown station {st_id}"); continue
            select_station_and_camera(ctrls, st_name, cam_id)
            fill_fields_and_show(ctrls, frame, x, y)
            time.sleep(0.3 + float(cfg.get("post_show_delay_s",0.7)))
            text=ocr_intersect_line(dlg); X,Y,Z,T,_=parse_intersect(text)
            df.at[idx,"X_NEW"]=X or ""; df.at[idx,"Y_NEW"]=Y or ""; df.at[idx,"Z_NEW"]=Z or ""; df.at[idx,"TIME_MS"]=T or ""
            print(f"[OK] Nr {nr} -> ({X},{Y},{Z}) {T} ms")
            time.sleep(float(cfg.get("per_row_pause_s",0.2)))
        except Exception as e:
            print(f"[ERR] Row {idx}: {e}")
    out_path=cfg.get("output_excel_path") or excel_path
    with pd.ExcelWriter(out_path, engine="openpyxl", mode="w") as w: df.to_excel(w, sheet_name=sheet, index=False)
    print(f"[DONE] Saved to {out_path}")

if __name__=="__main__": main()
