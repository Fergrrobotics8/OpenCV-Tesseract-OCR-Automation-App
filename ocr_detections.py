# -*- coding: utf-8 -*-

# ========= CONFIGURACIÓN RÁPIDA =========
INPUT_IMAGE = r"C:\Users\fernando.garcia\Documents\COSAS FERNANDO\DEVELOPER\Eines automation porsche\debug_footer_img_for_ocr.png"
LANG        = "eng+spa"
MIN_CONF    = 60          # confianza mínima por palabra (pasada 1)
PSM         = "11"        # 6 o 11 van bien para UI
SCALE       = 2.0         # reescalado previo
LABEL_SIZE  = 0.4         # tamaño de etiqueta
ALPHA       = 0.25        # opacidad [0..1]
SHOW_CONF   = False       # mostrar confianza junto al texto

# ---- Coordenadas fijas (en lienzo ya ESCALADO) ----
FIXED_SHOW_POINT  = (1032, 108)  # x,y definitivos del botón "Show"
FIXED_SHOW_SCALE  = 2.0          # escala a la que se midió el punto fijo

# Offsets para ajustar la cruz (se aplican tras calcular el punto fijo/reescalado)
OFFSET_X    = 0          # píxeles hacia la DERECHA (negativo = izquierda)
OFFSET_Y    = 0          # píxeles en VERTICAL (positivo = abajo, negativo = arriba)
# =======================================

import os, sys, re, cv2, numpy as np, pytesseract, shutil, platform
from typing import List, Dict, Tuple

def ensure_tesseract_available():
    if shutil.which("tesseract") is None and platform.system() == "Windows":
        exe = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        if os.path.exists(exe):
            pytesseract.pytesseract.tesseract_cmd = exe
    pytesseract.get_tesseract_version()

def preprocess_for_gui(img_bgr, scale=SCALE):
    if scale != 1.0:
        img_bgr = cv2.resize(img_bgr, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
    gray = cv2.convertScaleAbs(gray, alpha=1.8, beta=0)
    thr  = cv2.adaptiveThreshold(gray,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY,35,9)
    return thr, img_bgr

def draw_label(canvas, x, y, w, h, label, font_scale=LABEL_SIZE, alpha=ALPHA,
               box_color=(0,255,0), text_color=(255,255,255)):
    cv2.rectangle(canvas, (x,y), (x+w,y+h), box_color, 2)
    H,W = canvas.shape[:2]
    font = cv2.FONT_HERSHEY_SIMPLEX
    (tw,th), _ = cv2.getTextSize(label, font, font_scale, 1)
    pad = 3
    options = [
        (x, max(0, y - th - 2*pad), x + tw + 2*pad, max(0, y)),
        (x, min(H, y + h), x + tw + 2*pad, min(H, y + h + th + 2*pad)),
        (min(W, x + w), max(0, y), min(W, x + w + tw + 2*pad), max(0, y + th + 2*pad)),
    ]
    for (a,b,c,d) in options:
        if 0 <= a < W and 0 <= c <= W and 0 <= b < H and 0 <= d <= H:
            rx1,ry1,rx2,ry2 = a,b,c,d
            break
    else:
        rx1,ry1,rx2,ry2 = max(0,x), max(0,y-th-2*pad), min(W,x+tw+2*pad), max(0,y)
    overlay = canvas.copy()
    cv2.rectangle(overlay, (rx1,ry1), (rx2,ry2), box_color, -1)
    cv2.addWeighted(overlay, alpha, canvas, 1-alpha, 0, dst=canvas)
    cv2.putText(canvas, label, (rx1+pad, ry2-pad), font, font_scale, text_color, 1, cv2.LINE_AA)

def draw_cross(canvas, cx, cy, size=14, color=(0,0,255), thickness=2):
    H, W = canvas.shape[:2]
    x1, x2 = max(0, cx - size), min(W-1, cx + size)
    y1, y2 = max(0, cy - size), min(H-1, cy + size)
    cv2.line(canvas, (x1, cy), (x2, cy), color, thickness, cv2.LINE_AA)
    cv2.line(canvas, (cx, y1), (cx, y2), color, thickness, cv2.LINE_AA)

def first_pass(ocr_img, lang=LANG, psm=PSM, min_conf=MIN_CONF) -> List[Dict]:
    cfg = f"--oem 3 --psm {psm}"
    d = pytesseract.image_to_data(ocr_img, lang=lang, config=cfg, output_type=pytesseract.Output.DICT)
    out = []
    for i in range(len(d["text"])):
        txt = (d["text"][i] or "").strip()
        try: conf = float(d["conf"][i])
        except: conf = -1.0
        if txt and conf >= min_conf:
            out.append({"x":int(d["left"][i]), "y":int(d["top"][i]),
                        "w":int(d["width"][i]), "h":int(d["height"][i]),
                        "text":txt, "conf":conf})
    out.sort(key=lambda r: (r["y"], r["x"]))
    return out

def _vertical_overlap(a:Dict, b:Dict) -> float:
    ay1, ay2 = a["y"], a["y"] + a["h"]
    by1, by2 = b["y"], b["y"] + b["h"]
    inter = max(0, min(ay2, by2) - max(ay1, by1))
    return inter / (min(a["h"], b["h"]) + 1e-6)

def _bbox_union(a:Dict, b:Dict) -> Tuple[int,int,int,int]:
    x1 = min(a["x"], b["x"])
    y1 = min(a["y"], b["y"])
    x2 = max(a["x"]+a["w"], b["x"]+b["w"])
    y2 = max(a["y"]+a["h"], b["y"]+b["h"])
    return x1, y1, x2-x1, y2-y1

def merge_adjacent_words(dets: List[Dict],
                         max_gap_px:int=16,
                         min_overlap:float=0.5) -> List[Dict]:
    """
    Agrupa tokens seguidos en la misma línea si el hueco entre cajas es pequeño.
    Incluye también 'Position' con su letra/símbolo (X/Y/Z/?/??/�).
    """
    if not dets: return dets[:]
    merged = []
    line = [dets[0]]
    for cur in dets[1:]:
        prev = line[-1]
        gap = cur["x"] - (prev["x"] + prev["w"])
        same_line = _vertical_overlap(prev, cur) >= min_overlap
        if same_line and 0 <= gap <= max_gap_px:
            new_text = (prev["text"] + " " + cur["text"]).strip()
            x,y,w,h = _bbox_union(prev, cur)
            prev.update({"text": new_text, "x":x, "y":y, "w":w, "h":h, "conf": max(prev["conf"], cur["conf"])})
            line[-1] = prev
        else:
            merged.extend(line)
            line = [cur]
    merged.extend(line)
    merged.sort(key=lambda r: (r["y"], r["x"]))
    return merged

def fix_second_position_Y_in_merged(dets_merged: List[Dict]) -> Tuple[List[Dict], bool]:
    """
    En la lista YA AGRUPADA:
    Localiza la 2ª 'Position' (o la única) y normaliza a 'Position Y'
    si el sufijo tras 'Position' es '?', '??', '�', 1–2 no alfanuméricos o letra ≠ 'Y'.
    """
    pos_idx = [i for i, d in enumerate(dets_merged)
               if d["text"].strip().lower().startswith("position")]
    if not pos_idx:
        print("[INFO] No hay 'Position' en MERGED."); return dets_merged, False
    pos_idx.sort(key=lambda i: dets_merged[i]["y"])
    idx = pos_idx[1] if len(pos_idx) >= 2 else pos_idx[0]
    t = dets_merged[idx]["text"].strip()
    suffix = re.sub(r"(?i)^position", "", t).strip()
    print(f"[DBG] 2ª 'Position' candidato: '{t}'  |  sufijo='{suffix}'")

    def set_Y(i):
        before = dets_merged[i]["text"]
        dets_merged[i]["text"] = "Position Y"
        dets_merged[i]["conf"] = max(dets_merged[i].get("conf", 0), 99.0)
        print(f"[FIX] '{before}'  ->  'Position Y'")
        return True

    if suffix == "": return dets_merged, False
    if suffix in {"?", "??", "�"}: return dets_merged, set_Y(idx)
    if len(suffix) <= 2 and not re.search(r"[A-Za-z0-9]", suffix): return dets_merged, set_Y(idx)
    if "?" in suffix: return dets_merged, set_Y(idx)

    m = re.fullmatch(r"[A-Za-z]", suffix)
    if m:
        return (dets_merged, False) if m.group(0).upper()=="Y" else (dets_merged, set_Y(idx))

    m2 = re.search(r"([A-Za-z])", suffix)
    if m2:
        if m2.group(1).upper()=="Y":
            dets_merged[idx]["text"] = "Position Y"
            dets_merged[idx]["conf"] = max(dets_merged[idx].get("conf", 0), 99.0)
            print("[OK] Ya contiene 'Y' (normalizo)."); return dets_merged, True
        else:
            return dets_merged, set_Y(idx)

    if len(suffix) <= 3 and not re.search(r"[A-Za-z0-9]", suffix): return dets_merged, set_Y(idx)
    print("[INFO] Sufijo tras 'Position' no requiere cambio."); return dets_merged, False

def print_dets(title: str, dets: List[Dict]):
    print(f"\n=== {title} ===")
    print(f"Total: {len(dets)} | MIN_CONF: {MIN_CONF} | PSM: {PSM} | SCALE: {SCALE}\n")
    print(f"{'#':>3}  {'texto':<24}  {'conf':>5}  {'x':>4} {'y':>4} {'w':>4} {'h':>4}")
    print("-"*70)
    for i, r in enumerate(dets, 1):
        txt = r['text'] if len(r['text']) <= 24 else r['text'][:21] + "…"
        print(f"{i:>3}  {txt:<24}  {int(r['conf']):>5}  {r['x']:>4} {r['y']:>4} {r['w']:>4} {r['h']:>4}")
    print("-"*70)

def main():
    ensure_tesseract_available()

    in_path = INPUT_IMAGE
    if not os.path.isfile(in_path):
        print(f"ERROR: no existe el archivo: {in_path}", file=sys.stderr); return

    img_bgr = cv2.imread(in_path, cv2.IMREAD_COLOR)
    if img_bgr is None:
        print("ERROR: no se pudo abrir la imagen.", file=sys.stderr); return

    # Preproceso + lienzo escalado
    ocr_img, canvas_base = preprocess_for_gui(img_bgr, scale=SCALE)

    # PASADA 1 (RAW)
    dets_raw = first_pass(ocr_img, LANG, PSM, MIN_CONF)
    print_dets("DETECCIONES RAW (pasada 1)", dets_raw)

    # Guardar overlay RAW
    root, ext = os.path.splitext(in_path)
    canvas_raw = canvas_base.copy()
    for r in dets_raw:
        lbl = f"{r['text']} ({int(r['conf'])})" if SHOW_CONF else r['text']
        draw_label(canvas_raw, r["x"], r["y"], r["w"], r["h"], lbl)
    out_raw = f"{root} Detections{ext if ext else '.png'}"
    cv2.imwrite(out_raw, canvas_raw)
    print(f"[OK] Imagen RAW: {out_raw}")

    # MERGE (incluye 'Position X' y 'Position Y')
    dets_merged = merge_adjacent_words(dets_raw, max_gap_px=16, min_overlap=0.5)
    print_dets("DETECCIONES MERGED (agrupadas)", dets_merged)

    # Corrección robusta de 'Position Y' en MERGED (2ª 'Position')
    dets_final, changed = fix_second_position_Y_in_merged(dets_merged)
    if changed:
        print("[OK] 'Position Y' corregido en MERGED.")
    else:
        print("[INFO] 'Position Y' ya estaba correcto o no aplica.")

    # Overlay FINAL (MERGED + corrección)
    print_dets("DETECCIONES FINALES (para overlay)", dets_final)
    canvas_final = canvas_base.copy()
    for r in dets_final:
        lbl = f"{r['text']} ({int(r['conf'])})" if SHOW_CONF else r['text']
        draw_label(canvas_final, r["x"], r["y"], r["w"], r["h"], lbl)
    out_final = f"{root} Detections 2{ext if ext else '.png'}"
    cv2.imwrite(out_final, canvas_final)
    print(f"[OK] Imagen FINAL: {out_final}")

    # ===== PASADA 3: usar punto FIJO de 'Show' y pintar cruz =====
    # Reescala el punto fijo si cambia SCALE
    if FIXED_SHOW_SCALE and FIXED_SHOW_SCALE > 0 and FIXED_SHOW_SCALE != SCALE:
        factor = SCALE / FIXED_SHOW_SCALE
        midx = int(round(FIXED_SHOW_POINT[0] * factor))
        midy = int(round(FIXED_SHOW_POINT[1] * factor))
    else:
        midx, midy = FIXED_SHOW_POINT

    # aplicar offsets
    midx += int(OFFSET_X)
    midy += int(OFFSET_Y)

    canvas_show = canvas_final.copy()
    draw_cross(canvas_show, midx, midy, size=14, color=(0,0,255), thickness=2)
    cv2.putText(canvas_show, f"({midx},{midy})", (midx+8, midy-8),
                cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0,0,255), 1, cv2.LINE_AA)

    out_show = f"{root} Detections 3{ext if ext else '.png'}"
    cv2.imwrite(out_show, canvas_show)

    print("\n[SHOW] Punto FIJO del botón 'Show':")
    print(f"       Coordenadas (x,y) = ({midx}, {midy})  [SCALE={SCALE}, OFF=({OFFSET_X},{OFFSET_Y})]")
    print(f"       Punto base fijo = {FIXED_SHOW_POINT} @ SCALE={FIXED_SHOW_SCALE}")
    print(f"[OK] Imagen con cruz: {out_show}")

if __name__ == "__main__":
    main()
