# EINES mini app (Windows)
Automatiza `eines.system3d.dentcheck.viewer.exe` leyendo tu Excel y capturando `Intersect position` por OCR.

## Requisitos
- Windows 10/11 + Python 3.10+ (64-bit)
- Tesseract OCR instalado y en PATH
- `pip install -r requirements.txt`

## Configuración
Edita `config.json`:
- `excel_path`: ruta a `POSITION.xlsx`
- `stations_xml_path`: ruta a `stations.xml`
- `viewer_exe_path`: ruta a `eines.system3d.dentcheck.viewer.exe`
- `window_title_contains`: texto que aparece en el título de la ventana
- `post_show_delay_s`, `per_row_pause_s`
- `output_excel_path`: opcional (si vacío, sobrescribe)

## Uso
- Prueba sin automatizar: `python eines_automation.py --dry`
- Ejecución completa: `python eines_automation.py --run`

Ajusta `find_controls()` y `ocr_intersect_line()` si tu UI tiene nombres distintos o el texto aparece en otra zona.
