#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GEO-SPHERE · Mapa de ubicaciones de Timbío, Cauca
-------------------------------------------------
Reemplaza al index.html. Genera una página HTML autónoma y la abre en el navegador.

Flujo (según la imagen de referencia):
  1) Globo terráqueo -> clic (o buscar nombre / "lat, lon") para elegir ubicación inicial
  2) "NAVEGAR A LA UBICACIÓN" -> el globo gira 2 rotaciones mostrando LAT/LON en vivo
  3) "UBICACIÓN ADQUIRIDA" -> aterriza en mapa satelital con todos los lugares
  4) "CONFIRMAR UBICACIÓN" -> encuadra todos los lugares visibles

Uso:
  python geo_sphere_timbio.py                      # usa los 38 lugares incluidos aquí
  python geo_sphere_timbio.py --xlsx datos.xlsx    # lee los lugares desde un Excel
  python geo_sphere_timbio.py --no-open            # solo genera el HTML

Requiere internet en el navegador (globo, mapas Esri y Leaflet se cargan por CDN).
Para --xlsx: pip install pandas openpyxl
"""
import argparse
import json
import sys
import unicodedata
import webbrowser
from pathlib import Path

# ----------------------------------------------------------------------------
# DATOS  (n=nombre, c=categoría, d=dirección, o=observaciones, i=ícono, col=color,
#         cn=direcciones cercanas N/S/E/O)
# ----------------------------------------------------------------------------
DATA = [
 {"id": 1, "n": "Estación de gasolina GLIX", "c": "Estación de servicio", "d": "Cl. 18 # 20-41", "lat": 2.35364666666667, "lng": -76.684995, "o": "Misma dirección que Transtimbio Transportes", "i": "⛽", "col": "#F59E0B", "cn": {"N": "", "S": "", "E": "", "O": ""}},
 {"id": 2, "n": "Polvorería El vaquero", "c": "Comercio", "d": "Popayan - Rosas", "lat": 2.360035, "lng": -76.6824583333333, "o": "Dirección aproximada (vía Popayán - Rosas); la app no dio nomenclatura urbana", "i": "🧨", "col": "#0078D4", "cn": {"N": "", "S": "", "E": "", "O": ""}},
 {"id": 3, "n": "Hospital Timbío", "c": "Salud", "d": "Cl. 18 #2122", "lat": 2.35433833333333, "lng": -76.6858283333333, "o": "", "i": "🏥", "col": "#E11D48", "cn": {"N": "", "S": "", "E": "", "O": ""}},
 {"id": 4, "n": "Escuela San Cayetano", "c": "Educación", "d": "Cl. 17 # 8-32", "lat": 2.35104833333333, "lng": -76.674785, "o": "Nombre escrito 'Escuela san ayetano' en la captura", "i": "🏫", "col": "#7C3AED", "cn": {"N": "", "S": "", "E": "", "O": ""}},
 {"id": 5, "n": "Santo San Cayetano", "c": "Religion", "d": "Cl. 17 #938", "lat": 2.35098666666667, "lng": -76.6759633333333, "o": "Nombre escrito 'Santo sancayetano' en la captura; tipo de lugar no indicado", "i": "⛪", "col": "#A16207", "cn": {"N": "", "S": "", "E": "", "O": ""}},
 {"id": 6, "n": "Estadio municipal", "c": "Deporte y recreación", "d": "Cl. 17 #1047", "lat": 2.35111833333333, "lng": -76.6770666666667, "o": "", "i": "🏟️", "col": "#16A34A", "cn": {"N": "", "S": "", "E": "", "O": ""}},
 {"id": 7, "n": "Droguería Droxi", "c": "Salud", "d": "Cl. 17 # 11-90", "lat": 2.351665, "lng": -76.6785066666667, "o": "", "i": "💊", "col": "#E11D48", "cn": {"N": "", "S": "", "E": "", "O": ""}},
 {"id": 8, "n": "Panadería La Esquina", "c": "Comercio", "d": "Cl. 17 #12-81", "lat": 2.35174666666667, "lng": -76.6789483333333, "o": "", "i": "🥖", "col": "#0078D4", "cn": {"N": "", "S": "", "E": "", "O": ""}},
 {"id": 9, "n": "Clínica dental", "c": "Salud", "d": "Cl. 17 # 13-12", "lat": 2.351835, "lng": -76.67926, "o": "Misma dirección y coordenadas casi idénticas a Acertemos", "i": "🦷", "col": "#E11D48", "cn": {"N": "", "S": "", "E": "", "O": ""}},
 {"id": 10, "n": "Acertemos", "c": "Comercio", "d": "Cl. 17 # 13-12", "lat": 2.351855, "lng": -76.67929, "o": "Misma dirección y coordenadas casi idénticas a Clínica dental", "i": "💼", "col": "#0078D4", "cn": {"N": "", "S": "", "E": "", "O": ""}},
 {"id": 11, "n": "Vidrios Martínez", "c": "Comercio", "d": "Cra. 14 # 17-3", "lat": 2.35215333333333, "lng": -76.68039666666671, "o": "Misma dirección que Parqueadero La Virgen y Escuela La Sucre", "i": "🪟", "col": "#0078D4", "cn": {"N": "", "S": "", "E": "", "O": ""}},
 {"id": 12, "n": "Parqueadero La Virgen", "c": "Servicios", "d": "Cra. 14 # 17-3", "lat": 2.3522172884449346, "lng": -76.6805, "o": "Misma dirección que Vidrios Martínez y Escuela La Sucre", "i": "🅿️", "col": "#0891B2", "cn": {"N": "", "S": "", "E": "", "O": ""}},
 {"id": 13, "n": "Escuela La Sucre", "c": "Educación", "d": "Cra. 14 # 17-3", "lat": 2.352096044888405, "lng": -76.6805, "o": "Coordenadas idénticas a Parqueadero La Virgen (posible captura repetida)", "i": "🏫", "col": "#7C3AED", "cn": {"N": "", "S": "", "E": "", "O": ""}},
 {"id": 14, "n": "Barberías", "c": "Servicios", "d": "Cl. 17 # 14-77", "lat": 2.3522, "lng": -76.6811616666667, "o": "", "i": "💈", "col": "#0891B2", "cn": {"N": "", "S": "", "E": "", "O": ""}},
 {"id": 15, "n": "Parque Simón Bolívar", "c": "Espacio público", "d": "Cl. 17 #1703", "lat": 2.352565, "lng": -76.6828, "o": "Misma dirección que Clínica Digital Dentalab", "i": "🌳", "col": "#65A30D", "cn": {"N": "", "S": "", "E": "", "O": ""}},
 {"id": 16, "n": "Clínica Digital Dentalab", "c": "Salud", "d": "Cl. 17 #1703", "lat": 2.35256666666667, "lng": -76.6829516666667, "o": "Misma dirección que Parque Simón Bolívar; hora no visible en la captura", "i": "🦷", "col": "#E11D48", "cn": {"N": "", "S": "", "E": "", "O": ""}},
 {"id": 17, "n": "Betplay", "c": "Comercio", "d": "Cl. 17 # 19-17", "lat": 2.35281333333333, "lng": -76.6841233333333, "o": "Misma dirección que Moto Montaña del Sur", "i": "🎰", "col": "#0078D4", "cn": {"N": "", "S": "", "E": "", "O": ""}},
 {"id": 18, "n": "Moto Montaña del Sur", "c": "Comercio", "d": "Cl. 17 # 19-17", "lat": 2.35280833333333, "lng": -76.68427, "o": "Misma dirección que Betplay", "i": "🏍️", "col": "#0078D4", "cn": {"N": "", "S": "", "E": "", "O": ""}},
 {"id": 19, "n": "Transtimbio Transportes", "c": "Transporte", "d": "Cl. 18 # 20-41", "lat": 2.35387666666667, "lng": -76.6849683333333, "o": "Misma dirección que Estación de gasolina GLIX", "i": "🚌", "col": "#475569", "cn": {"N": "", "S": "", "E": "", "O": ""}},
 {"id": 20, "n": "Estación de gasolina Primax", "c": "Estación de servicio", "d": "Cra. 21 # 19-10", "lat": 2.35437166666667, "lng": -76.6852866666667, "o": "", "i": "⛽", "col": "#F59E0B", "cn": {"N": "", "S": "", "E": "", "O": ""}},
 {"id": 21, "n": "Cancha San Judas", "c": "Deporte y recreación", "d": "Cl. 18 # 23-49", "lat": 2.35488, "lng": -76.68706333333333, "o": "", "i": "⚽", "col": "#16A34A", "cn": {"N": "Cl. 18 # 23-71", "S": "Cl. 18 #2324", "E": "Urgencias, Cl. 18 #23-10", "O": "Cl. 18 # 23-71"}},
 {"id": 22, "n": "Iglesia San Judas", "c": "Religion", "d": "Calle 19 # 23-33", "lat": 2.355273333333333, "lng": -76.686775, "o": "", "i": "⛪", "col": "#A16207", "cn": {"N": "Calle 19 # 23-50", "S": "Cra. 23 # 18-7", "E": "Calle 19 # 23-33", "O": "Calle 19 # 23-33"}},
 {"id": 23, "n": "Droguería Medlife plus", "c": "Salud", "d": "Cl. 19 # 22-12", "lat": 2.3548366666666665, "lng": -76.68591666666666, "o": "", "i": "💊", "col": "#E11D48", "cn": {"N": "Cl. 19 # 22-24", "S": "Cl. 19 #2213", "E": "Cl. 19 # 22-8", "O": "Cl. 19 # 22-31"}},
 {"id": 24, "n": "Panadería Maxipan", "c": "Comercio", "d": "Cl. 18 # 21-11", "lat": 2.35393, "lng": -76.68540833333333, "o": "", "i": "🥖", "col": "#0078D4", "cn": {"N": "Cl. 18 #2084", "S": "Cra. 21 # 17-57", "E": "Cl. 18 #2056", "O": "Cl. 18 # 21-19"}},
 {"id": 25, "n": "Clínica Odontoestetica", "c": "Salud", "d": "Cl. 17 # 21-8", "lat": 2.35327, "lng": -76.68543666666666, "o": "", "i": "🦷", "col": "#E11D48", "cn": {"N": "Cra. 21 # 17-32", "S": "Cl. 17 # 20-64", "E": "Cl. 17 # 20-42", "O": "Cra. 21 # 17-3"}},
 {"id": 26, "n": "Ferretería El baratos de el sur", "c": "Comercio", "d": "Cra. 21 # 17-22", "lat": 2.3533416666666667, "lng": -76.68541, "o": "Nombre escrito 'Ferretería El baratos de el sur' en la captura", "i": "🔧", "col": "#0078D4", "cn": {"N": "Cra. 21 # 17-38", "S": "Cl. 17 # 20-66", "E": "Cra. 21 # 17-24", "O": "Cra. 21 # 17-17"}},
 {"id": 27, "n": "Asmet salud", "c": "Salud", "d": "Cra. 20 #1551", "lat": 2.3522883333333335, "lng": -76.68498666666666, "o": "", "i": "🏥", "col": "#E11D48", "cn": {"N": "Cra. 20 # 16-11", "S": "Cra. 20 # 15-42", "E": "Cl. 16 # 19-74", "O": "Cl. 16 #20-07"}},
 {"id": 28, "n": "Ferretería Don Carlos", "c": "Comercio", "d": "Cl. 16 # 19-60", "lat": 2.3522883333333335, "lng": -76.684735, "o": "", "i": "🔧", "col": "#0078D4", "cn": {"N": "Cl. 16 # 20-36", "S": "Cl. 16 # 19-57", "E": "Cl. 16 # 19-70", "O": "Cra. 20 #1551"}},
 {"id": 29, "n": "Paola Herrera perfumería", "c": "Comercio", "d": "Cra. 19 #calle 16", "lat": 2.35203, "lng": -76.68419666666667, "o": "La app indicó 'Centro' como barrio", "i": "🧴", "col": "#0078D4", "cn": {"N": "Cra. 19 # 16-9", "S": "Cra. 19 # 15-36", "E": "Cl. 16 # 18-54", "O": "Cl. 16 # 19-9"}},
 {"id": 30, "n": "Colegio los andes", "c": "Educación", "d": "Cl. 15 #18-19", "lat": 2.3513183333333334, "lng": -76.68404666666667, "o": "", "i": "🏫", "col": "#7C3AED", "cn": {"N": "Cl. 15 # 18-36", "S": "Cra. 18 # 14A-63", "E": "Cl. 15 # 18-6", "O": "Cl. 15 # 18-49"}},
 {"id": 31, "n": "Alcaldía Municipal", "c": "Gobierno", "d": "Cl. 15 # 17-11", "lat": 2.3511333333333333, "lng": -76.68336666666667, "o": "Categoría 'Gobierno' asignada por el nombre", "i": "🏛️", "col": "#334155", "cn": {"N": "Cl. 15 # 17-22", "S": "Cra. 17 # 14A-49", "E": "Parque, Cra. 17 #18-24", "O": "Cl. 15 #17-33"}},
 {"id": 32, "n": "Parque Boyaca", "c": "Espacio público", "d": "Cl. 15 # 15-60", "lat": 2.350925, "lng": -76.68267333333333, "o": "", "i": "🌳", "col": "#65A30D", "cn": {"N": "Cl. 15 # 16-24", "S": "Cra. 16 # 14-64", "E": "Cl. 15 # 15-60", "O": "Cra. 16 # 14-64"}},
 {"id": 33, "n": "Iglesia San Pedro", "c": "Religion", "d": "Cra. 16 # 15-4", "lat": 2.3510050000000002, "lng": -76.682525, "o": "", "i": "⛪", "col": "#A16207", "cn": {"N": "Cra. 16 #15-26", "S": "Cl. 15 # 15-47", "E": "Cl. 15 # 15-34", "O": "Cl. 15 # 16-24"}},
 {"id": 34, "n": "Colegio San Antonio", "c": "Educación", "d": "Cra. 15 #15-04", "lat": 2.3507883333333335, "lng": -76.682015, "o": "", "i": "🏫", "col": "#7C3AED", "cn": {"N": "Cra. 15 # 03C-15", "S": "Cl. 15 # 15-3", "E": "Cl. 15 # 14-114", "O": "Cl. 15 # 15-3"}},
 {"id": 35, "n": "Supermercado D1", "c": "Comercio", "d": "Cl. 18 #15-18", "lat": 2.353135, "lng": -76.68136833333332, "o": "", "i": "🛒", "col": "#0078D4", "cn": {"N": "Cl. 18 # 15 - 22", "S": "Cra. 15 # 17-89", "E": "Cl. 18 # 15-18", "O": "Cl. 18 # 15-34"}},
 {"id": 36, "n": "Minimarket 13", "c": "Comercio", "d": "Cl. 18 # 14-41", "lat": 2.3527683333333336, "lng": -76.68064, "o": "", "i": "🛒", "col": "#0078D4", "cn": {"N": "Cl. 18 # 14-40", "S": "Cl. 18 # 14-19", "E": "Cl. 18 # 14-26", "O": "Cl. 18 # 14-49"}},
 {"id": 37, "n": "Servimetalicas", "c": "Comercio", "d": "Cra. 14 #17-47", "lat": 2.35246, "lng": -76.68025666666668, "o": "", "i": "🔩", "col": "#0078D4", "cn": {"N": "Cl. 18 # 14-2", "S": "Cra. 14 # 17-20", "E": "Cl. 18 # 13-66", "O": "Cra. 14 #17-45"}},
 {"id": 38, "n": "Minimarket Merka Cauca", "c": "Comercio", "d": "Cl. 17 #1283", "lat": 2.35192, "lng": -76.67962000000001, "o": "Captura repetida dos veces (se guardó una sola)", "i": "🛒", "col": "#0078D4", "cn": {"N": "Cl. 18 # 13-6", "S": "Cra. 13 #16-76", "E": "Cl. 17 # 13-2", "O": "Cl. 17 # 13-10"}},
]

PALETTE = ["#F59E0B", "#0078D4", "#E11D48", "#7C3AED", "#A16207", "#16A34A",
           "#0891B2", "#65A30D", "#475569", "#334155", "#DB2777", "#EA580C"]
ICON_KW = [("drogu", "💊"), ("farmacia", "💊"), ("hospital", "🏥"), ("salud", "🏥"),
           ("clinica", "🦷"), ("dental", "🦷"), ("odonto", "🦷"), ("escuela", "🏫"),
           ("colegio", "🏫"), ("iglesia", "⛪"), ("parque", "🌳"), ("estadio", "🏟️"),
           ("cancha", "⚽"), ("panader", "🥖"), ("ferreter", "🔧"), ("super", "🛒"),
           ("minimarket", "🛒"), ("gasolina", "⛽"), ("alcald", "🏛️"), ("taller", "🛠️"),
           ("transporte", "🚌"), ("perfumer", "🧴")]


def _norm(s):
    s = unicodedata.normalize("NFD", str(s)).encode("ascii", "ignore").decode()
    return s.strip().lower()


def _num(v):
    return float(str(v).strip().replace(",", "."))


def load_xlsx(path):
    import pandas as pd
    xl = pd.ExcelFile(path)
    sheet = "Todas" if "Todas" in xl.sheet_names else xl.sheet_names[0]
    df = xl.parse(sheet).fillna("")
    cols = {_norm(c): c for c in df.columns}

    def pick(*names):
        for n in names:
            if n in cols:
                return cols[n]
        return None

    cn_, cc_, cd_ = pick("nombre del lugar", "nombre", "lugar"), pick("categoria"), pick("direccion")
    clat, clng = pick("latitud", "lat"), pick("longitud", "lng", "lon")
    cobs, cid = pick("observaciones"), pick("id")
    near = {k: pick("direccion cercana " + k.lower()) for k in "NSEO"}
    if not (cn_ and clat and clng):
        sys.exit("El Excel necesita al menos: Nombre del lugar, Latitud y Longitud.")
    cat_color, out = {}, []
    for k, r in df.iterrows():
        try:
            lat, lng = _num(r[clat]), _num(r[clng])
        except ValueError:
            continue
        name = str(r[cn_]).strip()
        cat = str(r[cc_]).strip() if cc_ else "Otros"
        cat_color.setdefault(cat, PALETTE[len(cat_color) % len(PALETTE)])
        key = _norm(name + " " + cat)
        icon = next((i for kw, i in ICON_KW if kw in key), "📍")
        out.append({
            "id": int(r[cid]) if cid and str(r[cid]).strip() != "" else k + 1,
            "n": name, "c": cat or "Otros", "d": str(r[cd_]).strip() if cd_ else "",
            "lat": lat, "lng": lng, "o": str(r[cobs]).strip() if cobs else "",
            "i": icon, "col": cat_color[cat],
            "cn": {k2: (str(r[v]).strip() if v else "") for k2, v in near.items()},
        })
    return out


HTML = r'''<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>GEO-SPHERE · Timbío, Cauca</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.css">
<script src="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.js"></script>
<script src="https://cdn.jsdelivr.net/npm/globe.gl@2"></script>
<style>
:root{--bg:#070d13;--pan:#0d1822;--pan2:#12222e;--ac:#2dd4bf;--tx:#cfe8ee;--mut:#6b8794;--line:#1c3340}
*{box-sizing:border-box}
html,body{height:100%;margin:0}
body{font-family:Segoe UI,Montserrat,Arial,sans-serif;background:radial-gradient(ellipse at 40% 40%,#0c2233 0%,var(--bg) 70%);color:var(--tx);display:flex;overflow:hidden}
#main{flex:1;display:flex;flex-direction:column;min-width:0}
#top{padding:12px 14px;display:flex;gap:8px;border-bottom:1px solid var(--line);background:#0a131b}
#q{flex:1;background:#0b1620;border:1px solid var(--line);border-radius:8px;color:var(--tx);padding:10px 12px;font-size:12px;letter-spacing:.06em;outline:none}
#q:focus{border-color:var(--ac)}#q::placeholder{color:var(--mut);text-transform:uppercase}
#stage{position:relative;flex:1;min-height:0}
#globe,#map{position:absolute;inset:0}
#map{visibility:hidden}
body.s-map #map{visibility:visible}body.s-map #globe{visibility:hidden}
#side{width:330px;background:var(--pan);border-left:1px solid var(--line);display:flex;flex-direction:column;padding:16px;gap:12px;overflow:hidden}
.brand{font-size:12px;letter-spacing:.2em;color:var(--mut)}.brand b{display:block;font-size:22px;letter-spacing:.12em;color:#fff;margin-top:2px}
.tag{border-left:3px solid var(--ac);background:#0f2a30;padding:8px 10px;font-size:11px;letter-spacing:.12em;color:var(--ac);text-transform:uppercase}
.btn{cursor:pointer;background:#0f2a30;color:var(--tx);border:1px solid var(--ac);border-radius:6px;padding:9px 12px;font-size:11px;letter-spacing:.12em;text-transform:uppercase}
.btn:hover{background:var(--ac);color:#04201c}.btn.ghost{border-color:var(--line);color:var(--mut)}
#selc{font-size:13px;line-height:1.7}#selc b{color:#fff}#selc small{color:var(--mut)}
#p-sel,#p-map,#p-select{display:none;flex-direction:column;gap:10px;min-height:0}
body.s-select #p-select{display:flex}
body.s-select.has-sel #p-sel{display:flex}
body.s-nav #p-sel{display:flex}body.s-nav #p-sel .btn{display:none}
body.s-map #p-map{display:flex;flex:1}
#kpi{display:flex;gap:8px}#kpi div{flex:1;background:#0f2a30;border:1px solid var(--line);border-radius:8px;padding:6px;text-align:center;font-size:10px;color:var(--mut);letter-spacing:.08em;text-transform:uppercase}
#kpi b{display:block;font-size:22px;color:var(--ac)}
#cats{display:flex;flex-wrap:wrap;gap:6px}
.chip{cursor:pointer;padding:3px 9px;border-radius:14px;font-size:11px;border:2px solid;user-select:none;color:#fff}
.chip.off{opacity:.3}
#list{flex:1;overflow:auto;display:flex;flex-direction:column;gap:6px;min-height:0}
.it{background:var(--pan2);border-radius:8px;padding:7px 9px;cursor:pointer;display:flex;gap:9px;align-items:center;border-left:5px solid}
.it:hover{background:#18313f}.it b{font-size:13px}.it span{font-size:11px;color:var(--mut)}.it em{font-size:21px;font-style:normal}
#navov,#acq{position:absolute;inset:0;display:none;align-items:center;justify-content:center;flex-direction:column;pointer-events:none;z-index:900}
body.s-nav #navov{display:flex}
#acq.on{display:flex}
.navt{text-align:center;font-size:26px;letter-spacing:.14em;font-weight:600;text-shadow:0 0 18px #000}.navt small{font-size:15px;color:var(--ac)}
.coords{position:absolute;bottom:60px;text-align:center;font-size:12px;letter-spacing:.2em;color:var(--mut)}
.coords b{display:block;font-size:26px;color:var(--ac);letter-spacing:.08em}
#prog{position:absolute;bottom:26px;width:220px;height:3px;background:var(--line)}#prog i{display:block;height:100%;width:0;background:var(--ac)}
.ring{width:150px;height:150px;border-radius:50%;border:2px solid var(--ac);background:#06131bcc;box-shadow:0 0 40px #2dd4bf66;display:flex;flex-direction:column;align-items:center;justify-content:center;font-size:30px;text-align:center}
.ring span{font-size:11px;letter-spacing:.14em;margin-top:4px}
#bar{position:absolute;left:0;right:0;bottom:0;z-index:1000;display:none;gap:16px;align-items:center;padding:10px 16px;background:#070d13ee;border-top:1px solid var(--line);font-size:11px;letter-spacing:.1em;color:var(--mut)}
body.s-map #bar{display:flex}#bar b{color:#fff}#bar .sp{flex:1}
.pin{width:38px;height:38px;border-radius:50% 50% 50% 0;transform:rotate(-45deg);border:3px solid #fff;display:flex;align-items:center;justify-content:center;box-shadow:0 2px 6px #0007}
.pin i{transform:rotate(45deg);font-style:normal;font-size:19px}
.selpin{width:18px;height:18px;border-radius:50%;background:var(--ac);border:3px solid #fff;box-shadow:0 0 0 6px #2dd4bf55}
.leaflet-popup-content{font-size:12px;line-height:1.5}
.nb{display:grid;grid-template-columns:1fr 1fr;gap:2px 10px;margin-top:4px;font-size:11px;color:#334155}
</style></head>
<body class="s-select">
<div id="main">
  <div id="top"><input id="q" placeholder="Buscar una ubicación o punto de coordenadas (ej: 2.35, -76.68)"></div>
  <div id="stage">
    <div id="globe"></div><div id="map"></div>
    <div id="navov"><div class="navt">NAVEGANDO…<br><small></small></div>
      <div class="coords">COORDENADAS:<b id="nlat">LAT: …</b><b id="nlon">LON: …</b></div><div id="prog"><i></i></div></div>
    <div id="acq"><div class="ring">📍<span>UBICACIÓN<br>ADQUIRIDA</span></div></div>
    <div id="bar"><span>LAT <b id="blat">-</b></span><span>LON <b id="blon">-</b></span><span class="sp"></span>
      <button class="btn ghost" id="btnBack">↩ Globo</button><button class="btn" id="btnOk">Confirmar ubicación</button></div>
  </div>
</div>
<aside id="side">
  <div class="brand">PROYECTO<b>GEO-SPHERE</b></div>
  <div id="p-select"><div class="tag">Selecciona ubicación inicial</div>
    <div style="font-size:12px;color:var(--mut);line-height:1.5">Haz clic en el globo, escribe un lugar de la lista o coordenadas y pulsa Enter.</div>
    <button class="btn ghost" id="btnTim">📍 Ir a Timbío, Cauca</button></div>
  <div id="p-sel"><div class="tag">Coordenadas seleccionadas</div><div id="selc"></div>
    <button class="btn" id="btnNav">Navegar a la ubicación</button></div>
  <div id="p-map">
    <div id="kpi"><div><b id="k1">0</b>lugares visibles</div><div><b id="k2">0</b>categorías</div></div>
    <div id="cats"></div><div id="list"></div></div>
</aside>
<script>
const D=/*DATA*/[];
const C={};D.forEach(p=>{if(!C[p.c])C[p.c]=p.col});
const $=id=>document.getElementById(id), body=document.body, stageEl=$('stage');
const esc=s=>String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const wrap=l=>((l+540)%360)-180;
let sel=null,map=null,layer=null,selMk=null,mk={},on=new Set(Object.keys(C));

function setStage(s){body.classList.remove('s-select','s-nav','s-map');body.classList.add('s-'+s)}

/* ---------- GLOBO ---------- */
const G=Globe()($('globe'))
  .globeImageUrl('https://cdn.jsdelivr.net/npm/three-globe/example/img/earth-blue-marble.jpg')
  .bumpImageUrl('https://cdn.jsdelivr.net/npm/three-globe/example/img/earth-topology.png')
  .backgroundColor('rgba(0,0,0,0)').atmosphereColor('#3ad7ff').atmosphereAltitude(0.22)
  .pointColor(()=>'#2dd4bf').pointAltitude(0.03).pointRadius(0.55)
  .ringColor(()=>t=>'rgba(45,212,191,'+(1-t)+')').ringMaxRadius(4).ringPropagationSpeed(2.5).ringRepeatPeriod(900)
  .onGlobeClick(({lat,lng})=>{if(body.classList.contains('s-select'))select(lat,lng)});
function fit(){G.width(stageEl.clientWidth).height(stageEl.clientHeight);if(map)map.invalidateSize()}
window.addEventListener('resize',fit);fit();
G.pointOfView({lat:5,lng:-75,altitude:2.4},0);
G.controls().autoRotate=true;G.controls().autoRotateSpeed=0.6;

function select(lat,lng,label){
  sel={lat:lat,lng:wrap(lng)};body.classList.add('has-sel');
  $('selc').innerHTML='LAT: <b>'+sel.lat.toFixed(4)+'</b> | LON: <b>'+sel.lng.toFixed(4)+'</b>'+(label?'<br><small>'+esc(label)+'</small>':'');
  G.pointsData([sel]).ringsData([sel]);
  G.controls().autoRotate=false;
  G.pointOfView({lat:sel.lat,lng:sel.lng,altitude:2.0},900);
}

function navigate(){
  if(!sel)return;
  setStage('nav');G.controls().autoRotate=false;
  const s=G.pointOfView(),t0=performance.now(),dur=7500;
  const d=((sel.lng-s.lng+540)%360)-180, total=720+d;
  const ease=p=>p<.5?4*p*p*p:1-Math.pow(-2*p+2,3)/2;
  function step(now){
    const p=Math.min(1,(now-t0)/dur),e=ease(p);
    const lng=s.lng+total*e,lat=s.lat+(sel.lat-s.lat)*e;
    const alt=p<.6?2.4:2.4-(2.4-0.12)*((p-.6)/.4);
    G.pointOfView({lat:lat,lng:lng,altitude:alt},0);
    $('nlat').textContent='LAT: '+lat.toFixed(3);$('nlon').textContent='LON: '+wrap(lng).toFixed(3);
    document.querySelector('#prog i').style.width=(p*100)+'%';
    if(p<1)requestAnimationFrame(step);else{
      $('acq').classList.add('on');
      setTimeout(()=>{$('acq').classList.remove('on');enterMap()},1300);
    }
  }
  requestAnimationFrame(step);
}

/* ---------- MAPA ---------- */
function initMap(){
  if(map)return;
  map=L.map('map',{zoomControl:true});
  const esri=(p,z)=>L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/'+p+'/MapServer/tile/{z}/{y}/{x}',{maxZoom:z,attribution:'Tiles © Esri'});
  const base={"Satélite":esri('World_Imagery',19),"Calles":esri('World_Street_Map',19),"Oscuro":esri('Canvas/World_Dark_Gray_Base',16)};
  base["Satélite"].addTo(map);L.control.layers(base).addTo(map);
  layer=L.layerGroup().addTo(map);
  D.forEach(p=>{
    const ic=L.divIcon({className:'',iconSize:[38,38],iconAnchor:[19,38],popupAnchor:[0,-34],
      html:'<div class="pin" style="background:'+p.col+'"><i>'+p.i+'</i></div>'});
    const cn=p.cn||{};const nb=['N','S','E','O'].filter(k=>cn[k]).map(k=>'<span><b>'+k+':</b> '+esc(cn[k])+'</span>').join('');
    mk[p.id]=L.marker([p.lat,p.lng],{icon:ic,title:p.n}).bindTooltip(p.n,{direction:'top',offset:[0,-34]})
      .bindPopup('<b>'+p.i+' '+esc(p.n)+'</b><br>'+esc(p.c)+'<br>'+esc(p.d)+(p.o?'<br><i>'+esc(p.o)+'</i>':'')+(nb?'<div class="nb">'+nb+'</div>':''));
  });
  Object.keys(C).forEach(c=>{const e=document.createElement('div');e.className='chip';e.textContent=c;
    e.style.borderColor=e.style.background=C[c];
    e.onclick=()=>{on.has(c)?on.delete(c):on.add(c);e.classList.toggle('off',!on.has(c));draw(false)};$('cats').appendChild(e)});
}
function draw(doFit){
  if(!map)return;
  const q=$('q').value.toLowerCase(),L2=$('list');
  layer.clearLayers();L2.innerHTML='';
  const vis=D.filter(p=>on.has(p.c)&&(p.n+' '+p.d).toLowerCase().includes(q));
  vis.forEach(p=>{layer.addLayer(mk[p.id]);
    const d=document.createElement('div');d.className='it';d.style.borderColor=p.col;
    d.innerHTML='<em>'+p.i+'</em><div><b>'+esc(p.n)+'</b><br><span>'+esc(p.c)+' · '+esc(p.d)+'</span></div>';
    d.onclick=()=>{map.flyTo([p.lat,p.lng],18);mk[p.id].openPopup()};L2.appendChild(d)});
  $('k1').textContent=vis.length;$('k2').textContent=new Set(vis.map(p=>p.c)).size;
  if(doFit&&vis.length)map.fitBounds(L.latLngBounds(vis.map(p=>[p.lat,p.lng])).pad(.2),{maxZoom:17});
}
function enterMap(){
  setStage('map');G.pauseAnimation();initMap();map.invalidateSize();
  map.setView([sel.lat,sel.lng],16);
  if(selMk)map.removeLayer(selMk);
  selMk=L.marker([sel.lat,sel.lng],{icon:L.divIcon({className:'',iconSize:[18,18],iconAnchor:[9,9],html:'<div class="selpin"></div>'})}).addTo(map);
  $('blat').textContent=sel.lat.toFixed(5);$('blon').textContent=sel.lng.toFixed(5);
  $('q').value='';draw(false);
}

/* ---------- EVENTOS ---------- */
$('btnNav').onclick=navigate;
$('btnOk').onclick=()=>draw(true);
$('btnBack').onclick=()=>{setStage('select');G.resumeAnimation();$('q').value='';fit();
  G.pointOfView({lat:sel?sel.lat:5,lng:sel?sel.lng:-75,altitude:2.0},800)};
$('btnTim').onclick=()=>{const la=D.reduce((a,p)=>a+p.lat,0)/D.length,lo=D.reduce((a,p)=>a+p.lng,0)/D.length;select(la,lo,'Timbío, Cauca (centro de los lugares)')};
$('q').addEventListener('input',()=>{if(body.classList.contains('s-map'))draw(false)});
$('q').addEventListener('keydown',async e=>{
  if(e.key!=='Enter')return;
  if(body.classList.contains('s-map')){draw(true);return}
  const v=$('q').value.trim();if(!v)return;
  const m=v.match(/^(-?\d+(?:\.\d+)?)\s*[,;\s]\s*(-?\d+(?:\.\d+)?)$/);
  if(m){select(parseFloat(m[1]),parseFloat(m[2]));return}
  const f=D.find(p=>(p.n+' '+p.d).toLowerCase().includes(v.toLowerCase()));
  if(f){select(f.lat,f.lng,f.n);return}
  try{
    const r=await fetch('https://nominatim.openstreetmap.org/search?format=json&limit=1&q='+encodeURIComponent(v));
    const j=await r.json();
    if(j.length)select(parseFloat(j[0].lat),parseFloat(j[0].lon),j[0].display_name);else alert('No se encontró: '+v);
  }catch(err){alert('No se pudo buscar (¿sin internet?)')}
});
</script></body></html>'''


def main():
    ap = argparse.ArgumentParser(description="GEO-SPHERE · mapa de ubicaciones de Timbío")
    ap.add_argument("--xlsx", help="Excel con los lugares (hoja 'Todas' o la primera)")
    ap.add_argument("--out", default="geo_sphere_timbio.html", help="HTML de salida")
    ap.add_argument("--no-open", action="store_true", help="no abrir el navegador")
    a = ap.parse_args()

    if a.xlsx:
        data = load_xlsx(a.xlsx)
        print(f"Lugares leídos de {a.xlsx}: {len(data)}")
    else:
        data = DATA
        print(f"Lugares incluidos en el script: {len(data)}")

    payload = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    html = HTML.replace("/*DATA*/[]", payload)
    out = Path(a.out).resolve()
    out.write_text(html, encoding="utf-8")
    print("Generado:", out)
    if not a.no_open:
        webbrowser.open(out.as_uri())


if __name__ == "__main__":
    main()
