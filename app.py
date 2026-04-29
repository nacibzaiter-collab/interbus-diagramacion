import os, io, re, json
from datetime import datetime
from flask import Flask, request, render_template, jsonify, send_file, session
from openpyxl import load_workbook

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "interbus-2025-secret")

UPLOAD_FOLDER = "/tmp/interbus"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
MAESTRO_PATH = os.path.join(UPLOAD_FOLDER, "maestro.xlsm")

MESES = ["ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic"]

LEGAJOS_DEFAULT = {
    "2040":"AGUERRE P.","2044":"BELTRAMO J.",
    "2037":"DEL CANTO M.","2035":"DEL CANTO M.",
    "2050":"FERREYRA G.","2011":"GIULIANI M.",
    "2056":"KANCHEFF C.","2041":"LAMBRUSCHI J.",
    "2017":"LASO M.","2057":"LUQUEZ D.",
    "2038":"MURCIA S.","2022":"PAOLLICELLI A.",
    "2048":"PASCASIO G.","2015":"PEREYRA W.",
    "2003":"ROSCHINI M.","2452":"RUIZ H.",
    "2460":"VESPRINI R.","2046":"VOLPINI A.",
    "3003":"PACI G.","2111":"GUARDIA",
}

def es_legajo(v):
    s = str(v or "").strip()
    return bool(re.match(r'^\d{4,5}$', s)) and int(s) > 1000

def es_guardia(v):
    s = str(v or "").strip().upper()
    return s in ("", "****", "GUARDIA", "ALVARELLOS")

def fmt_fecha(f):
    m = re.match(r'^(\d{4})-(\d{2})-(\d{2})', str(f))
    if m:
        return f"{int(m.group(3)):02d}-{MESES[int(m.group(2))-1]}"
    return str(f)

def fmt_larga(f):
    m = re.match(r'^(\d{4})-(\d{2})-(\d{2})', str(f))
    if m:
        return f"{int(m.group(3)):02d} {MESES[int(m.group(2))-1].upper()} {m.group(1)}"
    return str(f).upper()

def parsear_citacion(wb):
    result = {}
    for sn in wb.sheetnames:
        ws = wb[sn]
        fecha = None
        for row in ws.iter_rows(values_only=True):
            r = [str(c or "").strip() for c in row]
            while len(r) < 10:
                r.append("")
            fc = r[4]
            if fc:
                md = re.match(r'^(\d{4}-\d{2}-\d{2})', fc)
                if md:
                    fecha = md.group(1)
                    if fecha not in result:
                        result[fecha] = {"ALVAREZ": {}, "ACEBAL": {}}
                    continue
                md2 = re.match(r'^(\d{2})[/\-](\d{2})[/\-](\d{4})', fc)
                if md2:
                    fecha = f"{md2.group(3)}-{md2.group(2)}-{md2.group(1)}"
                    if fecha not in result:
                        result[fecha] = {"ALVAREZ": {}, "ACEBAL": {}}
                    continue
            if not fecha:
                continue
            svc_a, cond_a, ayud_a = r[0], r[1], r[2]
            svc_c, cond_c, ayud_c = r[6], r[7], r[8]
            if svc_a and re.match(r'^\d+$', svc_a):
                n = int(svc_a)
                if n not in result[fecha]["ALVAREZ"]:
                    result[fecha]["ALVAREZ"][n] = {}
                if es_legajo(cond_a):
                    result[fecha]["ALVAREZ"][n]["PRIMERA"] = cond_a
                if es_legajo(ayud_a):
                    result[fecha]["ALVAREZ"][n]["RELEVO"] = ayud_a
            if svc_c and re.match(r'^\d+$', svc_c):
                n = int(svc_c)
                if n not in result[fecha]["ACEBAL"]:
                    result[fecha]["ACEBAL"][n] = {}
                if not es_guardia(cond_c) and es_legajo(cond_c):
                    result[fecha]["ACEBAL"][n]["PRIMERA"] = cond_c
                if es_legajo(ayud_c):
                    result[fecha]["ACEBAL"][n]["RELEVO"] = ayud_c
    return result

def parsear_slots_maestro(ws):
    rows = list(ws.iter_rows(values_only=True))
    rows = [[str(c or "").strip() for c in r] for r in rows]
    starts = [i for i, r in enumerate(rows) if len(r) > 1 and r[1] == "PLANILLA"]
    slots = []
    for start in starts:
        base = start - 4
        f0 = rows[base]       if base < len(rows) else []
        f4 = rows[base + 4]   if base + 4 < len(rows) else []
        f7 = rows[base + 7]   if base + 7 < len(rows) else []
        f9 = rows[base + 9]   if base + 9 < len(rows) else []
        desc = f0[15] if len(f0) > 15 else ""
        desc_up = desc.upper()
        ruta = "ALVAREZ" if "LVAR" in desc_up else ("ACEBAL" if "ACEBAL" in desc_up else "")
        tipo = ("NOCTURNO" if "NOCTURNO" in desc_up
                else "RELEVO" if "RELEVO" in desc_up
                else "PRIMERA")
        sm = re.search(r'S(\d+)', desc_up)
        snum = int(sm.group(1)) if sm else 1
        dia_nombre = ""
        for d in ["SABADO","DOMINGO","LUNES","MARTES","MIERCOLES","MIÉRCOLES","JUEVES","VIERNES"]:
            if d in desc_up:
                dia_nombre = d
                break
        slots.append({
            "base": base, "start": start, "desc": desc,
            "ruta": ruta, "tipo": tipo, "snum": snum,
            "dia_nombre": dia_nombre,
            "nro_plan": f4[2] if len(f4) > 2 else "",
            "coche":    f4[9] if len(f4) > 9 else "",
            "dia":  f7[1] if len(f7) > 1 else "",
            "mes":  f7[2] if len(f7) > 2 else "",
            "anio": f7[3] if len(f7) > 3 else "",
            "conductor": f9[1] if len(f9) > 1 else "",
            "legajo":    f9[3] if len(f9) > 3 else "",
        })
    return slots

def aplicar_citacion_al_maestro(ws, slots, cit_data, contador_inicio, legajos):
    fechas = sorted(cit_data.keys())
    log = []
    planillas = []
    num = contador_inicio

    # Orden de días tal como aparecen en el Maestro
    dias_en_maestro = []
    seen = []
    for s in slots:
        dn = s["dia_nombre"]
        if dn and dn not in seen:
            seen.append(dn)
            dias_en_maestro.append(dn)

    if len(fechas) > len(dias_en_maestro):
        log.append(f"⚠ Citación tiene {len(fechas)} fechas, Maestro tiene {len(dias_en_maestro)} días.")

    for fi, fecha in enumerate(fechas):
        if fi >= len(dias_en_maestro):
            log.append(f"⚠ Fecha extra sin slot: {fmt_fecha(fecha)}")
            continue

        dia_nombre = dias_en_maestro[fi]
        cit_dia    = cit_data[fecha]

        m = re.match(r'^(\d{4})-(\d{2})-(\d{2})', fecha)
        anio_n = m.group(1) if m else ""
        mes_n  = str(int(m.group(2))) if m else ""
        dia_n  = str(int(m.group(3))) if m else ""

        log.append(f"\n📅 {fmt_larga(fecha)} → día '{dia_nombre}'")

        slots_dia = [s for s in slots if s["dia_nombre"] == dia_nombre]

        for slot in slots_dia:
            ruta  = slot["ruta"]
            snum  = slot["snum"]
            tipo  = slot["tipo"]
            base  = slot["base"]

            # Siempre actualizamos N° planilla y fecha
            ws.cell(row=base+5, column=3,  value=num)
            ws.cell(row=base+8, column=2,  value=dia_n)
            ws.cell(row=base+8, column=3,  value=mes_n)
            ws.cell(row=base+8, column=4,  value=anio_n)

            if tipo == "NOCTURNO":
                cond_actual = slot["conductor"]
                leg_actual  = slot["legajo"]
                log.append(f"  ✓ #{num:05d}  {cond_actual:<18} ({leg_actual}) | {ruta} S{snum} NOCTURNO [sin cambio]")
                planillas.append({"nro":num,"fecha":fmt_fecha(fecha),"conductor":cond_actual,
                                  "legajo":leg_actual,"ruta":ruta,"svc":snum,"tipo":tipo,"coche":slot["coche"]})
                num += 1
                continue

            legajo_cit = cit_dia.get(ruta, {}).get(snum, {}).get(tipo, "")
            conductor_nuevo = legajos.get(legajo_cit, legajo_cit) if legajo_cit else ""

            ws.cell(row=base+10, column=2, value=conductor_nuevo)
            ws.cell(row=base+10, column=4, value=legajo_cit)

            log.append(f"  ✓ #{num:05d}  {conductor_nuevo:<18} ({legajo_cit}) | {ruta} S{snum} {tipo} | Coche {slot['coche']}")
            planillas.append({"nro":num,"fecha":fmt_fecha(fecha),"conductor":conductor_nuevo,
                              "legajo":legajo_cit,"ruta":ruta,"svc":snum,"tipo":tipo,"coche":slot["coche"]})
            num += 1

    log.append(f"\n✅ {len(planillas)} planillas actualizadas. Próxima: #{num}")
    return num, "\n".join(log), planillas

# ════ RUTAS ════════════════════════════════════════════════════════════════

@app.route("/")
def index():
    maestro_ok = os.path.exists(MAESTRO_PATH)
    maestro_info = ""
    if maestro_ok:
        ts = os.path.getmtime(MAESTRO_PATH)
        dt = datetime.fromtimestamp(ts)
        maestro_info = f"Planillas Maestro {dt.month:02d}/{str(dt.year)[-2:]}"
    return render_template("index.html", maestro_ok=maestro_ok, maestro_info=maestro_info)

@app.route("/upload_maestro", methods=["POST"])
def upload_maestro():
    f = request.files.get("maestro")
    if not f:
        return jsonify({"ok": False, "error": "Sin archivo"}), 400
    f.save(MAESTRO_PATH)
    now = datetime.now()
    nombre = f"Planillas Maestro {now.month:02d}/{str(now.year)[-2:]}"
    try:
        wb = load_workbook(MAESTRO_PATH, read_only=True, keep_vba=False)
        sheets = wb.sheetnames
        wb.close()
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500
    return jsonify({"ok": True, "nombre": nombre, "sheets": sheets})

@app.route("/get_sheets")
def get_sheets():
    if not os.path.exists(MAESTRO_PATH):
        return jsonify({"ok": False, "error": "No hay Maestro cargado"}), 400
    try:
        wb = load_workbook(MAESTRO_PATH, read_only=True, keep_vba=False)
        sheets = wb.sheetnames
        wb.close()
        return jsonify({"ok": True, "sheets": sheets})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route("/generar", methods=["POST"])
def generar():
    cit_file     = request.files.get("citacion")
    sheet_name   = request.form.get("sheet_name", "").strip()
    contador     = int(request.form.get("contador", 6800))
    legajos_json = request.form.get("legajos", "{}")

    if not cit_file:
        return jsonify({"ok": False, "error": "Falta la planilla de citación"}), 400
    if not os.path.exists(MAESTRO_PATH):
        return jsonify({"ok": False, "error": "No hay Maestro cargado. Subilo primero."}), 400
    if not sheet_name:
        return jsonify({"ok": False, "error": "Seleccioná el tipo de semana"}), 400
    try:
        legajos = json.loads(legajos_json)
    except:
        legajos = LEGAJOS_DEFAULT

    try:
        cit_wb   = load_workbook(io.BytesIO(cit_file.read()), data_only=True)
        cit_data = parsear_citacion(cit_wb)
        cit_wb.close()

        if not cit_data:
            return jsonify({"ok": False, "error": "No se detectaron fechas en la citación."}), 400

        mae_wb = load_workbook(MAESTRO_PATH, keep_vba=False)
        if sheet_name not in mae_wb.sheetnames:
            return jsonify({"ok": False, "error": f"Pestaña '{sheet_name}' no encontrada"}), 400

        ws = mae_wb[sheet_name]
        slots = parsear_slots_maestro(ws)
        nuevo_contador, log_txt, planillas = aplicar_citacion_al_maestro(
            ws, slots, cit_data, contador, legajos
        )

        buf = io.BytesIO()
        mae_wb.save(buf)
        buf.seek(0)
        mae_wb.close()

        OUT_PATH = os.path.join(UPLOAD_FOLDER, "resultado.xlsx")
        with open(OUT_PATH, "wb") as fout:
            fout.write(buf.read())

        por_fecha = {}
        for p in planillas:
            por_fecha.setdefault(p["fecha"], []).append(p)
        fechas_resumen = [
            {"fecha": k, "cantidad": len(v), "items": v}
            for k, v in sorted(por_fecha.items())
        ]

        return jsonify({
            "ok": True,
            "total": len(planillas),
            "primera": planillas[0]["nro"] if planillas else 0,
            "ultima":  planillas[-1]["nro"] if planillas else 0,
            "nuevo_contador": nuevo_contador,
            "log": log_txt,
            "fechas": fechas_resumen,
            "sheet_name": sheet_name,
        })

    except Exception as e:
        import traceback
        return jsonify({"ok": False, "error": str(e), "trace": traceback.format_exc()}), 500

@app.route("/descargar")
def descargar():
    OUT_PATH = os.path.join(UPLOAD_FOLDER, "resultado.xlsx")
    if not os.path.exists(OUT_PATH):
        return "Sin datos generados", 400
    sheet = request.args.get("sheet", "planillas")
    nombre = f"Planillas_{sheet.replace(' ','_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    return send_file(OUT_PATH, as_attachment=True, download_name=nombre,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
