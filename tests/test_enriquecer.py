# -*- coding: utf-8 -*-
"""Pruebas de enriquecer-datos.py. Sin red: las respuestas de CartoCiudad y de
las webs son inventadas, igual que todos los datos. En este repositorio no
entra nunca un dato real de un cliente."""

import importlib.util
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("enriquecer_datos", RAIZ / "enriquecer-datos.py")
ed = importlib.util.module_from_spec(spec)
sys.modules["enriquecer_datos"] = ed
spec.loader.exec_module(ed)

CARTO = [{
    "type": "portal", "address": "CALLE MAYOR 12, Villaficticia", "tip_via": "CALLE",
    "portalNumber": 12, "postalCode": "29100", "muni": "Villaficticia", "poblacion": "Villaficticia",
    "province": "Málaga", "provinceCode": "29",
}, {
    "type": "portal", "address": "CALLE MAYORAL 2, Otro Sitio", "tip_via": "CALLE",
    "portalNumber": 2, "postalCode": "43892", "muni": "Otro Sitio", "poblacion": "Otro Sitio",
    "province": "Tarragona", "provinceCode": "43",
}]
PORTADA = '<html><a href="/aviso-legal">Aviso legal</a><p>Llámanos: 952 000 111</p></html>'
AVISO = ('<html><p>TALLERES FICTICIOS DEL SUR, S.L. · CIF B-12.345.674 · '
         'Calle Mayor 12, Villaficticia · info@talleresficticios.example</p></html>')


class RedFalsa(ed.Red):
    """Contesta de memoria, y apunta lo que le preguntan."""

    def __init__(self):
        super().__init__(None, activa=True)
        self.preguntas = []

    def descargar(self, url, tipo, max_bytes=600_000):
        self.preguntas.append(url)
        self.consultas[tipo] += 1
        if "cartociudad" in url:
            return "callback(" + json.dumps(CARTO) + ")"
        if url.endswith("/robots.txt"):
            return "User-agent: *\nDisallow: /privado\n"
        if url.rstrip("/").endswith("talleresficticios.example"):
            return PORTADA
        if "aviso-legal" in url:
            return AVISO
        if url.rstrip("/").endswith("gestoria.example"):
            return '<p>GESTORÍA INVENTADA SL · CIF A58818501 · 952 999 888</p>'
        return None


class Piezas(unittest.TestCase):
    def test_clasificar(self):
        self.assertEqual(ed.clasificar("", "Talleres Norte S.L."), "empresa")
        self.assertEqual(ed.clasificar("B12345674", ""), "empresa")
        self.assertEqual(ed.clasificar("12345678Z", "Ana"), "persona")
        self.assertEqual(ed.clasificar("X1234567L", ""), "persona")
        self.assertEqual(ed.clasificar("", "Ana Pérez"), "dudoso")

    def test_provincias(self):
        self.assertEqual(ed.provincia_de_cp("4001"), "04")
        self.assertEqual(ed.provincia_de_cp("99000"), None)
        self.assertEqual(ed.codigo_provincia("VIZCAYA"), "48")
        self.assertEqual(ed.codigo_provincia("Málaga"), "29")

    def test_cif_y_telefono_en_texto(self):
        t = "CIF: B-12.345.674, tel +34 600-111-222 y 952 12 34 56".upper()
        self.assertEqual(["".join(m.groups()) for m in ed.RE_CIF_TEXTO.finditer(t)], ["B12345674"])
        self.assertTrue(ed.cif_valido("B12345674"))
        self.assertFalse(ed.cif_valido("B12345675"))
        tels = [ed.re.sub(r"\D", "", m.group(0))[-9:] for m in ed.RE_TEL.finditer(t)]
        self.assertEqual(tels, ["600111222", "952123456"])

    def test_calle(self):
        self.assertEqual(ed.nucleo_calle("C/ Marqués de Larios, 5 - 2ºB"), "marques larios")
        self.assertEqual(ed.numero_portal("Avda. Andalucía 12, 1º"), 12)


class Cartociudad(unittest.TestCase):
    def test_solo_vale_si_cuadra(self):
        red = RedFalsa()
        cand, conf = ed.cartociudad(red, "C/ Mayor 12", "", "Villaficticia", "")
        self.assertEqual(cand["postalCode"], "29100")
        self.assertEqual(conf, "alta")
        # La misma calle con un CP de otra provincia: ningún candidato vale
        cand, motivo = ed.cartociudad(red, "C/ Mayor 12", "08001", "", "")
        self.assertIsNone(cand)
        self.assertIn("ningún candidato", motivo)


def escribir_xlsx(ruta, filas):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Proveedores"
    for f in filas:
        ws.append(f)
    wb.save(ruta)


class DePuntaAPunta(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.d = Path(self.tmp.name)
        self.fichero = self.d / "proveedores.xlsx"
        escribir_xlsx(self.fichero, [
            ["Nombre", "CIF", "Dirección", "CP", "Población", "Provincia", "Teléfono", "Email", "Web"],
            ["Talleres Ficticios del Sur, S.L.", "", "C/ Mayor 12", "", "Villaficticia", "", "",
             "", "www.talleresficticios.example"],
            ["Ana Pérez Inventada", "12345678Z", "C/ Mayor 12", "", "Villaficticia", "", "", "", ""],
            ["Otra Empresa Inventada SA", "A00000000", "", "29001", "", "Barcelona", "600000000",
             "hola@otra.example", ""],
        ])
        self.cwd = os.getcwd()
        os.chdir(self.d)

    def tearDown(self):
        os.chdir(self.cwd)
        self.tmp.cleanup()

    def test_propuestas(self):
        import pandas as pd
        df = pd.read_excel(self.fichero, dtype=str, keep_default_na=False)
        roles = ed.asignar_roles(list(df.columns), {})
        red = RedFalsa()
        r = ed.enriquecer(df, "Proveedores", str(self.fichero), roles, red, ed.Opciones())
        prop = {(p.fila, p.columna): p for p in r.propuestas}

        self.assertEqual(prop[(2, "CIF")].propuesto, "B12345674")
        self.assertEqual(prop[(2, "CIF")].confianza, "alta", "el nombre está junto al CIF")
        self.assertEqual(prop[(2, "CP")].propuesto, "29100")
        self.assertEqual(prop[(2, "Provincia")].propuesto, "Málaga")
        self.assertEqual(prop[(2, "Teléfono")].propuesto, "952000111")
        self.assertEqual(prop[(4, "Provincia")].confianza, "revisar", "Barcelona no cuadra con el CP 29001")

        self.assertFalse(any(p.fila == 3 for p in r.propuestas), "una persona física no se busca")
        self.assertEqual(r.clases["persona"], 1)
        self.assertFalse(any("privado" in u for u in red.preguntas))

    def test_excel_con_formato_de_auditoria(self):
        codigo = ed.main([str(self.fichero), "--salida", str(self.d / "p"), "--sin-red", "--permitir-git"])
        self.assertEqual(codigo, 0)
        from openpyxl import load_workbook
        wb = load_workbook(next((self.d / "p").glob("Enriquecimiento-*.xlsx")))
        ws = wb["proveedores"]
        cab = [c.value for c in ws[1]]
        for imprescindible in ("Fila", "Columna", "Valor actual", "💡 Valor propuesto", "✏️ Valor corregido",
                               "Decisión"):
            self.assertIn(imprescindible, cab)
        self.assertTrue(any(r[4] == "Barcelona" or r[4] == "Málaga"
                            for r in ws.iter_rows(min_row=2, values_only=True)))

    def test_web_de_otra_empresa(self):
        """El correo es de la gestoría: su CIF no puede salir con confianza alta."""
        import pandas as pd
        df = pd.DataFrame([{"Nombre": "Bodegas Inventadas SL", "CIF": "", "Email": "bodegas@gestoria.example",
                            "Teléfono": ""}])
        r = ed.enriquecer(df, "P", "p.xlsx", ed.asignar_roles(list(df.columns), {}), RedFalsa(), ed.Opciones())
        prop = {p.columna: p for p in r.propuestas}
        self.assertEqual(prop["CIF"].propuesto, "A58818501")
        self.assertEqual(prop["CIF"].confianza, "baja")
        self.assertIn("otra empresa", prop["CIF"].porque)
        self.assertEqual(prop["Teléfono"].confianza, "baja")

    def test_aviso_legal_al_pie(self):
        """Con muchos enlaces delante, el aviso legal se lee igual, y los
        ficheros estáticos no gastan visitas."""
        enlaces = "".join(f'<a href="/empresa/{i}">x</a>' for i in range(500))
        enlaces += '<a href="/wp-content/plugins/legal.css">c</a><a href="/aviso-legal">Aviso legal</a>'

        class Red2(RedFalsa):
            def descargar(self, url, tipo, max_bytes=600_000):
                self.preguntas.append(url)
                if url.endswith("robots.txt"):
                    return ""
                if url.endswith("aviso-legal"):
                    return AVISO
                return "<html>" + enlaces + "</html>"
        red = Red2()
        d = ed.leer_web(red, "https://talleresficticios.example")
        self.assertIn("B12345674", d["cifs"])
        self.assertFalse(any(u.endswith(".css") for u in red.preguntas))

    def test_nombre_palabras_enteras(self):
        self.assertFalse(ed.nombre_aparece("Ron SL", "Ronda del sur"))
        self.assertTrue(ed.nombre_aparece("Bodegas Inventadas, S.L.", "BODEGAS INVENTADAS SL, CIF..."))

    def test_sin_red_no_pregunta(self):
        red = ed.Red(None, activa=False)
        self.assertIsNone(red.get("https://example.invalid/", "web"))
        self.assertEqual(sum(red.consultas.values()), 0)


if __name__ == "__main__":
    unittest.main()
