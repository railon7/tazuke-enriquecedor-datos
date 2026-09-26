# Enriquecedor de datos

Propone los datos que le faltan a un maestro de empresas —CIF, código postal,
municipio, provincia, teléfono, correo, web— buscándolos en **fuentes públicas
y gratuitas, sin gastar tokens**, y los deja en un Excel para que el cliente
los acepte o los rechace.

Un solo fichero, `enriquecer-datos.py`. Python 3.10+ con `pandas` y `openpyxl`:

```bash
pip install -r requirements.txt
```

## Propone, no rellena

**No cambia ningún dato.** Saca un Excel con una fila por dato encontrado:
el valor propuesto, la confianza, la fuente con su enlace y por qué. El cliente
marca «Se acepta la propuesta» o escribe el valor bueno, y
[`limpiar-datos.py`](https://github.com/railon7/tazuke-limpiador-datos) aplica
lo aceptado.

Un CIF que parece bueno y es de otra empresa es el peor error de una migración:
está bien formado, pasa todas las validaciones y factura a quien no es. Por eso
nada entra sin que alguien lo mire.

## Solo empresas

Los datos de una sociedad son públicos; los de una persona física, no, y el
RGPD no deja buscarlos sin motivo. **Una fila con DNI o NIE, o sin
identificador y sin forma jurídica en el nombre, no se busca ni sale del
ordenador.** Con `--incluir-dudosos` se buscan también las filas sin nada que
las identifique: úsalo sabiendo lo que haces.

## De dónde sale cada dato

| Qué | Fuente | Sale del ordenador |
|---|---|---|
| Provincia | Las dos primeras cifras del código postal | No |
| Provincia que no cuadra con el CP | Lo mismo: se marca **REVISAR** | No |
| CP, municipio, provincia | [CartoCiudad](https://www.cartociudad.es) (Instituto Geográfico Nacional), desde la dirección | La dirección |
| CIF, teléfono, correo | La web de la empresa. La LSSI la obliga a publicar NIF, razón social y domicilio en el aviso legal | La URL |
| Web | El dominio de su correo, si no es genérico (gmail, hotmail…) | El dominio |
| ¿Operador intracomunitario? | [VIES](https://ec.europa.eu/taxation_customs/vies/), con `--vies`. Para España solo dice sí o no: no da nombre ni dirección | El CIF |

## Confianza

| Confianza | Cuándo |
|---|---|
| **ALTA** | CP → provincia · dirección de CartoCiudad con el portal exacto y que cuadra con el CP, la población o la provincia · CIF que aparece en la web **junto al nombre de la empresa**, con sus palabras enteras |
| **MEDIA** | Dirección que cuadra pero sin portal exacto · dato de una web donde aparece el nombre de la empresa |
| **BAJA** | Dato de una web donde **no aparece su nombre**: el correo puede ser de la gestoría, del grupo o de quien le lleva la informática |
| **REVISAR** | Dos datos del registro que se contradicen (provincia y CP) |

CartoCiudad es tolerante con las erratas, y a veces demasiado: «c/ Larios 5,
29005» (Málaga) le devolvió una calle «Larix» en Tarragona. Por eso **un
candidato solo vale si cuadra con lo que ya sabemos del registro**: el CP, la
población o la provincia. Si no cuadra, se descarta y queda anotado.

## Uso

```bash
python enriquecer-datos.py proveedores.xlsx --salida propuestas/
python enriquecer-datos.py clientes.csv --columna "nombre=Razón social" --columna "cp=C.P."
python enriquecer-datos.py clientes.csv --sin-red      # solo lo que no sale del ordenador
python enriquecer-datos.py clientes.csv --simular      # cuenta, no escribe
python enriquecer-datos.py --help
```

Las columnas se reconocen solas por el nombre, también las de SAP Business One
(`CardName`, `LicTradNum`, `ZipCode`, `City`, `County`, `Phone1`, `E_Mail`,
`IntrntSite`). Lo que no se reconozca se indica con `--columna rol=Columna`.

Dentro de un proyecto del **Método de Implantación Tazuke**, desde su raíz y
sin argumentos, lee `06-Migracion/Maestros-Limpios/` y deja
`06-Migracion/Enriquecimiento-<fecha>.xlsx`.

**Mejor sobre datos ya limpios**: con los CIF sin puntos y los vacíos como
vacíos, encuentra más y se equivoca menos.

## Cortesía y límites

- Una pausa por servidor: 1 s entre páginas de la misma web, 0,3 s en
  CartoCiudad. Se identifica con su nombre en cada consulta
- **Respeta el `robots.txt`** de cada web, y lee como mucho cuatro páginas por
  web: la portada y las de aviso legal o contacto
- **Caché**: lo consultado queda en `.cache-enriquecer.json` en la carpeta de
  salida y no se vuelve a preguntar. Lleva direcciones y webs del cliente:
  **no se publica**, y el `.gitignore` de este repositorio lo excluye
- Con varios cientos de empresas, cuenta con unos minutos: van saliendo por
  pantalla de cincuenta en cincuenta

## Lo que todavía no hace

**Buscar un CIF solo con el nombre**, cuando no hay web ni correo propio. No
hay fuente oficial gratuita que lo dé. La siguiente versión lo hará con una API
de búsqueda de cuota gratuita, sin tokens, y siempre como propuesta con enlace.

## Los datos del cliente

- No escribe en una carpeta `02-Datos-Origen`, ni dentro de un repositorio git
  si la salida no está ignorada
- Por pantalla solo salen recuentos
- Las pruebas no salen a la red y usan datos inventados. **En este repositorio
  no entra nunca un dato real**, ni el nombre de un cliente

## Pruebas

```bash
python -m unittest discover -s tests -v
```

## Relación con el método

Este repositorio es **la fuente**. El Método de Implantación Tazuke lleva una
copia idéntica en `_herramientas/scripts/enriquecer-datos.py`. Un cambio se
hace aquí y se copia allí; `check-metodo.sh` avisa si difieren.
