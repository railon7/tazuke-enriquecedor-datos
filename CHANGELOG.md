# Cambios

## 1.1.1 · 2026-09-27

- **Un listado de documentos no se enriquece.** Probado en un proyecto real,
  trataba un listado de 958 facturas como un maestro de 584 «empresas»: el
  nombre del cliente se repite en cada factura. Si los valores distintos no
  llegan a la mitad de las filas, se salta con una nota

## 1.1.0 · 2026-09-27

Sale de revisar extractores de *Impressum* alemanes, de *mentions légales*
francesas, extruct, las fuentes abiertas españolas y cómo empareja nombres la
gente que más lo ha hecho (japoneses, chinos, rusos).

- **VIES como verificador.** Para España no devuelve el nombre, pero compara
  el que se le manda: cada CIF propuesto se confirma (alta) o se degrada
  (baja). Con la forma jurídica abreviada, porque con «Sociedad Anónima» dice
  que no. `--vies` comprueba también los CIF que ya hay
- **`--bdns`: el CIF a partir del nombre**, en la Base de Datos Nacional de
  Subvenciones. Tres filtros: todas las palabras enteras, la letra del CIF
  cuadrando con la forma jurídica, y un único candidato claro. Probándolo,
  «Inditex SA» salía como «BLINDITEX, S.L.» con un 87 % de parecido: por eso
  los filtros. La BDNS busca por el principio de la frase y no casa la Ñ: se
  le pregunta hasta la primera Ñ. Opcional, por su aviso legal
- **La web por bloques**: el CIF del pie que firma la agencia no cuenta; el
  del bloque con «Tomo, Folio, Hoja» o «titular» gana a los demás, y en la
  misma línea pesa el doble que en la de al lado
- **Un CIF o un teléfono en las webs de dos empresas del fichero** no se
  propone: es de la agencia, del alojamiento o del grupo
- **schema.org (JSON-LD)**: teléfono, CP y CIF que la web declara de sí misma
- **Pestaña «Comprobaciones»**: letra del CIF frente a forma jurídica, prefijo
  del fijo frente al CP, y lo que diga VIES de los CIF que ya hay
- Rastreo: `Crawl-delay`, parar con un servidor que responde 429 o 503, y no
  esquivar las protecciones contra robots (se anotan en el resumen)
- Las páginas legales, también en catalán, gallego y euskera

Probado contra un maestro real de 678 interlocutores (566 empresas), con
`--vies --bdns`: 199 propuestas en 14 minutos (505 consultas a VIES a una por
segundo). La BDNS encontró el CIF de 8 empresas que no lo tenían, 5 confirmados
por VIES. Y una lección: **VIES solo reconoce la razón social casi completa**.
111 de los 390 CIF que conoce no casaban con el nombre del maestro, casi todos
por ser nombres comerciales o razones sociales recortadas. Por eso, para los
CIF que ya hay, que VIES no case es un aviso y no un error.

## 1.0.0 · 2026-09-26

Primera versión, sin tokens: fuentes públicas y gratuitas, y la propia web de
cada empresa.

- Provincia desde el código postal, y aviso cuando la del registro no cuadra
- CP, municipio y provincia desde la dirección con CartoCiudad (IGN), **solo
  si el candidato cuadra con el CP, la población o la provincia del
  registro**: a «c/ Larios 5, 29005» (Málaga) le devolvió una calle de
  Tarragona
- CIF, teléfono y correo desde la web de la empresa, empezando por el aviso
  legal, que la LSSI obliga a publicar con el NIF
- Web desde el dominio del correo, si no es genérico
- VIES opcional (`--vies`): para España solo dice sí o no
- Solo empresas: una persona física no se busca ni sale del ordenador
- El Excel de propuestas tiene el formato del de auditoría, y
  `limpiar-datos.py` 1.1.0 aplica lo aceptado

Lo que salió al probarlo contra un maestro real de 678 interlocutores:

- Un CIF de confianza alta sacado de una web que no era la de la empresa
  (el correo era de otro dominio). Resultó ser su marca comercial y el CIF era
  bueno, pero por casualidad: la comprobación del nombre buscaba trozos de
  texto. Ahora es alta solo si **todas las palabras del nombre, enteras,
  están junto al CIF**
- El aviso legal suele ser el último enlace de la página, y se perdía detrás
  de los demás: ahora los enlaces se ordenan por prioridad antes de elegir
- La pausa global de un segundo convertía el maestro en nueve minutos: ahora
  es por servidor, con ocho webs distintas a la vez y nunca dos consultas a la
  misma. Tres minutos
- La caché guardaba páginas enteras (37 MB): ahora solo su texto y los
  enlaces útiles
