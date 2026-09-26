# Cambios

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
