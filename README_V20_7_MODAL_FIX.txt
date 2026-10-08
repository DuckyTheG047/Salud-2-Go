SALUD 2 GO — V20.7

Fix:
- "Quiero recibir información" vuelve a abrir el modal correctamente.
- "Participar en el lanzamiento" mobile usa el mismo handler robusto.
- Se conserva el tracking de cta_origen.
- No requiere ningún cambio adicional en Apps Script respecto a V20.6.

Causa:
El CTA del resultado puede volver a renderizarse después de la carga inicial.
El handler original dependía de bindings previos. V20.7 usa un listener delegado
en capture para abrir siempre el modal de waitlist.
