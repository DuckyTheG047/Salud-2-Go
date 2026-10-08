SALUD 2 GO — V20.8

Fix definitivo del CTA final "Quiero recibir información":
- El modal se abre directamente manipulando el DOM y clonando waitlistTemplate.
- No depende del handler openModal original.
- Funciona también desde el preview/local; NO es necesario subir a GitHub para que aparezca.
- Conserva cta_origen = resultado / mobile_cta.
- No requiere cambios nuevos en Apps Script respecto a V20.6.
