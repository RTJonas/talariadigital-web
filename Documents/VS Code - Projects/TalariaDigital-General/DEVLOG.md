# DEVLOG — Talaria Digital

Bitácora cronológica de cambios, mejoras y decisiones técnicas.  
Formato: `## YYYY-MM-DD — Descripción breve` + detalle de lo hecho.

---

## 2026-04-06 — Inicialización del repositorio

- Creado repo privado en GitHub: `TalariaDigital-General`
- Establecida estructura base del proyecto (CLAUDE.md, DEVLOG.md, .gitignore, scripts/)
- Documentados todos los scripts de Apps Script existentes en `scripts/talaria-ledger.gs`
- Agregado `Talaria Ledger 2026.ods` como referencia local del spreadsheet
- Contexto del proyecto capturado en `CLAUDE.md` para carga automática en sesiones futuras

### Estado del sistema al inicio
- 3 creadoras activas, todas operativas
- Liquidación semanal funcionando
- Facturación mensual funcionando
- Make.com + Telegram activos
- Web en dominio propio (pendiente: agregar Paygold/Bizum Empresas)
- Apps Script: 7 funciones activas (`cargarGasto`, `cargarTipoCambio`, `actualizarLiquidaciones`, `generarExtractoPDF`, `enviarExtractosPorEmail`, `generarFacturasMensuales`, `enviarFacturasPorEmail`)
