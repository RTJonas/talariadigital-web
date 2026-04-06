# Talaria Digital — Contexto del Proyecto

## Descripción General
Agencia boutique B2B radicada en España (Pinseque, Zaragoza). Presta servicios de consultoría, gestión administrativa y liquidación de pagos para **creadoras de contenido digital residentes en Argentina**.

**Titular:** Jonas Gabriel Ramirez Torres | **NIF:** Z3318599A  
**Email:** info@talariadigital.es  
**Estado actual:** Operativo. 3 creadoras activas, liquidación semanal, facturación mensual.

---

## Modelo de Negocio

### Flujo de Capital
1. **Cobros (Frontend):** Revolut Pro (CHF, GBP), BBVA Negocios (Paygold/Bizum Empresas), Stripe
2. **Retención:** 20% honorarios de agencia sobre el neto
3. **Conversión:** El 80% restante se convierte a USDT vía Kraken (transferencia SEPA en EUR)
4. **Liquidación:** USDT enviado a wallets de las creadoras en Argentina

> **Nota Compliance:** De cara a bancos y web pública, Talaria opera como agencia B2B corporativa estándar. No se menciona USDT, cripto ni contenido de alto riesgo en ninguna comunicación pública.

---

## Infraestructura Contable y Fiscal

- **IVA:** Facturas al 0% (Art. 69 Ley 37/1992) — destinatarios fuera de la UE (Argentina)
- **IRPF:** Sin retención (destinatario no residente). Internamente se reserva el 20% para Modelo 130 trimestral
- **Gastos deducibles:** Solo los estrictamente justificados con documento (ej. mobiliario pagado por banco). Sin dietas diarias

---

## El "Cerebro" Tecnológico

### Archivo Principal
`Talaria Ledger 2026.ods` — vive en Google Drive (copia local en la raíz de este repo para versionado)

### Pestañas del Spreadsheet
| Pestaña | Función |
|---|---|
| `Transacciones` | Registro diario de ingresos y gastos. Estado: `Pendiente` → `Procesado` |
| `Semanal - Liquidacion` | Panel de control. Datos de creadoras (Col O=ID, Q=Email, S=Nombre, T=CUIT). Historial tipo de cambio en cols K-L |
| `Historial de Pagos` | Resumen consolidado por liquidación (9 columnas: fecha, ID, bruto, comisión, honorarios, a liquidar, USDT, tasa, reserva hacienda) |
| `Fondeo Kraken` | Registro de envíos SEPA a Kraken |
| `Registro Facturas` | Numeración correlativa de facturas |

### Columnas clave de `Transacciones`
A=Fecha/Hora, B=Origen/Tipo, C=ID Transacción, D=Email Cliente, E=ID Proveedor, F=Bruto, G=Comisión, H=Neto(fórmula), I=Honorarios(fórmula), J=A Liquidar(fórmula), K=Fecha Liquidación, L=Estado Liquidación, M=USDT

### Apps Script — Funciones activas
Todas están en `scripts/talaria-ledger.gs` y activadas por botones en el spreadsheet.

| Función | Qué hace |
|---|---|
| `cargarGasto()` | Registra un gasto manual (con ID creadora, monto, concepto, referencia). Descuenta al balance de la creadora |
| `cargarTipoCambio()` | Pide el precio EUR/USDT de Kraken y calcula los USDT exactos en transacciones pendientes (divide por la tasa) |
| `actualizarLiquidaciones()` | Cierra semana: marca como `Procesado`, protege filas, escribe en Historial, genera saldos deudores si hay balance negativo, registra fondeo Kraken |
| `generarExtractoPDF()` | Genera PDFs por creadora con desglose de transacciones. Los guarda en carpeta `Extractos Semanales` de Drive |
| `enviarExtractosPorEmail()` | Envía los PDFs generados por email a cada creadora. Anota fecha de envío en col R de `Semanal - Liquidacion` |
| `generarFacturasMensuales()` | Suma honorarios del mes, genera factura PDF legal con datos fiscales, los guarda en carpeta `Facturas` de Drive y registra en `Registro Facturas` |
| `enviarFacturasPorEmail()` | Envía las facturas mensuales generadas a los emails de las creadoras |

---

## Automatizaciones e Integraciones

- **Make.com:** Guardado automático de documentos en Google Drive. Conexión via proyecto propio en Google Cloud Console (en producción, no en modo test) para evitar que los tokens de Gmail expiren cada 7 días
- **Telegram Bot:** Notificaciones directas a creadoras vía Make — pueden consultar saldos acumulados y avisos de pago de forma privada
- **Google Cloud Console:** Proyecto propio configurado para las integraciones de Make

---

## Web y Presencia Digital

- Dominio propio (migrado desde Carrd.co para pasar auditoría Redsys/BBVA)
- Footer con marco legal europeo completo: Aviso Legal, Política de Privacidad, Cookies, T&C
- **Próximo:** Agregar Paygold/Bizum Empresas a la web para procesar pagos de clientes

---

## Contratos y KYC

- Contrato firmado con cada creadora
- Nombre completo + CUIT (Argentina) registrados en la matriz del spreadsheet
- 3 creadoras activas actualmente (IDs en el spreadsheet)

---

## Estructura del Repositorio

```
TalariaDigital-General/
├── CLAUDE.md                    ← Este archivo (contexto para IA)
├── DEVLOG.md                    ← Bitácora de cambios diarios
├── .gitignore
├── Talaria Ledger 2026.ods      ← Copia local del spreadsheet (referencia)
└── scripts/
    └── talaria-ledger.gs        ← Código Apps Script completo
```

---

## Flujo de Trabajo Semanal Tipo

```
Lunes-Viernes: Se registran transacciones en "Transacciones" (manual o via Stripe/Revolut)
Viernes:
  1. cargarTipoCambio()         → Ingresar precio Kraken del día
  2. actualizarLiquidaciones()  → Cerrar semana, fondear Kraken si aplica
  3. generarExtractoPDF()       → Generar PDFs (dejar vacío = TODOS)
  4. enviarExtractosPorEmail()  → Enviar a creadoras
```

```
Fin de mes:
  1. generarFacturasMensuales() → Generar facturas
  2. enviarFacturasPorEmail()   → Enviar facturas
  Presentar Modelo 130 con el fondo IRPF reservado
```
