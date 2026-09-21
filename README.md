# Análisis Descriptivo con Excel — Kiva Crowdfunding

Proyecto individual de análisis de datos en Microsoft Excel sobre el dataset **Data Science for Good: Kiva Crowdfunding** (préstamos de microfinanciación).
El trabajo cubre el ciclo completo: importación y transformación con Power Query, modelado y análisis con tablas dinámicas, y construcción de un dashboard interactivo.

**Autor:** Franco Recabarren
**Curso:** Bootcamp Data Analytics P3 — Factoría F5 Madrid

---

## Contenido del repositorio

| Archivo | Descripción |
|---|---|
| `kiva_loans_proyecto_1.xlsx` | Libro de trabajo completo: datos, modelo, análisis y dashboard |
| `kiva_loans.csv` | Datos de origen (fuente del proyecto) |
| `README.md` | Este documento |

---

## El dataset

- **Origen:** Kiva Crowdfunding (préstamos de microfinanciación concedidos a través de la plataforma Kiva).
- **Volumen tras la limpieza:** 88.140 registros y 26 columnas en la tabla `kiva_loans`.
- **Periodo cubierto:** 2014 – 2017.
- **Campos principales:** `id`, `country`, `region`, `sector`, `activity`, `use`, `funded_amount`, `loan_amount`, `currency`, `partner_id`, `posted_time`, `disbursed_time`, `funded_time`, `term_in_months`, `lender_count`, `tags`, `borrower_genders`, `repayment_interval`, `date`.

---

## Estructura del libro

| Hoja | Contenido |
|---|---|
| `kiva_loans` | Tabla principal cargada y transformada desde Power Query, en formato de tabla de Excel (`kiva_loans`) |
| `Calendario` | Tabla de fechas generada con Power Query (2014-01-01 → 2017-12-31, 1.461 días) con columnas `Fecha`, `Año`, `Mes`, `Nombre_Mes` |
| `Estadistica Descriptiva` | Resumen estadístico de `funded_amount` y `loan_amount` generado con la herramienta *Análisis de datos* de Excel |
| `Tabla Dinamica` | Tablas dinámicas que alimentan los gráficos y los KPIs del dashboard |
| `Dashboard` | Panel interactivo final (líneas de cuadrícula ocultas, KPIs, gráficos y segmentaciones) |
| `test`, `Hoja1` | Hojas auxiliares de trabajo: pruebas de medidas por año y volcado de detalle del modelo |

---

## 1. Importación y transformación (Power Query)

Consulta **`kiva_loans`** — pasos aplicados:

1. Importación del CSV (delimitador `,`, codificación UTF-8) y promoción de encabezados.
2. Tipado explícito de columnas: enteros, moneda, texto, `datetime` y `date`.
3. Limpieza de la columna `tags`: separación, eliminación de duplicados dentro de cada celda y recomposición del texto.
4. Expansión de `borrower_genders` por delimitador a filas individuales.
5. Filtrado de filas con `loan_amount > 0` y eliminación de filas con errores en `funded_amount`.
6. Columna personalizada `Personalizado` = `funded_amount` / `loan_amount` (ratio de financiación).
7. Agrupación por `sector` y `country` para calcular las columnas agregadas `Total_Prestamos` (nº de préstamos) y `Monto_Total` (suma financiada), reincorporadas después al detalle.
8. Reordenación de columnas y ordenación cronológica por `date`.

Consulta **`Calendario`** — tabla de fechas creada por código M a partir de `#date(2014,1,1)` y `#date(2017,12,31)`, ampliada con `Año`, `Mes` y `Nombre_Mes` para el análisis temporal.

---

## 2. Análisis

**Columna calculada en la tabla**

```excel
Categoria_Monto = SI(D2>4375;"Alto";SI(D2>1000;"Medio";"Bajo"))
```

Clasifica cada préstamo en *Bajo / Medio / Alto* según el importe financiado.

**Estadística descriptiva** (`funded_amount`, 88.140 registros)

| Medida | Valor |
|---|---|
| Media | 1.692,56 |
| Mediana | 1.000 |
| Moda | 200 |
| Desviación estándar | 1.761,52 |
| Mínimo / Máximo | 25 / 50.000 |
| Rango | 49.975 |
| Coeficiente de asimetría | 2,22 |
| Curtosis | 20,44 |
| Suma | 149.182.200 |

La distribución es claramente asimétrica a la derecha: la mayoría de los préstamos son pequeños (mediana 1.000, moda 200) y una minoría de importes muy altos tira de la media hacia arriba.

**Modelo de datos (Power Pivot)**

Las tablas `kiva_loans` y `Calendario` se cargan al modelo de datos, sobre el que se definen las medidas:

- `Total_Financiado`
- `Prestamos`
- `Ticket_Promedio`

**Tablas dinámicas** (hoja `Tabla Dinamica`)

- Importe financiado por **sector**.
- Plazo medio (`term_in_months`) por **país**.
- Importe financiado por **país × tipo de devolución** (`repayment_interval`).
- Tabla de apoyo para los KPIs del dashboard.

---

## 3. Dashboard

Hoja `Dashboard` — **Kiva Loans Analysis**.

**KPIs dinámicos:** cuadros de texto vinculados por fórmula a celdas de las tablas dinámicas, de modo que se actualizan al filtrar:

- `TOTAL LOANS` → `'Tabla Dinamica'!$A$2`
- `SELECTED COUNTRY` → `'Tabla Dinamica'!$C$2`
- `AVG TERM (MONTHS)` → `'Tabla Dinamica'!$D$2`

**Gráficos:**

- *Funding Share by Sector* — gráfico de anillo con el reparto del importe financiado por sector.
- *Funded Amount by Repayment Type* — gráfico de barras por país y tipo de devolución.

**Interactividad:** dos segmentaciones de datos, **Año** (tabla `Calendario`) y **country** (tabla `kiva_loans`), conectadas a las cuatro tablas dinámicas del libro, de forma que un clic actualiza simultáneamente KPIs y gráficos.

**Presentación:** líneas de cuadrícula ocultas, títulos y tarjetas con formas redondeadas para dar aspecto de aplicación.

---

## Cómo usarlo

1. Abrir `kiva_loans_proyecto_1.xlsx` con Microsoft Excel (365 / 2021 o superior; requiere Power Query y Power Pivot).
2. Situarse en la hoja `Dashboard` y usar las segmentaciones **Año** y **country** para filtrar.
3. Para volver a cargar los datos: *Datos → Actualizar todo*.

### Actualizar la fuente de datos

La consulta `kiva_loans` apunta a una ruta local del equipo donde se construyó el libro. Para usar el CSV incluido en este repositorio:

1. *Datos → Consultas y conexiones → kiva_loans → Editar*.
2. En el paso **Origen**, sustituir la ruta por la de `kiva_loans.csv` de este repositorio.
3. *Cerrar y cargar* y, después, *Actualizar todo*.

El resto de pasos de transformación se reaplican automáticamente, por lo que el libro admite una actualización de datos con una intervención mínima.

---

## Herramientas

Microsoft Excel (Power Query · Power Pivot / modelo de datos · tablas dinámicas · segmentaciones · herramienta de Análisis de datos) · Git / GitHub.
