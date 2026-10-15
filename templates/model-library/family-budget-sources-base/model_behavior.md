# Contrato funcional del modelo · Base de presupuesto familiar

> Este documento es normativo: describe lo que la plantilla debe hacer, no lo que se observó en un XLSX ya construido.
> Si el libro no cumple una regla, el libro falla la verificación y debe corregirse el diseño o la implementación.

## Guía rápida para el usuario

**Objetivo:** Registrar movimientos una sola vez, comparar presupuesto y gasto real por mes y categoría, y convertir el resultado en flujo de caja, metas y deudas.

1. Introduce o modifica únicamente las columnas marcadas como `input`.
2. No sobrescribas columnas `derived`: contienen fórmulas y deben actualizarse.
3. Usa las listas y estados declarados por la ficha del producto; no escribas porcentajes en columnas de cantidades o importes.
4. Comprueba el resultado en las hojas de resumen después de añadir un registro.

## Hojas que debe contener el libro

`Inicio` · `Resumen` · `Presupuesto` · `Gastos reales` · `Movimientos` · `Flujo de caja` · `Metas y fondo` · `Deudas` · `Listas`

## Tablas y columnas esperadas

Cada rango es una tabla de datos. Las filas de total indicadas después del rango deben quedar fuera del ListObject y fuera del filtro de entradas.

| Hoja | Rango de datos | Columnas esperadas |
|---|---|---|
| `Resumen` | `B11:F18` | `Categoría` (input, text); `Plan` (derived, amount); `Real` (derived, amount); `Diferencia` (derived, amount); `% gastado` (derived, percent). |
| `Presupuesto` | `B7:O14` | `Categoría` (input, text); `Enero` (input, amount); `Febrero` (input, amount); `Marzo` (input, amount); `Abril` (input, amount); `Mayo` (input, amount); `Junio` (input, amount); `Julio` (input, amount); `Agosto` (input, amount); `Septiembre` (input, amount); `Octubre` (input, amount); `Noviembre` (input, amount); `Diciembre` (input, amount); `Total anual` (derived, amount). Total fuera: `B15:O15` (`TOTAL DEL HOGAR`). |
| `Gastos reales` | `B7:O14` | `Categoría` (derived, text); `Enero` (derived, amount); `Febrero` (derived, amount); `Marzo` (derived, amount); `Abril` (derived, amount); `Mayo` (derived, amount); `Junio` (derived, amount); `Julio` (derived, amount); `Agosto` (derived, amount); `Septiembre` (derived, amount); `Octubre` (derived, amount); `Noviembre` (derived, amount); `Diciembre` (derived, amount); `Total anual` (derived, amount). Total fuera: `B15:O15` (`TOTAL REAL`). |
| `Movimientos` | `B7:I27` | `Fecha` (input, date); `Mes` (input, text); `Tipo` (input, text); `Categoría` (input, text); `Concepto` (input, text); `Importe` (input, amount); `Pagado por` (input, text); `Estado` (input, text). |
| `Flujo de caja` | `B7:J19` | `Mes` (derived, text); `Ingresos plan` (derived, amount); `Ingresos real` (derived, amount); `Gastos plan` (derived, amount); `Gastos real` (derived, amount); `Ahorro plan` (derived, amount); `Ahorro real` (derived, amount); `Disponible` (derived, amount); `Comentarios` (input, text). |
| `Metas y fondo` | `B9:J12` | `Meta` (input, text); `Gasto mensual` (input, amount); `Meses` (input, count); `Meta necesaria` (derived, amount); `Ahorrado` (input, amount); `Pendiente` (derived, amount); `Progreso` (derived, percent); `Fecha objetivo` (input, date); `Estado` (derived, text). |
| `Deudas` | `B7:H10` | `Deuda` (input, text); `Saldo actual` (input, amount); `Interés anual` (input, percent); `Pago mensual` (input, amount); `Meses restantes` (input, count); `Interés estimado` (derived, amount); `Estado` (derived, text). |
| `Listas` | `B7:C13` | `Parámetro` (input, text); `Valor` (input, general). |

## Reglas de cálculo

- Movimientos es la única entrada operativa para gastos reales y flujo de caja.
- Presupuesto y Gastos reales comparten categorías y meses.
- Las filas TOTAL quedan fuera del rango de datos de cada tabla.
- Las columnas de conteo nunca se expresan como porcentaje.

### Fórmulas que deben existir

- `Resumen!E12` debe contener `C12-D12`: La diferencia es presupuesto menos gasto real.
- `Resumen!F12` debe contener `D12/C12`: El porcentaje compara gasto real con presupuesto.
- `Gastos reales!C8` debe contener `SUMIFS`: El gasto real nace de Movimientos y no se duplica manualmente.
- `Flujo de caja!I8` debe contener `D8-F8-H8`: Disponible = ingreso real - gasto real - ahorro real.
- `Metas y fondo!G10` debe contener `MAX(E10-F10,0)`: Pendiente nunca puede ser negativa.
- `Metas y fondo!H10` debe contener `F10/E10`: Progreso = ahorrado / meta necesaria.

## Casos de verificación

Los casos se ejecutan sobre una copia temporal del libro. El archivo original no se modifica.

### Caso 1: income flows into monthly cash flow

Comprueba que el ingreso neto declarado alimenta la lectura mensual.

**Entrada:** `Listas!C9 = 3500`

**Resultados esperados:**
- `Flujo de caja!C8` equals `3500`

## Criterios de fallo

- Una tabla tiene una cabecera vacía, genérica, duplicada o con un significado distinto.
- Una fila TOTAL queda dentro del rango filtrable de datos.
- Gastos reales deja de depender de Movimientos o Disponible deja de seguir su ecuación.
- Una columna de conteo usa formato porcentaje o una columna monetaria usa porcentaje.

## Regla de independencia

Este documento se actualiza desde el contrato de diseño declarado antes de generar el XLSX. No se debe redactar leyendo hacia atrás el libro generado ni cambiar el resultado esperado para que coincida con un error observado.
