# Contrato funcional del modelo · Base de planificador de boda

> Este documento es normativo: describe lo que la plantilla debe hacer, no lo que se observó en un XLSX ya construido.
> Si el libro no cumple una regla, el libro falla la verificación y debe corregirse el diseño o la implementación.

## Guía rápida para el usuario

**Objetivo:** Comparar presupuesto y pagos, seguir proveedores, invitados y tareas sin duplicar entradas.

1. Introduce o modifica únicamente las columnas marcadas como `input`.
2. No sobrescribas columnas `derived`: contienen fórmulas y deben actualizarse.
3. Usa las listas y estados declarados por la ficha del producto; no escribas porcentajes en columnas de cantidades o importes.
4. Comprueba el resultado en las hojas de resumen después de añadir un registro.

## Hojas que debe contener el libro

`Inicio` · `Resumen` · `Presupuesto` · `Pagos reales` · `Proveedores` · `Flujo de pagos` · `Checklist` · `Invitados` · `Listas`

## Tablas y columnas esperadas

Cada rango es una tabla de datos. Las filas de total indicadas después del rango deben quedar fuera del ListObject y fuera del filtro de entradas.

| Hoja | Rango de datos | Columnas esperadas |
|---|---|---|
| `Resumen` | `B11:F18` | `Categoría` (derived, text); `Plan` (derived, amount); `Pagado` (derived, amount); `Pendiente` (derived, amount); `% pagado` (derived, percent). |
| `Presupuesto` | `B7:O14` | `Categoría` (input, text); `Enero` (input, amount); `Febrero` (input, amount); `Marzo` (input, amount); `Abril` (input, amount); `Mayo` (input, amount); `Junio` (input, amount); `Julio` (input, amount); `Agosto` (input, amount); `Septiembre` (input, amount); `Octubre` (input, amount); `Noviembre` (input, amount); `Diciembre` (input, amount); `Total` (derived, amount). Total fuera: `B15:O15` (`TOTAL DE LA BODA`). |
| `Pagos reales` | `B7:O14` | `Categoría` (derived, text); `Enero` (derived, amount); `Febrero` (derived, amount); `Marzo` (derived, amount); `Abril` (derived, amount); `Mayo` (derived, amount); `Junio` (derived, amount); `Julio` (derived, amount); `Agosto` (derived, amount); `Septiembre` (derived, amount); `Octubre` (derived, amount); `Noviembre` (derived, amount); `Diciembre` (derived, amount); `Total` (derived, amount). Total fuera: `B15:O15` (`TOTAL PAGADO`). |
| `Proveedores` | `B7:I27` | `Fecha` (input, date); `Mes` (input, text); `Tipo` (input, text); `Categoría` (input, text); `Concepto` (input, text); `Importe` (input, amount); `Pagado por` (input, text); `Estado` (input, text). |
| `Flujo de pagos` | `B7:J19` | `Mes` (derived, text); `Presupuesto plan` (derived, amount); `Ingresos` (derived, amount); `Pagos previstos` (derived, amount); `Pagos reales` (derived, amount); `Ajustes` (derived, amount); `Saldo previsto` (derived, amount); `Comentarios` (input, text); `Estado` (derived, text). |
| `Checklist` | `B9:J12` | `Tarea` (input, text); `Área` (input, text); `Responsable` (input, text); `Fecha límite` (input, date); `Estado` (input, text); `Prioridad` (input, text); `Días` (derived, count); `Notas` (input, text); `Progreso` (derived, percent). |
| `Invitados` | `B7:H10` | `Grupo` (input, text); `Invitados previstos` (input, count); `Confirmados` (input, count); `Menú` (input, text); `Mesa` (input, count); `Estado` (input, text); `Notas` (input, text). Total fuera: `B11:H11` (`TOTAL`). |
| `Listas` | `B7:C13` | `Parámetro` (input, text); `Valor` (input, general). |

## Reglas de cálculo

- Proveedores es el diario operativo; Pagos reales y Flujo de pagos se alimentan de él.
- Invitados previstos, Confirmados y Mesa son conteos enteros.
- Los porcentajes solo describen progreso o pagado, nunca número de invitados.
- Las filas TOTAL están fuera de las tablas de datos.

### Fórmulas que deben existir

- `Resumen!E12` debe contener `C12-D12`: Pendiente = plan menos pagado.
- `Pagos reales!C8` debe contener `SUMIFS`: Los pagos reales proceden del diario de proveedores.
- `Flujo de pagos!H8` debe contener `C8-F8+D8-G8`: Saldo previsto combina presupuesto, ingresos, pagos y ajustes.
- `Checklist!J10` debe contener `IF(F10="Confirmado",1,0)`: Progreso refleja el estado de la tarea.
- `Invitados!C11` debe contener `SUM(C8:C10)`: El total de invitados previstos suma grupos, no porcentajes.

## Casos de verificación

Los casos se ejecutan sobre una copia temporal del libro. El archivo original no se modifica.

### Caso 1: guest total follows group counts

Comprueba que el total de invitados suma personas y no interpreta la columna como porcentaje.

**Entrada:** `Invitados!C8 = 40`

**Resultados esperados:**
- `Invitados!C11` equals `130`

## Criterios de fallo

- Invitados previstos, Confirmados o Mesa usan formato porcentaje.
- Los pagos reales se introducen duplicados en vez de salir de Proveedores.
- Pendiente o saldo previsto deja de ser una relación calculada.
- Una fila TOTAL queda dentro del rango filtrable de la tabla.

## Regla de independencia

Este documento se actualiza desde el contrato de diseño declarado antes de generar el XLSX. No se debe redactar leyendo hacia atrás el libro generado ni cambiar el resultado esperado para que coincida con un error observado.
