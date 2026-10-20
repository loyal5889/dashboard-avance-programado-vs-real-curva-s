# Métricas y supuestos

## Avance programado

La medida `Avance Programado Curva S` selecciona actividades activas, no resumen,
no hito y con duración de línea base positiva. Cada actividad aporta según su peso
en la duración total de línea base. El avance de actividad se limita entre 0 y 1.

El cálculo utiliza fechas de línea base y `NETWORKDAYS(..., 11)`: ese código
considera domingo como fin de semana. No reproduce automáticamente todos los
calendarios, turnos, festivos o restricciones personalizados de Microsoft Project.

## Avance real

La ejecución semanal proviene de `DATOS` y se vincula por WBS. El modelo incluye
medidas para cantidad ejecutada acumulada, porcentaje real por actividad y aporte
real. La cantidad total debe ser coherente entre semanas para evitar denominadores
contradictorios.

## Desviación

El modelo incorpora medidas de desviación y estado para comparar avance real y
programado. Una diferencia de porcentajes se interpreta en **puntos porcentuales**,
no como días de retraso. Las comparaciones requieren una fecha de corte común.

## No confundir con valor ganado

La ponderación por duración no equivale a ponderación por costo o presupuesto.
Esta curva de avance no debe llamarse automáticamente EVM, SPI o CPI. Para esas
métricas se necesitan definiciones y fuentes adicionales.

## Fuentes y publicación

Se incluyen fuentes reales de DARCA II y el PBIX guardado al momento de la copia. Las definiciones editables mantienen rutas de instalación `C:\Datos\Demo`; el README explica cómo conectarlas a los archivos reales publicados. No se ha actualizado el modelo en Desktop durante esta publicación. Los generadores sintéticos se conservan solo para las pruebas automatizadas.
