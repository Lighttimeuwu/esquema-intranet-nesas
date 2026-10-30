# Centro de Control — Base Informes

Panel web para consultar los reportes contenidos en un archivo Excel.

## Uso

1. Abre `index.html` en el navegador.
2. Selecciona el Excel que quieres consultar. El panel y el plano de puestos solo se habilitan cuando el archivo se valida correctamente.
3. Usa **Actualizar Excel** para cargar otro archivo; las tablas y métricas se recalculan con esa nueva carga.

El libro debe incluir las hojas `ASIGNACION RANKING`, `COMISIONES`, `INFORME H A H`, `DATOS ASESORES`, `PUESTOS`, `DOMINGO 1`, `DOMINGO 2` y `FESTIVO 1`. Los datos se leen en memoria en el navegador y no se guardan en caché ni en almacenamiento local.
