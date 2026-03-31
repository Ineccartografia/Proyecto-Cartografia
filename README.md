# Proyecto-Carto
Hola

## Carlos (INEC)
- Amanzanado en promedio 50 casas (maximo se realizan 60 sin problema) al dia
- Dispersos:: en promedio se visitan 40 casas al dia

Realizar cambios
Agregar el numero de manzanas e dispersos por equipo ()


## Errores a corregir
Al subir el archivo Man_SEC_ENCIET.gpkg y darle a "Procesar", aparece el siguiente mensaje:
"index 1 is out of bounds for axis 0 with size 1"

### Posible causa

- El archivo de la ENCIET guarda en una sola capa tanto los sectores amanzanados como los dispersos, contrario al archivo gpkg de la ENDI.
- El programa app.py está quemado para identificar las dos capas y unirlas en una sola tabla.



### Posible solución:

Que el programa identifique, en los arhcivos con una sola capa, que los sectores dispersos contienen en su código "ManSec" los caracteres 999. Se debe además extraer de estos códigos la provincia, cantón y parroquia, ya que el archivo de la ENCIET no cuenta con esta información para las zonas dispersas.



## CAMBIOS

Cambios realizados en app.py el dia 31/3/26. Agregar estética de Franklin, permitir modificaciones manuales y corregir planificación (necesario, los slots no son independientes)
