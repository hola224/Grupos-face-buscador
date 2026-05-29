# Vendedor Automatico - Base de grupos de Facebook

Este workspace contiene un script de consola para armar una base local de
grupos de Facebook sin automatizar acciones dentro de Facebook.

La forma recomendada de uso es:

1. Buscas grupos manualmente en Facebook.
2. Copias el nombre, URL si la tienes y cantidad de miembros visible.
3. El script guarda la informacion en `facebook_groups.json`.
4. El script exporta un Excel en `facebook_groups.xlsx`.
5. Si el grupo ya existe, no lo duplica: lo actualiza si hay datos mejores.
6. Si detecta menos de 100 miembros/seguidores, omite el grupo.

## Agregar un grupo manual

```bash
python3 facebook_group_crm.py add \
  --name "Compra Venta Curico" \
  --url "https://www.facebook.com/groups/123456789" \
  --members "56 mil miembros" \
  --query "compra venta Curico" \
  --location "Curico"
```

## Importar varios grupos pegando texto

Puedes usar lineas separadas por `|`:

```text
Compra Venta Curico | https://www.facebook.com/groups/123 | 56 mil miembros
Autos Usados Maule | https://www.facebook.com/groups/456 | 22.500 miembros
Parcelas VII Region | https://www.facebook.com/groups/789 | 18 mil miembros
```

Luego importas:

```bash
python3 facebook_group_crm.py import-text \
  --query "compra venta Curico" \
  --location "Curico" \
  --file resultados.txt
```

Tambien puedes pegar directo en la consola:

```bash
python3 facebook_group_crm.py import-text --query "autos usados Maule"
```

Pegas el contenido y terminas con `Ctrl+D`.

## Capturar mientras buscas

Este es el flujo mas comodo si vas a estar haciendo varias busquedas seguidas.

```bash
python3 facebook_group_crm.py capture \
  --query "compra venta Curico" \
  --location "Curico"
```

Dentro del modo captura pegas resultados en este formato:

```text
Compra Venta Curico | https://www.facebook.com/groups/123 | 56 mil miembros
Autos Usados Maule | https://www.facebook.com/groups/456 | 22.500 miembros
```

Luego escribes:

```text
:save
```

El script guarda lo nuevo, evita duplicados y actualiza `facebook_groups.xlsx`.
Por defecto omite grupos con menos de 100 miembros/seguidores detectados.

Para cambiar de busqueda sin salir:

```text
:query autos usados Maule
:location Maule
```

Para salir:

```text
:q
```

## Capturar desde la consola del navegador

Para reducir el copiado manual, puedes usar `browser_console_capture.js`.

1. Abre Facebook y haz una busqueda de grupos.
2. Abre DevTools y entra a `Console`.
3. Pega el contenido de `browser_console_capture.js`.
4. Presiona Enter.
5. El script copia al portapapeles los grupos visibles en formato compatible.
6. Pegas el resultado dentro del modo `capture` y usas `:save`.

El helper no hace clicks, no publica y no navega. Solo lee datos visibles en la
pantalla actual.

## Captura con un clic desde Facebook

Para no copiar datos manualmente, usa el servidor local:

```bash
python3 capture_server.py \
  --location "Curico" \
  --category "Compra/Venta general"
```

Luego abre la URL que muestra la consola, normalmente:

```text
http://127.0.0.1:8765/
```

En esa pagina arrastra el boton `Capturar grupos FB` a la barra de marcadores.
Despues:

1. Abre Facebook y busca grupos.
2. Haz clic en el marcador `Capturar grupos FB`.
3. El marcador abre una ventana local, baja lentamente por los resultados y
   acumula los grupos detectados.
4. Al terminar, envia los datos por `POST` al servidor local.
5. `facebook_groups.xlsx` queda actualizado.

El Excel tambien guarda `Publicaciones/día` cuando Facebook muestra ese dato en
los resultados.

Puedes controlar cuanto baja:

```bash
python3 capture_server.py \
  --location "Curico" \
  --scrolls 20 \
  --delay-ms 1400
```

`--scrolls` es la cantidad maxima de bajadas. `--delay-ms` es la pausa entre
bajadas para darle tiempo a Facebook a cargar nuevos resultados.

Si pasas `--query`, esa consulta queda fija:

```bash
python3 capture_server.py \
  --query "compra venta Curico" \
  --location "Curico"
```

Si no pasas `--query`, el bookmarklet intenta tomar la busqueda desde la URL de
Facebook.

Si el navegador bloquea la ventana emergente, permite popups para Facebook o
vuelve a hacer clic en el marcador. La ventana local es la que permite enviar
muchos datos sin depender del portapapeles ni del largo de una URL.

## Cambiar el minimo de miembros

La regla actual es omitir grupos con menos de 100 miembros/seguidores.
Si quieres usar otro minimo para una captura o importacion:

```bash
python3 facebook_group_crm.py capture \
  --query "compra venta Curico" \
  --location "Curico" \
  --min-members 500
```

## Ver ranking por miembros

```bash
python3 facebook_group_crm.py list --sort members --limit 50
```

## Ver ranking por prioridad

```bash
python3 facebook_group_crm.py list --sort priority
```

La prioridad inicial se calcula segun la cantidad de miembros:

- 100.000 o mas: prioridad 5
- 50.000 a 99.999: prioridad 4
- 20.000 a 49.999: prioridad 3
- 5.000 a 19.999: prioridad 2
- Menos de 5.000: prioridad 1

## Marcar que ya publicaste

Puedes usar el ID que sale en `list`:

```bash
python3 facebook_group_crm.py mark-posted grp_0001
```

O una fecha especifica:

```bash
python3 facebook_group_crm.py mark-posted grp_0001 --date 2026-05-29
```

## Resumen operativo

```bash
python3 facebook_group_crm.py summary
```

## Limpiar nombres ya guardados

Si una captura guardo nombres con ruido como `Foto del perfil de ...`:

```bash
python3 facebook_group_crm.py clean-names
```

Esto actualiza `facebook_groups.json` y regenera `facebook_groups.xlsx`.

## Archivos generados

- `facebook_groups.json`: base interna editable, fuente principal de datos.
- `facebook_groups.xlsx`: Excel generado para revisar, filtrar y ordenar.

Si editas el Excel manualmente, esos cambios no vuelven al JSON. Lo ideal es
usar el script para agregar datos y usar el Excel como vista de trabajo.
