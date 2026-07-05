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

## Publicar tus servicios cada dia (planificador)

Cuando ya tienes la base de grupos, `posting_planner.py` arma un plan diario de
publicacion de tus servicios. Esta pensado para publicar de forma sostenible y
**sin caer en spam**, y con una regla clave de seguridad:

- **No inicia sesion en Facebook.** No guarda tu usuario, contrasena ni cookies.
- **No publica solo.** Genera una pagina local con el enlace de cada grupo y el
  mensaje listo para copiar; tu publicas manualmente en tu navegador donde ya
  estas conectado. Asi tus credenciales nunca quedan expuestas.
- **Ritmo poco agresivo.** Respeta un tope diario de grupos y un enfriamiento
  por grupo (no repite el mismo grupo antes de X dias).
- **Anti texto identico.** Rota entre varias variantes de tu mensaje y las
  personaliza por grupo (`{grupo}`, `{ubicacion}`, `{categoria}`).

### 1. Configuración inicial (una sola vez)

```bash
python3 posting_planner.py init
```

Esto crea dos archivos a partir de las plantillas `*.example.*`:

- `mensajes_servicios.txt`: tus mensajes. Escribe **varias variantes** separadas
  por una linea con `---`. Puedes usar `{grupo}`, `{ubicacion}` y `{categoria}`.
- `posting_config.json`: ajustes del plan.

```json
{
  "mensajes_file": "mensajes_servicios.txt",
  "daily_limit": 6,
  "cooldown_days": 12,
  "min_priority": 0,
  "min_members": 0,
  "categorias": [],
  "spread_categorias": true
}
```

- `daily_limit`: cuantos grupos como maximo publicas por dia.
- `cooldown_days`: dias que deben pasar antes de repetir un mismo grupo.
- `min_priority` / `min_members`: filtros opcionales de calidad.
- `categorias`: si la dejas vacia usa todas; si pones `["Autos"]` filtra.
- `spread_categorias`: reparte el dia entre categorias para no publicar todo del
  mismo rubro.

Estos dos archivos quedan fuera de git (tienen tu contacto). Las plantillas
`mensajes_servicios.example.txt` y `posting_config.example.json` si se versionan.

### 2. Generar el plan del día

```bash
python3 posting_planner.py plan
```

Elige los grupos que tocan hoy (respetando enfriamiento y tope), asigna una
variante a cada uno y genera:

- `plan_hoy.html`: abrela en tu navegador. Cada tarjeta tiene un boton
  **Abrir grupo**, un boton **Copiar mensaje** y una casilla para marcar cuando
  ya publicaste.
- `plan_hoy.json`: registro interno del plan.

Puedes ajustar sobre la marcha:

```bash
python3 posting_planner.py plan --limit 4 --cooldown 15 --categoria "Autos"
```

### 3. Publicar

Abre `plan_hoy.html`, y para cada grupo: **Abrir grupo → Copiar mensaje →**
pega en Facebook, revisa, publica y marca la casilla.

### 4. Registrar lo publicado

Al terminar, avanza la rotacion (marca todos los del plan como publicados hoy):

```bash
python3 posting_planner.py done
```

O solo algunos, si no alcanzaste a publicar en todos:

```bash
python3 posting_planner.py done grp_0001 grp_0007
```

### Ver estado de la rotación

```bash
python3 posting_planner.py status
```

Muestra cuantos grupos hay elegibles hoy y cuales estan en enfriamiento.

### Ejecutar todos los días

El script solo **prepara** el plan; la publicacion siempre la haces tu (esa es
la parte que protege tu cuenta y tus credenciales). Para tenerlo listo cada dia
puedes automatizar solo la generacion del plan:

- **Linux/Mac (cron):**

  ```cron
  0 10 * * * cd /ruta/a/Grupos-face-buscador && python3 posting_planner.py plan
  ```

- **Windows (Programador de tareas):** crea una tarea diaria que ejecute
  `python posting_planner.py plan` en la carpeta del proyecto.

Luego abres `plan_hoy.html`, publicas y corres `done`.

### Buenas prácticas para no caer en spam

- No subas el `daily_limit` a numeros altos ni bajes mucho el `cooldown_days`.
- Manten varias variantes de mensaje reales y utiles, no solo publicidad.
- Publica en horarios distintos y responde comentarios.
- Respeta las reglas de cada grupo; algunos no permiten promociones.

## Archivos generados

- `facebook_groups.json`: base interna editable, fuente principal de datos.
- `facebook_groups.xlsx`: Excel generado para revisar, filtrar y ordenar.
- `plan_hoy.html` / `plan_hoy.json`: plan de publicacion del dia (regenerables).

Si editas el Excel manualmente, esos cambios no vuelven al JSON. Lo ideal es
usar el script para agregar datos y usar el Excel como vista de trabajo.
