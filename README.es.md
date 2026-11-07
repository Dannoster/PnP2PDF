# PnP2PDF

[![tests](https://github.com/Dannoster/PnP2PDF/actions/workflows/tests.yml/badge.svg)](https://github.com/Dannoster/PnP2PDF/actions/workflows/tests.yml)

[English](README.md) · [Русский](README.ru.md) · [简体中文](README.zh-CN.md) · **Español** · [Deutsch](README.de.md)

PnP2PDF monta cartas print-and-play en un PDF listo para imprimir. Elige una
carpeta con las imágenes de las cartas, el papel y la cuadrícula, y obtén un solo
PDF con líneas de corte, zonas de seguridad y, cuando hacen falta, los dorsos
reflejados en la cara posterior.

- una sola ventana, sin línea de comandos;
- vista previa en vivo de las hojas antes de imprimir;
- cartas sin dorso, un dorso compartido, varios dorsos distintos y cartas con su
  propio dorso;
- valores por defecto para A4 y cartas de 62,5 × 88,7 mm (tamaño póker).

## Contenido

- [Aplicación lista](#aplicación-lista)
- [Ejecutar desde el código](#ejecutar-desde-el-código)
- [Inicio rápido](#inicio-rápido)
- [Estructura de carpetas](#estructura-de-carpetas)
- [Ajustes](#ajustes)
- [Ejemplos](#ejemplos)
- [Línea de comandos](#línea-de-comandos)
- [Compilar una aplicación independiente](#compilar-una-aplicación-independiente)
- [Si algo falla](#si-algo-falla)
- [Desarrollo](#desarrollo)
- [Problemas conocidos](#problemas-conocidos)
- [Licencia](#licencia)

## Aplicación lista

Descarga el archivo de tu sistema desde
[Releases](https://github.com/Dannoster/PnP2PDF/releases), descomprímelo y
ejecútalo. No hace falta instalar Python.

### Primera ejecución en macOS

Las versiones publicadas no están firmadas con una identidad de Apple, así que macOS
pide permiso la primera vez. Por orden:

1. Haz doble clic en la aplicación. Si macOS no la abre, abre la carpeta en Finder,
   **haz clic derecho (Control-clic) en la aplicación → Abrir → Abrir**. Después el
   sistema recuerda la elección.
2. Si el diálogo no tiene el botón *Abrir*, abre **Ajustes del Sistema → Privacidad y
   seguridad**, baja hasta la línea de PnP2PDF y pulsa **Abrir de todos modos**.
3. Si macOS dice que la aplicación *está dañada y no se puede abrir*, quita la
   cuarentena una vez y vuelve a abrirla:

   ```bash
   xattr -dr com.apple.quarantine /Applications/PnP2PDF.app
   open /Applications/PnP2PDF.app
   ```

### Primera ejecución en Windows

Un ejecutable sin firmar puede activar Defender SmartScreen: pulsa **Más información →
Ejecutar de todas formas**. El archivo está bien; tras el primer arranque el aviso no
vuelve.

## Ejecutar desde el código

Necesitas Python 3.11 o superior.

```bash
git clone https://github.com/Dannoster/PnP2PDF.git
cd PnP2PDF
pip install -r requirements.txt
python PnP2PDF.py
```

Instalar el paquete añade dos comandos: `pnp2pdf` abre la ventana y `pnp2pdf-cli`
genera el PDF sin ventana.

```bash
pip install .
pnp2pdf
```

## Inicio rápido

1. **Cards folder** — elige la carpeta con las imágenes. La vista previa aparece
   en cuanto se selecciona la carpeta.
2. **Papel y tamaño de carta** — por defecto A4 (210 × 297 mm) y carta de
   62,5 × 88,7 mm. Marca *Smaller size (98%)* para imprimir las cartas algo más
   pequeñas y dejar más papel entre ellas.
3. **Disposición** — filas × columnas, el espacio entre cartas y, si hace falta,
   zonas de seguridad (márgenes que la impresora no alcanza).
4. **Double sided cards** — márcalo cuando la carpeta tenga dorsos. Mira
   [Estructura de carpetas](#estructura-de-carpetas): sin esta casilla las
   imágenes de dorso se imprimen como cartas normales.
5. **Create PDF** — la barra llega al final, la ventana sigue abierta, la línea de
   estado se pone verde con el nombre del archivo y **Show PDF** lo muestra en el
   Finder, el Explorador o el gestor de archivos.

El PDF se guarda en una carpeta con la fecha junto a las cartas. Marca *Save the
PDF to a different folder* para elegir otro destino.

## Estructura de carpetas

La *carpeta de cartas* es la que eliges en la ventana; todo lo siguiente va dentro
de ella. Formatos admitidos: `.png`, `.jpg`, `.jpeg`, `.tif`, `.tiff`.

Las imágenes se estiran al área de la carta, así que usa imágenes con la misma
proporción que las cartas.

### Cartas sin dorso

Basta con poner las imágenes en la carpeta, sin marcar nada:

```
my-cards/
├── card-01.png
├── card-02.png
└── card-03.png
```

### Un dorso compartido por todas las cartas

El dorso es una imagen cuyo nombre termina con su tamaño en milímetros, y sus
cartas van en una subcarpeta con el mismo nombre **sin** el tamaño:

```
my-cards/
├── red-back_63,5_88,9.png     ← el dorso, con su tamaño en el nombre
└── red-back/                  ← cartas que se imprimen con ese dorso
    ├── card-01.png
    └── card-02.png
```

Marca **Double sided cards** y genera el PDF. Cada hoja de caras va seguida de la
misma hoja de dorsos, reflejada para que todo coincida al imprimir a doble cara.

### Varios dorsos

Repite el patrón: un archivo de dorso y una carpeta por dorso. Las cartas de
dorsos distintos pueden tener tamaños distintos y se maquetan juntas.

```
my-cards/
├── red-back_63,5_88,9.png
├── blue-back_64,5_91,3.png
├── red-back/
│   └── card-01.png …
└── blue-back/
    └── card-01.png …
```

El sufijo de tamaño es opcional: sin él el dorso se estira al tamaño de la carta;
con él, un dorso mayor que la carta se recorta a la ventana de la carta, de modo
que el sangrado y las marcas de registro no acaban en la carta impresa. Para la
parte decimal sirve coma o punto: `63,5` y `63.5` son equivalentes.

### Cartas con su propio dorso

Si el dorso pertenece a una sola carta, pon las dos imágenes en una carpeta
llamada `Double Sided`. Los archivos se emparejan por orden alfabético: cara,
dorso, cara, dorso…

```
my-cards/
└── Double Sided/
    ├── 1-front.png
    ├── 2-back.png
    └── 3-front.png …
```

Un número impar de archivos se rechaza con una explicación, porque el emparejado
se desplazaría.

### Cartas sin dorso en un conjunto mixto

Para eso está la carpeta `Single Sided`:

```
my-cards/
├── Double Sided/…
└── Single Sided/
    ├── card-01.png
    └── card-02.png
```

Las cuatro estructuras se pueden combinar en una sola carpeta y generar de una vez.

> **No olvides la casilla.** La estructura de carpetas dice *qué dorso
> corresponde a qué carta*; la casilla **Double sided cards** decide si esos
> dorsos se imprimen. Si está desmarcada, las imágenes de dorso se tratan como
> caras normales y acaban en la hoja junto a las cartas.

## Ajustes

| Ajuste | Por defecto | Notas |
| --- | --- | --- |
| Tamaño del papel | 210 × 297 mm | A4. Cualquier medida en milímetros. |
| Tamaño de la carta | 62,5 × 88,7 mm | Tamaño póker. Las imágenes se estiran a este rectángulo. |
| Smaller size (98%) | desactivado | Imprime cada carta al 98% del tamaño indicado. |
| Filas × columnas | 3 × 3 | Cartas por hoja. |
| Separación | 2 mm | Espacio entre cartas, por donde pasan las líneas de corte. |
| Zona segura arriba/abajo | 0 mm | Márgenes que la impresora no alcanza. |
| Zona segura izquierda/derecha | 0 mm | Lo mismo en horizontal. |
| Carpeta de cartas | — | La carpeta descrita arriba. |
| Guardar el PDF en | junto a las cartas | Opcional, se puede elegir otra carpeta. |
| Double sided cards | desactivado | Imprimir los dorsos. |

Los ajustes se guardan entre ejecuciones en `presets.json`, dentro de la carpeta
de datos de la aplicación (rutas en [Si algo falla](#si-algo-falla)). Al ejecutar
desde el código, el archivo queda junto al proyecto.

## Ejemplos

[`examples/`](examples/) contiene conjuntos de cartas listos para cada estructura
—unos pocos kilobytes— y `make_examples.py` para regenerarlos. Hay más detalles en
[`examples/README.md`](examples/README.md).

| Carpeta | Qué muestra | Páginas |
| --- | --- | --- |
| `single-sided/` | nueve cartas sin dorso | 1 |
| `one-back/` | seis cartas con un dorso compartido | 2 |
| `two-backs/` | dos dorsos, uno mayor que las cartas | 4 |
| `double-sided/` | cartas emparejadas y cartas de una cara en una ejecución | 2 |

```bash
python3 -m pnp2pdf.cli --cards examples/one-back --out /tmp/ex --double-sided --quiet
```

## Línea de comandos

`pnp2pdf-cli` genera los mismos PDF sin abrir ventana, útil para scripts y lotes.

```bash
pnp2pdf-cli --cards ~/cards --out ~/pdf
pnp2pdf-cli --cards ~/cards --out ~/pdf --double-sided
pnp2pdf-cli --cards ~/cards --preview /tmp/first-page.png
```

| Opción | Significado |
| --- | --- |
| `--cards DIR` | Carpeta con las imágenes (obligatorio). |
| `--out DIR` | Dónde se escribe el PDF. Por defecto, junto a las cartas. |
| `--paper WxH` | Papel en mm, por defecto `210x297`. |
| `--card WxH` | Tamaño de carta en mm, por defecto `62.5x88.7`. |
| `--grid RxC` | Filas × columnas, por defecto `3x3`. |
| `--gap MM` | Separación entre cartas en mm, por defecto `2`. |
| `--safe-ud MM` | Zona segura arriba/abajo en mm, por defecto `0`. |
| `--safe-lr MM` | Zona segura izquierda/derecha en mm, por defecto `0`. |
| `--smaller-size` | Imprimir al 98% del tamaño indicado. |
| `--double-sided` | Imprimir los dorsos. |
| `--preview PATH` | Guardar la primera página como PNG y salir. |
| `--quiet` | Sin salida de progreso. |
| `--version` | Mostrar la versión y salir. |

Códigos de salida: `0` éxito, `1` error inesperado, `2` datos incorrectos, `130`
cancelado.

## Compilar una aplicación independiente

Un solo comando compila la aplicación y repara su propio entorno:

```bash
./build_app.sh          # macOS y Linux
build_app.bat           # Windows
```

El script crea `.venv` si falta, instala o actualiza las dependencias, reemplaza un
entorno virtual roto o de otra plataforma, ejecuta las pruebas y empaqueta la
aplicación. Resultado:

- macOS: `dist/PnP2PDF.app`
- Windows: `dist/PnP2PDF.exe`
- Linux: `dist/PnP2PDF`

`--no-tests` omite las pruebas; `PNP2PDF_BUILD_DRY=1` revisa el entorno y se
detiene antes de empaquetar. En macOS, después de escribir la versión, la app se
firma de nuevo y la compilación falla con un mensaje claro si la firma no valida.

## Si algo falla

**La aplicación empaquetada se cierra o no hace nada.** Los errores que una
compilación con ventana oculta se escriben en `crash.log`, y el diálogo de error
indica la ruta. El mismo archivo registra lo que se pulsó; rota a los 320 KB y
conserva una copia anterior en `crash.log.1`. Envíalo al informar de un problema.

- macOS: `~/Library/Application Support/PnP2PDF/crash.log`
- Windows: `%APPDATA%\PnP2PDF\crash.log`
- Linux: `$XDG_DATA_HOME/pnp2pdf/crash.log` (o `~/.local/share/pnp2pdf/crash.log`)

**«PnP2PDF is damaged and can't be opened» (macOS).** La app descargada y sin
firmar queda en cuarentena: mira [Aplicación lista](#aplicación-lista).

**«No card images found».** Las imágenes deben terminar en `.png`, `.jpg`,
`.jpeg`, `.tif` o `.tiff`, y las cartas con dorso van en la subcarpeta con el
nombre de ese dorso.

**Los dorsos salieron en la hoja como cartas.** No está marcada la casilla
**Double sided cards**: mira [Estructura de carpetas](#estructura-de-carpetas).

**La app se abre como una carpeta (macOS).** El paquete o la base de datos de
inicio están desfasados: recompila con `./build_app.sh` y luego ejecuta
`/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister -f dist/PnP2PDF.app`.

## Desarrollo

```bash
pip install -r requirements-dev.txt
pytest
```

Las pruebas sustituyen CustomTkinter, así que corren sin pantalla. Estructura del
paquete:

```
pnp2pdf/
├── app.py          bucle de eventos y arranque
├── cards.py        recorrido de la carpeta y emparejado de cartas y dorsos
├── layout.py       reparto de las cartas en hojas
├── render.py       dibujo de la página: cartas, cortes, esquinas, zonas
├── pdf.py          escritura del PDF en CMYK
├── preview.py      render rápido de baja resolución para la vista previa
├── job.py          una compilación: cartas → hojas → PDF
├── cli.py          entrada de línea de comandos
└── ui/             ventana, ayudas de formulario, hilo de trabajo
```

## Problemas conocidos

- El archivo de macOS es una compilación arm64, porque se genera con el Python que
  ejecuta el script de compilación. En un Mac con Intel, compila la aplicación tú mismo
  con `./build_app.sh`.
- Ninguna de las dos compilaciones está firmada, así que el sistema pide permiso en el
  primer arranque — ver [Aplicación lista](#aplicación-lista).
- Las imágenes se estiran al área de la carta, así que una imagen con otra proporción
  que el tamaño elegido se ve deformada.

## Licencia

[MIT](LICENSE).
