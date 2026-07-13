# Reordenar Registros Horarios (Humand)

Página web (sin backend) que toma el Excel de "Registros horarios" tal
como lo descargás de Humand y te devuelve el mismo archivo con las
columnas reordenadas según el formato de "Registros horarios 16-30 Junio".
Mantiene la fuente, los colores, el formato de horas/fechas y los
hipervínculos de geolocalización del archivo original.

Todo el procesamiento pasa en el navegador (no se sube el archivo a
ningún servidor). Usa la librería [ExcelJS](https://github.com/exceljs/exceljs)
porque, a diferencia de otras librerías de Excel para JS, sí preserva
estilos (fuente, relleno, bordes) al escribir el archivo de salida.

## Archivos

- `index.html` — la página.
- `style.css` — estilos.
- `app.js` — lógica: detecta el encabezado, reordena columnas, genera el
  Excel de descarga.
- `config.js` — **el único archivo que necesitás editar** si cambia el
  orden de columnas. Ahí está la lista `TARGET_COLUMNS`.
- `vercel.json` — le dice a Vercel que esto es un sitio estático, sin build.

## Probarlo en tu computadora (opcional)

No hace falta nada instalado. Basta con abrir `index.html` con doble
click, o servirlo con cualquier servidor estático, por ejemplo:

```bash
npx serve .
```

## Subirlo a GitHub

Desde la carpeta del proyecto:

```bash
git init
git add .
git commit -m "Reordenar registros horarios"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/reporte-horarios.git
git push -u origin main
```

(Antes creá el repositorio vacío en GitHub, sin README ni .gitignore,
para que el `git push` no choque con nada.)

## Publicarlo en Vercel

1. Entrá a https://vercel.com y logueate (podés usar tu cuenta de GitHub).
2. "Add New..." → "Project".
3. Elegí el repositorio `reporte-horarios` que acabás de subir.
4. Framework Preset: dejalo en "Other" (es un sitio estático, no necesita build).
5. Deploy.

A los 30 segundos te da una URL tipo `reporte-horarios.vercel.app` que
podés abrir desde cualquier compu o celu.

## Modificarlo después

- **Cambiar el orden o agregar/sacar columnas:** editá el array
  `TARGET_COLUMNS` en `config.js`.
- Cada vez que hagas `git push`, Vercel vuelve a publicar sola la nueva
  versión (deploy automático).

## Qué pasa si falta una columna

Si el Excel que subís no tiene alguna de las columnas de
`TARGET_COLUMNS` (hoy pasa con "Legajo", "Observaciones",
"Horario ingreso programado" y "Horario egreso programado", que no
vienen en el reporte estándar), esa columna sale en blanco en el
archivo de salida. El mensaje de estado en la página te avisa cuáles
fueron.
