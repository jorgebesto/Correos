# Oficios a Outlook

Página web que lee oficios (Word .docx, PDF con texto o .txt) y por cada uno arma un correo listo para Outlook:

- **Para:** los correos que aparecen dentro del oficio.
- **Asunto:** el mismo texto que sigue a "ASUNTO:" en el oficio.
- **Cuerpo:** el texto completo del oficio.
- **Adjunto:** el oficio original (opcional).

Los archivos se leen en el navegador. No se suben a ningún servidor.

## Archivos

```
index.html        Página principal
css/styles.css    Estilos (colores, tipografía, modo oscuro)
js/extraer.js     Saca correos, asunto y radicado del texto
js/lectores.js    Lee .docx, .pdf y .txt
js/eml.js         Arma los correos .eml, el .csv y las descargas
js/app.js         Interfaz: carga, fichas y botones
```

## Publicar en GitHub Pages

1. Crea un repositorio nuevo en GitHub (por ejemplo `oficios-outlook`).
2. Sube todo el contenido de esta carpeta (`index.html` debe quedar en la raíz del repositorio).
3. En el repositorio entra a **Settings → Pages**.
4. En **Source** elige **Deploy from a branch**, rama **main**, carpeta **/ (root)** y guarda.
5. En uno o dos minutos la página queda en `https://TU-USUARIO.github.io/oficios-outlook/`.

También funciona abriendo `index.html` directamente con doble clic, sin subirlo a ningún lado (necesita internet para cargar las librerías).

## Uso

1. Carga los oficios.
2. Revisa cada ficha. Si el oficio trae el correo de tu propia entidad, escribe tu dominio en "No enviar a".
3. Descarga el .zip, descomprímelo y haz doble clic en cada `.eml`. Outlook de escritorio lo abre como correo nuevo. Revisa y dale **Enviar**.
4. Con Outlook en el navegador usa los botones **Copiar** de cada ficha.

## Limitaciones

- Los PDF escaneados (imagen) no tienen texto y no se pueden leer.
- Los `.doc` antiguos no se leen; guárdalos como `.docx`.
