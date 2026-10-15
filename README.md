# Oficios a Outlook

Página web que convierte los oficios de Cambio de Subtipo (Word + PDF) en correos listos para enviar desde Outlook.

Por cada oficio arma un correo con:

- **Para:** el correo de la alcaldía (el que aparece debajo de "Señores", antes del ASUNTO).
- **CC:** los correos de la SAE (@saesas.gov.co) que aparecen en el oficio.
- **Asunto:** el del oficio + " - MUNICIPIO".
- **Cuerpo:** copiado del Word con su formato (negrillas, cursiva, tabla, imagen, firma), desde "Bogotá D.C." hasta el nombre y cargo de quien firma.
- **Adjunto:** el PDF del oficio.

No se envía a la Oficina de Registro, aunque su correo aparezca en el oficio.

Los archivos se leen en el navegador. No se suben a ningún servidor.

## Cómo empareja el Word con el PDF

1. Si tienen el mismo nombre (solo cambia la extensión).
2. Si no, busca el nombre del municipio en el nombre del PDF. El municipio sale del nombre del Word: `Oficio_base_Cambio_de_Subtipo_CALDAS.docx` → CALDAS.

Lo que no se pueda relacionar (Word sin PDF, PDF sin Word, Word sin correo de alcaldía) aparece primero en "Para revisar a mano". Ahí se puede elegir el PDF o corregir los datos.

## Archivos

```
index.html          Página principal
css/styles.css      Estilos
js/docx-html.js     Lee el Word y lo pasa a HTML con su formato
js/extraer.js       Reglas del oficio: rango del cuerpo, Para, CC, asunto, municipio, emparejar con el PDF
js/eml.js           Arma los correos .eml, el listado .csv y las descargas
js/app.js           Interfaz: carga, fichas, revisión manual y botones
vendor/jszip.min.js Librería JSZip 3.10.1 (MIT) para leer el Word y armar el .zip
```

## Publicar en GitHub Pages

1. Crea un repositorio nuevo en GitHub (por ejemplo `oficios-outlook`).
2. Entra a **Add file → Upload files** y **arrastra** el contenido de esta carpeta (`index.html` y las carpetas `css`, `js` y `vendor`). Si eliges los archivos con el botón en vez de arrastrarlos, GitHub no sube las carpetas y la página no funciona.
   Al terminar, en el repositorio deben verse `index.html`, `README.md` y las carpetas `css`, `js` y `vendor`.
3. En el repositorio entra a **Settings → Pages**.
4. En **Source** elige **Deploy from a branch**, rama **main**, carpeta **/ (root)** y guarda.
5. En uno o dos minutos la página queda en `https://TU-USUARIO.github.io/oficios-outlook/`.

También funciona abriendo `index.html` con doble clic, sin internet.

Si la página muestra un aviso rojo diciendo que no cargó completa, falta alguna de las carpetas `js` o `vendor`.

## Uso

1. Carga todos los Word y PDF juntos.
2. Revisa "Para revisar a mano" y completa lo que falte.
3. Descarga el .zip, descomprímelo y haz doble clic en cada `.eml`. Outlook de escritorio lo abre como correo nuevo. Revisa y dale **Enviar**.
4. Con Outlook en el navegador usa los botones **Copiar** y **Copiar cuerpo con formato**, y adjunta el PDF a mano.
