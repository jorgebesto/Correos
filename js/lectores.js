/* lectores.js — Convierte cada archivo (.docx, .pdf, .txt) en texto plano. */
window.Oficios = window.Oficios || {};

(function (O) {
  "use strict";

  async function leerPdf(buffer) {
    var pdf = await pdfjsLib.getDocument({
      data: new Uint8Array(buffer.slice(0)),
      isEvalSupported: false,
      disableFontFace: true
    }).promise;
    var paginas = [];
    for (var p = 1; p <= pdf.numPages; p++) {
      var pagina = await pdf.getPage(p);
      var contenido = await pagina.getTextContent();
      var texto = "";
      contenido.items.forEach(function (it) { texto += it.str + (it.hasEOL ? "\n" : ""); });
      paginas.push(texto);
    }
    return paginas.join("\n\n");
  }

  async function leerDocx(buffer) {
    var r = await mammoth.extractRawText({ arrayBuffer: buffer.slice(0) });
    return r.value;
  }

  /**
   * Lee un archivo y devuelve { texto, bytes, ext }.
   * Lanza un error con mensaje en español si no se puede leer.
   */
  O.leerArchivo = async function (archivo) {
    var ext = (archivo.name.split(".").pop() || "").toLowerCase();
    var buffer = await archivo.arrayBuffer();
    var bytes = new Uint8Array(buffer.slice(0));
    var texto;

    if (ext === "docx") texto = await leerDocx(buffer);
    else if (ext === "pdf") texto = await leerPdf(buffer);
    else if (ext === "txt") texto = new TextDecoder("utf-8").decode(bytes);
    else if (ext === "doc") throw new Error("Los .doc antiguos no se pueden leer. Guárdalo como .docx o PDF.");
    else throw new Error("Formato no admitido. Usa .docx, .pdf o .txt.");

    return { texto: texto, bytes: bytes, ext: ext };
  };
})(window.Oficios);
