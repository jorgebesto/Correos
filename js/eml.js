/* eml.js — Arma los correos .eml (listos para Outlook), el listado .csv y las descargas. */
window.Oficios = window.Oficios || {};

(function (O) {
  "use strict";

  var enc = new TextEncoder();

  function base64(bytes) {
    var s = "", TROZO = 0x8000;
    for (var i = 0; i < bytes.length; i += TROZO) {
      s += String.fromCharCode.apply(null, bytes.subarray(i, i + TROZO));
    }
    return btoa(s);
  }
  function cortar76(s) { return s.replace(/.{1,76}/g, "$&\r\n"); }

  // Codifica encabezados con tildes y ñ.
  function codificar(str) {
    if (/^[\x20-\x7E]*$/.test(str) && str.length < 70) return str;
    var partes = [], actual = "";
    for (var ch of str) {
      if (enc.encode(actual + ch).length > 45) { partes.push(actual); actual = ch; }
      else actual += ch;
    }
    if (actual) partes.push(actual);
    return partes.map(function (p) { return "=?UTF-8?B?" + base64(enc.encode(p)) + "?="; }).join("\r\n ");
  }
  function nombreArchivo(n) {
    return /^[\x20-\x7E]*$/.test(n) ? '"' + n.replace(/"/g, "") + '"' : '"=?UTF-8?B?' + base64(enc.encode(n)) + '?="';
  }
  function direcciones(s) { return (s || "").split(/[;,\s]+/).filter(Boolean).join(", "); }

  /** Documento HTML completo del cuerpo del correo. */
  O.documentoHtml = function (cuerpo) {
    return '<!DOCTYPE html><html><head><meta charset="utf-8">' +
      '<meta name="viewport" content="width=device-width, initial-scale=1"></head>' +
      '<body style="margin:0;background:#ffffff;color:#000000;font-family:\'Arial Narrow\',Arial,sans-serif;font-size:11pt">' +
      cuerpo + "</body></html>";
  };

  /**
   * Arma el .eml:
   *  - "X-Unsent: 1" → Outlook lo abre como correo NUEVO listo para enviar.
   *  - Cuerpo en HTML con las imágenes del oficio incrustadas.
   *  - El PDF como adjunto.
   */
  O.armarEml = function (o) {
    var id = Date.now().toString(36) + "_" + o.id;
    var mixto = "----=_Mixto_" + id, relacionado = "----=_Rel_" + id;

    var enc1 = ["X-Unsent: 1", "To: " + direcciones(o.para)];
    if (direcciones(o.cc)) enc1.push("Cc: " + direcciones(o.cc));
    enc1.push("Subject: " + codificar(o.asunto.replace(/\s+/g, " ").trim()));
    enc1.push("MIME-Version: 1.0");
    enc1.push('Content-Type: multipart/mixed; boundary="' + mixto + '"');

    var partes = [];
    partes.push(
      "--" + mixto + "\r\n" +
      'Content-Type: multipart/related; boundary="' + relacionado + '"; type="text/html"\r\n\r\n' +
      "--" + relacionado + "\r\n" +
      'Content-Type: text/html; charset="UTF-8"\r\n' +
      "Content-Transfer-Encoding: base64\r\n\r\n" +
      cortar76(base64(enc.encode(O.documentoHtml(o.html)))) +
      o.imagenes.map(function (im) {
        return "--" + relacionado + "\r\n" +
          "Content-Type: " + im.tipo + "; name=" + nombreArchivo(im.nombre) + "\r\n" +
          "Content-Transfer-Encoding: base64\r\n" +
          "Content-ID: <" + im.cid + ">\r\n" +
          "Content-Disposition: inline; filename=" + nombreArchivo(im.nombre) + "\r\n\r\n" +
          cortar76(base64(im.bytes));
      }).join("") +
      "--" + relacionado + "--\r\n"
    );
    if (o.pdf) {
      partes.push(
        "--" + mixto + "\r\n" +
        "Content-Type: application/pdf; name=" + nombreArchivo(o.pdf.nombre) + "\r\n" +
        "Content-Transfer-Encoding: base64\r\n" +
        "Content-Disposition: attachment; filename=" + nombreArchivo(o.pdf.nombre) + "\r\n\r\n" +
        cortar76(base64(o.pdf.bytes))
      );
    }
    return enc1.join("\r\n") + "\r\n\r\n" + partes.join("") + "--" + mixto + "--\r\n";
  };

  /** Nombre del .eml: MUNICIPIO (y radicado si lo tiene). */
  O.nombreEml = function (o) {
    var base = (o.radicado ? o.radicado + " - " : "") + (o.municipio || o.nombre.replace(/\.[^.]+$/, ""));
    return base.replace(/[\\/:*?"<>|]+/g, " ").replace(/\s+/g, " ").trim().slice(0, 90) || "oficio";
  };

  /** Listado en .csv (separado por ; para Excel en español). */
  O.armarCsv = function (filas) {
    var q = function (s) { return '"' + String(s || "").replace(/"/g, '""') + '"'; };
    var cab = ["Municipio", "Word", "PDF", "Para", "CC", "Asunto", "Estado"];
    return "﻿" + [cab.join(";")].concat(filas.map(function (f) { return f.map(q).join(";"); })).join("\r\n");
  };

  /** Descarga un archivo desde el navegador. */
  O.descargar = function (nombre, datos, tipo) {
    var blob = datos instanceof Blob ? datos : new Blob([datos], { type: tipo || "application/octet-stream" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = nombre;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 2000);
  };
})(window.Oficios);
