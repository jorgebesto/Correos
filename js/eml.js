/* eml.js — Arma los correos .eml (listos para Outlook) y el listado .csv. */
window.Oficios = window.Oficios || {};

(function (O) {
  "use strict";

  var enc = new TextEncoder();
  var TIPOS = {
    pdf: "application/pdf",
    docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    txt: "text/plain"
  };

  function base64(bytes) {
    var s = "", TROZO = 0x8000;
    for (var i = 0; i < bytes.length; i += TROZO) {
      s += String.fromCharCode.apply(null, bytes.subarray(i, i + TROZO));
    }
    return btoa(s);
  }

  function cortar76(s) { return s.replace(/.{1,76}/g, "$&\r\n"); }

  // Codifica el asunto para que tildes y ñ lleguen bien.
  function codificarEncabezado(str) {
    if (/^[\x20-\x7E]*$/.test(str) && str.length < 70) return str;
    var partes = [], actual = "";
    for (var ch of str) {
      if (enc.encode(actual + ch).length > 45) { partes.push(actual); actual = ch; }
      else actual += ch;
    }
    if (actual) partes.push(actual);
    return partes.map(function (p) { return "=?UTF-8?B?" + base64(enc.encode(p)) + "?="; }).join("\r\n ");
  }

  function nombreAdjunto(nombre) {
    return /^[\x20-\x7E]*$/.test(nombre)
      ? '"' + nombre.replace(/"/g, "") + '"'
      : '"=?UTF-8?B?' + base64(enc.encode(nombre)) + '?="';
  }

  /**
   * Arma el texto de un archivo .eml.
   * "X-Unsent: 1" hace que Outlook lo abra como correo NUEVO listo para enviar.
   */
  O.armarEml = function (oficio, conAdjunto) {
    var para = oficio.para.split(/[;,\s]+/).filter(Boolean).join(", ");
    var limite = "----=_Oficio_" + Date.now().toString(36) + "_" + oficio.id;
    var cuerpo = oficio.cuerpo.replace(/\r?\n/g, "\r\n");

    var encabezados = [
      "X-Unsent: 1",
      "To: " + para,
      "Subject: " + codificarEncabezado(oficio.asunto.replace(/\s+/g, " ").trim()),
      "MIME-Version: 1.0"
    ];
    var parteTexto =
      'Content-Type: text/plain; charset="UTF-8"\r\n' +
      "Content-Transfer-Encoding: base64\r\n\r\n" +
      cortar76(base64(enc.encode(cuerpo)));

    if (!conAdjunto || !oficio.bytes) return encabezados.join("\r\n") + "\r\n" + parteTexto;

    encabezados.push('Content-Type: multipart/mixed; boundary="' + limite + '"');
    return encabezados.join("\r\n") + "\r\n\r\n" +
      "--" + limite + "\r\n" + parteTexto + "\r\n" +
      "--" + limite + "\r\n" +
      "Content-Type: " + (TIPOS[oficio.ext] || "application/octet-stream") + "; name=" + nombreAdjunto(oficio.nombre) + "\r\n" +
      "Content-Transfer-Encoding: base64\r\n" +
      "Content-Disposition: attachment; filename=" + nombreAdjunto(oficio.nombre) + "\r\n\r\n" +
      cortar76(base64(oficio.bytes)) + "\r\n" +
      "--" + limite + "--\r\n";
  };

  /** Nombre de archivo seguro para Windows. */
  O.nombreSeguro = function (s) {
    return s.replace(/\.[^.]+$/, "")
            .replace(/[\\/:*?"<>|]+/g, " ")
            .replace(/\s+/g, " ")
            .trim()
            .slice(0, 90) || "oficio";
  };

  /** Nombre del .eml: radicado + nombre del archivo. */
  O.nombreEml = function (oficio) {
    return (oficio.radicado ? oficio.radicado + " - " : "") + O.nombreSeguro(oficio.nombre);
  };

  /** Listado en .csv (separado por ; para Excel en español). */
  O.armarCsv = function (oficios, estadoDe) {
    var q = function (s) { return '"' + String(s || "").replace(/"/g, '""') + '"'; };
    var filas = [["Radicado", "Archivo", "Para", "Asunto", "Estado"].join(";")];
    oficios.forEach(function (o) {
      filas.push([q(o.radicado), q(o.nombre), q(o.para), q(o.asunto), q(estadoDe(o).texto)].join(";"));
    });
    return "﻿" + filas.join("\r\n");
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
