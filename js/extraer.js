/* extraer.js — Saca del texto del oficio: correos, asunto y número de radicado. */
window.Oficios = window.Oficios || {};

(function (O) {
  "use strict";

  var EMAIL_RE = /[A-Za-z0-9._%+\-]+@[A-Za-z0-9\-]+(?:\.[A-Za-z0-9\-]+)*\.[A-Za-z]{2,}/g;

  // Líneas con las que suele empezar el saludo; ahí termina el asunto.
  var SALUDO = /^\s*(cordial|respetad|apreciad|estimad|señor|senor|señores|senores|doctor|doctora|ingenier|de manera|por medio|reciba|con el fin)/i;

  // Etiquetas de asunto, en orden de prioridad.
  var ETIQUETAS = [
    /^\s*asunto\s*[:.\-–]\s*(.*)$/i,
    /^\s*ref(?:erencia)?\s*\.?\s*[:.\-–]\s*(.*)$/i
  ];

  /** Todos los correos del texto, sin repetir y en minúscula. */
  O.buscarCorreos = function (texto) {
    var vistos = {}, salida = [];
    (texto.match(EMAIL_RE) || []).forEach(function (c) {
      c = c.replace(/[.\-]+$/, "").toLowerCase();
      if (!vistos[c]) { vistos[c] = true; salida.push(c); }
    });
    return salida;
  };

  /** El asunto: lo que sigue a "ASUNTO:" (puede ocupar varias líneas). */
  O.buscarAsunto = function (texto) {
    var lineas = texto.split(/\r?\n/);
    for (var k = 0; k < ETIQUETAS.length; k++) {
      for (var i = 0; i < lineas.length; i++) {
        var m = lineas[i].match(ETIQUETAS[k]);
        if (!m) continue;
        var partes = [m[1].trim()], j = i + 1;
        if (!partes[0]) { while (j < lineas.length && !lineas[j].trim()) j++; }
        while (j < lineas.length && lineas[j].trim() && !SALUDO.test(lineas[j]) && partes.join(" ").length < 450) {
          partes.push(lineas[j].trim());
          j++;
        }
        var asunto = partes.join(" ").replace(/\s+/g, " ").trim();
        if (asunto) return asunto;
      }
    }
    return "";
  };

  /** Número de radicado de 14 dígitos que empieza por 20 (ej. 20264100242441). */
  O.buscarRadicado = function (texto) {
    var m = texto.match(/\b20\d{12}\b/);
    return m ? m[0] : "";
  };

  /** Limpia espacios y saltos de línea sobrantes. */
  O.limpiarTexto = function (t) {
    return t.replace(/\r\n?/g, "\n")
            .replace(/[ \t ]+\n/g, "\n")
            .replace(/\n{3,}/g, "\n\n")
            .trim();
  };

  /** ¿El correo coincide con algo de la lista "No enviar a"? */
  O.estaExcluido = function (correo, listaExcluir) {
    var partes = (listaExcluir || "").toLowerCase().split(/[\s,;]+/).filter(Boolean);
    return partes.some(function (p) { return correo.indexOf(p) !== -1; });
  };
})(window.Oficios);
