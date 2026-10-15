/* extraer.js — Reglas del oficio: qué parte va al correo, a quién va, asunto y municipio.
   Pensado para la plantilla de "Cambio de Subtipo" (siempre el mismo formato). */
window.Oficios = window.Oficios || {};

(function (O) {
  "use strict";

  var EMAIL_RE = /[A-Za-z0-9._%+\-]+@[A-Za-z0-9\-]+(?:\.[A-Za-z0-9\-]+)*\.[A-Za-z]{2,}/g;
  var DOMINIO_SAE = "@saesas.gov.co";
  // Correos de la SAE que NO van en copia, aunque aparezcan en el oficio.
  var CC_QUITAR = ["correspondenciaentrada@saesas.gov.co"];
  // Correos que SIEMPRE van en copia.
  var CC_AGREGAR = ["apatino@saesas.gov.co"];

  /** Mayúsculas, sin tildes y sin signos: "Bogotá D.C." → "BOGOTA D C" */
  O.normalizar = function (s) {
    return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "")
      .toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
  };

  O.buscarCorreos = function (texto) {
    var vistos = {}, salida = [];
    (texto.match(EMAIL_RE) || []).forEach(function (c) {
      c = c.replace(/[.\-]+$/, "").toLowerCase();
      if (!vistos[c]) { vistos[c] = true; salida.push(c); }
    });
    return salida;
  };

  O.buscarRadicado = function (texto) {
    var m = texto.match(/\b20\d{12}\b/);
    return m ? m[0] : "";
  };

  function indice(bloques, desde, hasta, prueba) {
    for (var i = desde; i <= hasta; i++) if (prueba(O.normalizar(bloques[i].texto))) return i;
    return -1;
  }

  /**
   * Municipio a partir del nombre del archivo:
   * "Oficio_base_Cambio_de_Subtipo_CALDAS.docx" → "CALDAS"
   * "..._SAN_PELAYO.docx" → "SAN PELAYO"
   */
  O.municipioDeArchivo = function (nombre) {
    var base = nombre.replace(/\.[^.]+$/, "");
    var m = base.match(/subtipo[\s_\-]+(.+)$/i);
    var r = m ? m[1] : "";
    if (!r) {
      var partes = base.split(/[_]+/);
      r = partes.length > 1 ? partes[partes.length - 1] : "";
    }
    return r.replace(/[_\-]+/g, " ").replace(/\s+/g, " ").trim().toUpperCase();
  };

  // Respaldo: "ALCALDÍA DE CALDAS" / "ALCALDÍA MUNICIPAL DE CALDAS"
  function municipioDeTexto(texto) {
    var m = texto.match(/ALCALD[IÍ]A\s+(?:MUNICIPAL\s+)?DE\s+([^\n]+)/i);
    return m ? m[1].replace(/\s+/g, " ").trim().toUpperCase() : "";
  }

  /**
   * Analiza un oficio ya leído con O.leerDocx.
   * El correo va desde la línea "Bogotá D.C." hasta la firma que sigue a "Cordialmente"
   * (nombre y cargo).
   */
  O.analizarOficio = function (docx, nombreArchivo) {
    var b = docx.bloques, ultimo = b.length - 1;
    var r = { municipio: O.municipioDeArchivo(nombreArchivo), para: "", cc: "", asuntoBase: "", radicado: "", html: "", imagenes: [], texto: "", problema: "" };

    var ini = indice(b, 0, ultimo, function (t) { return t.indexOf("BOGOTA") === 0; });
    if (ini < 0) { r.problema = 'No encontré la línea que empieza con "Bogotá D.C."'; return r; }

    var cord = indice(b, ini, ultimo, function (t) { return t.indexOf("CORDIALMENTE") === 0; });
    var fin = cord < 0 ? -1 : indice(b, cord + 1, ultimo, function (t) { return t.length > 0; });
    if (fin < 0) { r.problema = 'No encontré la firma después de "Cordialmente"'; return r; }

    var textos = b.slice(ini, fin + 1).map(function (x) { return x.texto; });
    r.texto = textos.join("").replace(/\n{3,}/g, "\n\n").trim();
    r.radicado = O.buscarRadicado(b.map(function (x) { return x.texto; }).join(" "));
    if (!r.municipio) r.municipio = municipioDeTexto(r.texto);

    // Para: correos entre "Señores" y "ASUNTO" (la alcaldía), sin los de la SAE.
    var sen = indice(b, ini, fin, function (t) { return t.indexOf("SENOR") === 0; });
    var asu = indice(b, ini, fin, function (t) { return t.indexOf("ASUNTO") === 0; });
    if (sen >= 0 && asu > sen) {
      r.para = O.buscarCorreos(b.slice(sen, asu).map(function (x) { return x.texto; }).join("\n"))
        .filter(function (c) { return c.indexOf(DOMINIO_SAE) === -1; }).join("; ");
    }

    // CC: los correos de la SAE del oficio (sin los de CC_QUITAR) + los de CC_AGREGAR.
    var cc = O.buscarCorreos(r.texto).filter(function (c) {
      return c.indexOf(DOMINIO_SAE) !== -1 && CC_QUITAR.indexOf(c) === -1;
    });
    CC_AGREGAR.forEach(function (c) { if (cc.indexOf(c) === -1) cc.push(c); });
    r.cc = cc.join("; ");

    // Asunto: lo que sigue a "ASUNTO:".
    if (asu >= 0) {
      r.asuntoBase = b[asu].texto.replace(/^\s*asunto\s*[:.\-–]?\s*/i, "").replace(/\s+/g, " ").trim();
    }

    var conv = docx.convertir(ini, fin);
    r.html = conv.html;
    r.imagenes = conv.imagenes;
    return r;
  };

  /** Asunto: exactamente el del oficio (sin agregar el municipio). */
  O.armarAsunto = function (base) {
    return base;
  };

  /**
   * Empareja cada Word con su PDF.
   * 1) mismo nombre de archivo; 2) el nombre del PDF contiene el municipio.
   * Si hay dos PDF posibles para el mismo municipio, no adivina: queda para revisar.
   */
  O.emparejar = function (oficios, pdfs) {
    var libres = function () { return pdfs.filter(function (p) { return !p.oficioId; }); };
    var sinExt = function (n) { return O.normalizar(n.replace(/\.[^.]+$/, "")); };

    oficios.forEach(function (o) {
      if (o.pdfId) return;
      var igual = libres().filter(function (p) { return sinExt(p.nombre) === sinExt(o.nombre); })[0];
      if (igual) { o.pdfId = igual.id; igual.oficioId = o.id; }
    });

    oficios.forEach(function (o) {
      if (o.pdfId || !o.municipio) return;
      var mun = " " + O.normalizar(o.municipio) + " ";
      var candidatos = libres().filter(function (p) { return (" " + sinExt(p.nombre) + " ").indexOf(mun) !== -1; });
      if (!candidatos.length) return;
      var extra = function (p) { return sinExt(p.nombre).split(" ").length; };
      candidatos.sort(function (a, c) { return extra(a) - extra(c); });
      if (candidatos.length > 1 && extra(candidatos[0]) === extra(candidatos[1])) return;
      // Evita que "CALDAS" tome el PDF de un municipio que solo lo contiene (ej. otro con más palabras).
      var otro = oficios.some(function (x) {
        return x !== o && x.municipio && x.municipio.length > o.municipio.length &&
          (" " + sinExt(candidatos[0].nombre) + " ").indexOf(" " + O.normalizar(x.municipio) + " ") !== -1;
      });
      if (otro) return;
      o.pdfId = candidatos[0].id;
      candidatos[0].oficioId = o.id;
    });
  };
})(window.Oficios);
