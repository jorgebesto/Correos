/* app.js — Interfaz: carga de Word y PDF, fichas, revisión manual y descargas. */
(function (O) {
  "use strict";

  var $ = function (id) { return document.getElementById(id); };
  var oficios = [], pdfs = [], contador = 0;

  /* ---------- Carga de archivos ---------- */
  async function agregarArchivos(lista) {
    var archivos = [].slice.call(lista);
    if (!archivos.length) return;
    var ignorados = [], repetidos = [];
    mensaje("Leyendo " + archivos.length + " archivo" + (archivos.length > 1 ? "s" : "") + "…", "ok");

    for (var i = 0; i < archivos.length; i++) {
      var f = archivos[i], ext = (f.name.split(".").pop() || "").toLowerCase();
      if (oficios.concat(pdfs).some(function (x) { return x.nombre === f.name; })) { repetidos.push(f.name); continue; }

      if (ext === "pdf") {
        pdfs.push({ id: ++contador, nombre: f.name, bytes: new Uint8Array(await f.arrayBuffer()), oficioId: null });
      } else if (ext === "docx") {
        oficios.push(await leerOficio(f));
      } else {
        ignorados.push(f.name);
      }
    }

    O.emparejar(oficios, pdfs);
    pintar();

    var avisos = [];
    if (ignorados.length) avisos.push("No se cargaron (solo .docx y .pdf): " + ignorados.join(", ") + ".");
    if (repetidos.length) avisos.push("Ya estaban cargados: " + repetidos.join(", ") + ".");
    mensaje(avisos.join(" "), "err");
  }

  async function leerOficio(f) {
    var o = { id: ++contador, nombre: f.name, pdfId: null, municipio: O.municipioDeArchivo(f.name),
              para: "", cc: "", asunto: "", asuntoBase: "", asuntoEditado: false, radicado: "",
              html: "", imagenes: [], texto: "", problema: "", vistaPrevia: null };
    try {
      var docx = await O.leerDocx(await f.arrayBuffer());
      var r = O.analizarOficio(docx, f.name);
      Object.assign(o, r);
      o.asunto = O.armarAsunto(r.asuntoBase, r.municipio);
    } catch (e) {
      o.problema = "No se pudo leer el Word" + (e && e.message ? ": " + e.message : ".");
    }
    return o;
  }

  /* ---------- Estado ---------- */
  function pdfDe(o) { return pdfs.filter(function (p) { return p.id === o.pdfId; })[0] || null; }
  function estado(o) {
    if (o.problema) return { clase: "err", texto: "No se pudo leer" };
    if (!o.pdfId) return { clase: "warn", texto: "Falta el PDF" };
    if (!o.para.trim()) return { clase: "warn", texto: "Falta correo de la alcaldía" };
    if (!o.municipio.trim()) return { clase: "warn", texto: "Falta municipio" };
    if (!o.asunto.trim()) return { clase: "warn", texto: "Falta asunto" };
    return { clase: "ok", texto: "Listo" };
  }
  function listos() { return oficios.filter(function (o) { return estado(o).clase === "ok"; }); }
  function buscar(id) { id = +id; return oficios.filter(function (o) { return o.id === id; })[0]; }
  function pdfsSinWord() { return pdfs.filter(function (p) { return !p.oficioId; }); }

  /* ---------- Pintar ---------- */
  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c];
    });
  }

  function campo(o, nombre, etiqueta, valor, copiable, mono) {
    return '<div class="row"><div class="rowhead">' +
      '<label for="' + nombre + "-" + o.id + '">' + etiqueta + "</label>" +
      (copiable ? '<button type="button" data-act="copiar" data-campo="' + nombre + '" data-id="' + o.id + '">Copiar</button>' : "") +
      '</div><input type="text"' + (mono ? ' class="mono"' : "") + ' id="' + nombre + "-" + o.id + '" data-campo="' + nombre + '" data-id="' + o.id + '" value="' + esc(valor) + '"></div>';
  }

  function selectorPdf(o) {
    var opciones = pdfs.filter(function (p) { return !p.oficioId || p.oficioId === o.id; });
    return '<div class="row"><div class="rowhead"><label for="pdf-' + o.id + '">PDF que se adjunta</label></div>' +
      '<select id="pdf-' + o.id + '" data-act="pdf" data-id="' + o.id + '">' +
      '<option value="">— Sin PDF —</option>' +
      opciones.map(function (p) {
        return '<option value="' + p.id + '"' + (p.id === o.pdfId ? " selected" : "") + ">" + esc(p.nombre) + "</option>";
      }).join("") +
      "</select></div>";
  }

  function ficha(o) {
    var st = estado(o);
    var etiquetas = (o.radicado ? '<span class="rad" title="Radicado">' + esc(o.radicado) + "</span>" : "") +
      '<span class="pill ' + st.clase + '" data-pill="' + o.id + '">' + st.texto + "</span>";
    var cabeza = '<div class="head"><div class="titulo"><div class="name">' + esc(o.municipio || "Sin municipio") + "</div>" +
      '<div class="archivo mono">' + esc(o.nombre) + "</div></div>" +
      '<div class="tags">' + etiquetas + "</div></div>";

    if (o.problema) {
      return '<article class="card">' + cabeza + '<p class="msg err">' + esc(o.problema) + "</p>" +
        '<div class="foot"><span></span><div class="btns"><button type="button" data-act="quitar" data-id="' + o.id + '">Quitar</button></div></div></article>';
    }

    return '<article class="card" id="ficha-' + o.id + '">' + cabeza +
      '<div class="grid2">' + campo(o, "municipio", "Municipio", o.municipio, false, false) + selectorPdf(o) + "</div>" +
      campo(o, "para", "Para (alcaldía)", o.para, true, true) +
      campo(o, "cc", "CC (SAE)", o.cc, true, true) +
      campo(o, "asunto", "Asunto", o.asunto, true, false) +
      '<details class="cuerpo" data-id="' + o.id + '"><summary>Ver el cuerpo del correo</summary>' +
      '<div class="cuerpo-acciones"><span class="note">Desde "Bogotá D.C." hasta la firma, con el formato del Word.</span>' +
      '<button type="button" data-act="copiarCuerpo" data-id="' + o.id + '">Copiar cuerpo con formato</button></div>' +
      '<div class="marco" id="marco-' + o.id + '"></div></details>' +
      '<div class="foot"><span class="note">' + (pdfDe(o) ? "Adjunto: " + esc(pdfDe(o).nombre) : "Sin PDF adjunto") + "</span>" +
      '<div class="btns"><button type="button" data-act="eml" data-id="' + o.id + '">Descargar este correo (.eml)</button>' +
      '<button type="button" data-act="quitar" data-id="' + o.id + '">Quitar</button></div></div>' +
      "</article>";
  }

  function pintar() {
    var manual = oficios.filter(function (o) { return estado(o).clase !== "ok"; });
    var ok = listos(), sueltos = pdfsSinWord();

    $("vacio").hidden = !!(oficios.length || pdfs.length);
    $("seccionManual").hidden = !(manual.length || sueltos.length);
    $("seccionListos").hidden = !ok.length;

    $("listaManual").innerHTML = manual.map(ficha).join("");
    $("pdfsSueltos").hidden = !sueltos.length;
    $("listaPdfs").innerHTML = sueltos.map(function (p) {
      return '<li><span class="mono">' + esc(p.nombre) + '</span><button type="button" data-act="quitarPdf" data-id="' + p.id + '">Quitar</button></li>';
    }).join("");
    $("listaListos").innerHTML = ok.map(ficha).join("");
    actualizarContador();
  }

  function actualizarContador() {
    var n = listos().length, m = oficios.length - n, s = pdfsSinWord().length;
    var partes = [];
    if (oficios.length || pdfs.length) {
      partes.push("<b>" + n + "</b> listo" + (n === 1 ? "" : "s"));
      if (m) partes.push("<b>" + m + "</b> Word por revisar");
      if (s) partes.push("<b>" + s + "</b> PDF sin Word");
    }
    $("count").innerHTML = partes.length ? partes.join(" · ") : "Sin archivos cargados";
    $("btnZip").disabled = !n;
    $("btnCsv").disabled = !oficios.length;
    $("tituloManual").textContent = "Para revisar a mano (" + (m + s) + ")";
    $("tituloListos").textContent = "Listos para enviar (" + n + ")";
  }

  function mensaje(texto, tipo) {
    var m = $("msg");
    m.hidden = !texto;
    m.textContent = texto || "";
    m.className = "msg " + (tipo || "ok");
  }

  /* ---------- Vista previa del cuerpo ---------- */
  function htmlVista(o) {
    if (!o.vistaPrevia) o.vistaPrevia = O.htmlConImagenes(o.html, o.imagenes);
    return o.vistaPrevia;
  }
  function mostrarVista(o) {
    var marco = $("marco-" + o.id);
    if (!marco || marco.firstChild) return;
    var ifr = document.createElement("iframe");
    ifr.title = "Cuerpo del correo de " + (o.municipio || o.nombre);
    ifr.setAttribute("sandbox", "allow-same-origin");
    ifr.addEventListener("load", function () {
      try { ifr.style.height = (ifr.contentDocument.documentElement.scrollHeight + 8) + "px"; } catch (e) {}
    });
    ifr.srcdoc = O.documentoHtml('<div style="padding:16px">' + htmlVista(o) + "</div>");
    marco.appendChild(ifr);
  }

  /* ---------- Copiar ---------- */
  function avisoBoton(b, texto) {
    var original = b.dataset.original || b.textContent;
    b.dataset.original = original;
    b.textContent = texto;
    setTimeout(function () { b.textContent = original; }, 1500);
  }
  function copiarTexto(b, input) {
    var manual = function () { input.select(); avisoBoton(b, "Usa Ctrl+C"); };
    try { navigator.clipboard.writeText(input.value).then(function () { avisoBoton(b, "Copiado"); }, manual); }
    catch (e) { manual(); }
  }
  function copiarCuerpo(b, o) {
    var html = O.documentoHtml(htmlVista(o));
    try {
      var item = new ClipboardItem({
        "text/html": new Blob([html], { type: "text/html" }),
        "text/plain": new Blob([o.texto], { type: "text/plain" })
      });
      navigator.clipboard.write([item]).then(function () { avisoBoton(b, "Copiado"); }, function () {
        mensaje("El navegador no dejó copiar con formato. Descarga el .eml de este oficio.", "err");
      });
    } catch (e) {
      mensaje("Este navegador no permite copiar con formato. Descarga el .eml de este oficio.", "err");
    }
  }

  /* ---------- Eventos de las fichas ---------- */
  function datosEml(o) { return Object.assign({}, o, { pdf: pdfDe(o) }); }

  document.addEventListener("input", function (e) {
    var t = e.target, o = buscar(t.dataset.id);
    if (!o || !t.dataset.campo) return;
    o[t.dataset.campo] = t.value;
    if (t.dataset.campo === "asunto") o.asuntoEditado = true;
    if (t.dataset.campo === "municipio" && !o.asuntoEditado) {
      o.asunto = O.armarAsunto(o.asuntoBase, t.value.trim().toUpperCase());
      var a = $("asunto-" + o.id);
      if (a) a.value = o.asunto;
    }
    var st = estado(o), pill = document.querySelector('[data-pill="' + o.id + '"]');
    if (pill) { pill.className = "pill " + st.clase; pill.textContent = st.texto; }
    actualizarContador();
  });

  document.addEventListener("change", function (e) {
    var t = e.target;
    if (t.dataset.act === "pdf") {
      var o = buscar(t.dataset.id);
      pdfs.forEach(function (p) { if (p.oficioId === o.id) p.oficioId = null; });
      o.pdfId = t.value ? +t.value : null;
      if (o.pdfId) pdfs.filter(function (p) { return p.id === o.pdfId; })[0].oficioId = o.id;
      pintar();
    } else if (t.dataset.campo === "municipio") {
      var om = buscar(t.dataset.id);
      om.municipio = t.value.trim().toUpperCase();
      pintar();
    } else if (t.dataset.campo === "para") {
      pintar();
    }
  });

  document.addEventListener("toggle", function (e) {
    var d = e.target;
    if (d.classList && d.classList.contains("cuerpo") && d.open) mostrarVista(buscar(d.dataset.id));
  }, true);

  document.addEventListener("click", function (e) {
    var b = e.target.closest ? e.target.closest("button[data-act]") : null;
    if (!b) return;
    var act = b.dataset.act;

    if (act === "quitarPdf") {
      pdfs = pdfs.filter(function (p) { return p.id !== +b.dataset.id; });
      pintar();
      return;
    }
    var o = buscar(b.dataset.id);
    if (!o) return;

    if (act === "quitar") {
      pdfs.forEach(function (p) { if (p.oficioId === o.id) p.oficioId = null; });
      oficios = oficios.filter(function (x) { return x !== o; });
      pintar();
    } else if (act === "copiar") {
      copiarTexto(b, $(b.dataset.campo + "-" + o.id));
    } else if (act === "copiarCuerpo") {
      copiarCuerpo(b, o);
    } else if (act === "eml") {
      var st = estado(o);
      if (st.clase !== "ok") { mensaje(o.municipio + ": " + st.texto.toLowerCase() + ". Corrígelo antes de descargar.", "err"); return; }
      O.descargar(O.nombreEml(o) + ".eml", O.armarEml(datosEml(o)), "message/rfc822");
    }
  });

  /* ---------- Barra de acciones ---------- */
  $("btnZip").addEventListener("click", async function () {
    var ok = listos();
    if (!ok.length) return;
    var zip = new JSZip(), usados = {};
    ok.forEach(function (o) {
      var base = O.nombreEml(o), nombre = base, k = 2;
      while (usados[nombre]) nombre = base + " (" + (k++) + ")";
      usados[nombre] = true;
      zip.file(nombre + ".eml", O.armarEml(datosEml(o)));
    });
    var blob = await zip.generateAsync({ type: "blob", compression: "DEFLATE" });
    O.descargar("correos-oficios.zip", blob);
    var faltan = oficios.length - ok.length;
    mensaje("Descargado correos-oficios.zip con " + ok.length + " correo" + (ok.length > 1 ? "s" : "") + "." +
      (faltan ? " " + faltan + " Word por revisar no se incluyeron." : ""), "ok");
  });

  $("btnCsv").addEventListener("click", function () {
    var filas = oficios.map(function (o) {
      var p = pdfDe(o);
      return [o.municipio, o.nombre, p ? p.nombre : "", o.para, o.cc, o.asunto, estado(o).texto];
    }).concat(pdfsSinWord().map(function (p) { return ["", "", p.nombre, "", "", "", "PDF sin Word"]; }));
    O.descargar("listado-oficios.csv", O.armarCsv(filas), "text/csv;charset=utf-8");
  });

  /* ---------- Zona de carga ---------- */
  var zona = $("drop"), entrada = $("file");
  zona.addEventListener("click", function () { entrada.click(); });
  zona.addEventListener("keydown", function (e) {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); entrada.click(); }
  });
  entrada.addEventListener("change", function () { agregarArchivos(entrada.files); entrada.value = ""; });
  ["dragenter", "dragover"].forEach(function (ev) {
    zona.addEventListener(ev, function (e) { e.preventDefault(); zona.classList.add("over"); });
  });
  ["dragleave", "drop"].forEach(function (ev) {
    zona.addEventListener(ev, function (e) { e.preventDefault(); zona.classList.remove("over"); });
  });
  zona.addEventListener("drop", function (e) {
    if (e.dataTransfer && e.dataTransfer.files) agregarArchivos(e.dataTransfer.files);
  });

  pintar();
})(window.Oficios);
