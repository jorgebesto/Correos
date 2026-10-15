/* app.js — Interfaz: carga de archivos, fichas, botones y descargas. */
(function (O) {
  "use strict";

  var $ = function (id) { return document.getElementById(id); };
  var excluir = $("exclude"), adjuntar = $("attach"), lista = $("list");
  var oficios = [], contador = 0;

  try { excluir.value = localStorage.getItem("oficios-excluir") || ""; } catch (e) {}

  /* ---------- Crear un oficio a partir del texto ---------- */
  function crearOficio(texto, nombre, bytes, ext) {
    texto = O.limpiarTexto(texto);
    var correos = O.buscarCorreos(texto);
    return {
      id: ++contador, nombre: nombre, ext: ext, bytes: bytes, texto: texto,
      correos: correos,
      para: correos.filter(function (c) { return !O.estaExcluido(c, excluir.value); }).join("; "),
      paraEditado: false,
      asunto: O.buscarAsunto(texto),
      cuerpo: texto,
      radicado: O.buscarRadicado(texto),
      error: ""
    };
  }

  function oficioConError(archivo, mensaje) {
    return { id: ++contador, nombre: archivo.name, ext: "", bytes: null, texto: "", correos: [], para: "",
             paraEditado: false, asunto: "", cuerpo: "", radicado: "", error: mensaje };
  }

  async function agregarArchivos(archivos) {
    archivos = [].slice.call(archivos);
    if (!archivos.length) return;
    oficios = oficios.filter(function (o) { return !o.ejemplo; });
    mensaje("Leyendo " + archivos.length + " archivo" + (archivos.length > 1 ? "s" : "") + "…", "ok");

    for (var i = 0; i < archivos.length; i++) {
      try {
        var r = await O.leerArchivo(archivos[i]);
        var o = crearOficio(r.texto, archivos[i].name, r.bytes, r.ext);
        if (!o.texto) {
          o.error = r.ext === "pdf"
            ? "Este PDF no tiene texto (parece escaneado). Exporta el oficio desde Word como PDF o súbelo en .docx."
            : "El archivo está vacío.";
        }
        oficios.push(o);
      } catch (e) {
        oficios.push(oficioConError(archivos[i], e.message || "No se pudo leer el archivo."));
      }
    }
    pintar();
    mensaje("");
  }

  /* ---------- Estado de cada oficio ---------- */
  function estado(o) {
    if (o.error) return { clase: "err", texto: "No se pudo leer" };
    if (!o.para.trim()) return { clase: "warn", texto: "Falta destinatario" };
    if (!o.asunto.trim()) return { clase: "warn", texto: "Falta asunto" };
    return { clase: "ok", texto: "Listo" };
  }
  function listos() { return oficios.filter(function (o) { return estado(o).clase === "ok"; }); }
  function buscar(id) { id = +id; return oficios.filter(function (o) { return o.id === id; })[0]; }

  /* ---------- Pintar la lista ---------- */
  function esc(s) {
    return String(s).replace(/[&<>"]/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c];
    });
  }

  function fila(o, campo, etiqueta, control) {
    return '<div class="row"><div class="rowhead">' +
      '<label for="' + campo + "-" + o.id + '">' + etiqueta + "</label>" +
      '<button type="button" data-act="copiar" data-campo="' + campo + '" data-id="' + o.id + '">Copiar</button>' +
      "</div>" + control + "</div>";
  }

  function ficha(o) {
    var st = estado(o);
    var etiquetas =
      (o.ejemplo ? '<span class="pill ex">Ejemplo</span>' : "") +
      (o.radicado ? '<span class="rad" title="Radicado">' + esc(o.radicado) + "</span>" : "") +
      '<span class="pill ' + st.clase + '" data-pill="' + o.id + '">' + st.texto + "</span>";
    var cabeza = '<div class="head"><div class="name">' + esc(o.nombre) + '</div><div class="tags">' + etiquetas + "</div></div>";

    if (o.error) {
      return '<article class="card">' + cabeza +
        '<p class="msg err">' + esc(o.error) + "</p>" +
        '<div class="foot"><span></span><button type="button" data-act="quitar" data-id="' + o.id + '">Quitar</button></div>' +
        "</article>";
    }

    var n = o.correos.length;
    var encontrados = n
      ? n + " correo" + (n > 1 ? "s" : "") + " encontrado" + (n > 1 ? "s" : "") + " en el oficio"
      : "No encontré correos en el oficio; escríbelos a mano";

    return '<article class="card' + (o.ejemplo ? " sample" : "") + '">' + cabeza +
      fila(o, "para", "Para", '<input type="text" class="mono" id="para-' + o.id + '" data-campo="para" data-id="' + o.id + '" value="' + esc(o.para) + '">') +
      '<p class="note">' + encontrados + ". Separa varios con punto y coma.</p>" +
      fila(o, "asunto", "Asunto", '<input type="text" id="asunto-' + o.id + '" data-campo="asunto" data-id="' + o.id + '" value="' + esc(o.asunto) + '">') +
      fila(o, "cuerpo", "Cuerpo del correo (texto completo del oficio)", '<textarea id="cuerpo-' + o.id + '" data-campo="cuerpo" data-id="' + o.id + '">' + esc(o.cuerpo) + "</textarea>") +
      '<div class="foot"><span class="note">' + (adjuntar.checked ? "Se adjunta " + esc(o.nombre) : "Sin adjunto") + "</span>" +
      '<div class="btns">' +
      '<button type="button" data-act="eml" data-id="' + o.id + '">Descargar este correo (.eml)</button>' +
      '<button type="button" data-act="quitar" data-id="' + o.id + '">Quitar</button>' +
      "</div></div></article>";
  }

  function pintar() {
    lista.innerHTML = oficios.length
      ? oficios.map(ficha).join("")
      : '<div class="empty">Todavía no hay oficios. Cárgalos arriba para ver aquí un correo por cada uno.</div>';
    actualizarContador();
  }

  function actualizarContador() {
    var n = listos().length, porRevisar = oficios.length - n;
    var reales = oficios.filter(function (o) { return !o.ejemplo; }).length;
    $("count").innerHTML = oficios.length
      ? "<b>" + n + "</b> correo" + (n === 1 ? "" : "s") + " listo" + (n === 1 ? "" : "s") +
        (porRevisar ? " · <b>" + porRevisar + "</b> por revisar" : "") +
        (!reales ? ' <span class="note">(ejemplo)</span>' : "")
      : "Sin oficios cargados";
    $("btnZip").disabled = !n;
    $("btnCsv").disabled = !oficios.length;
  }

  function mensaje(texto, tipo) {
    var m = $("msg");
    m.hidden = !texto;
    m.textContent = texto || "";
    m.className = "msg " + (tipo || "ok");
  }

  /* ---------- Botones de la barra ---------- */
  $("btnZip").addEventListener("click", async function () {
    var ok = listos();
    if (!ok.length) return;
    var zip = new JSZip(), usados = {};
    ok.forEach(function (o) {
      var base = O.nombreEml(o), nombre = base, k = 2;
      while (usados[nombre]) nombre = base + " (" + (k++) + ")";
      usados[nombre] = true;
      zip.file(nombre + ".eml", O.armarEml(o, adjuntar.checked));
    });
    var blob = await zip.generateAsync({ type: "blob", compression: "DEFLATE" });
    O.descargar("correos-oficios.zip", blob);
    var omitidos = oficios.length - ok.length;
    mensaje("Descargado correos-oficios.zip con " + ok.length + " correo" + (ok.length > 1 ? "s" : "") + "." +
      (omitidos ? " " + omitidos + " oficio" + (omitidos > 1 ? "s" : "") + " por revisar no se incluyeron." : ""), "ok");
  });

  $("btnCsv").addEventListener("click", function () {
    O.descargar("listado-oficios.csv", O.armarCsv(oficios, estado), "text/csv;charset=utf-8");
  });

  /* ---------- Edición dentro de las fichas ---------- */
  lista.addEventListener("input", function (e) {
    var t = e.target, o = buscar(t.dataset.id);
    if (!o || !t.dataset.campo) return;
    o[t.dataset.campo] = t.value;
    if (t.dataset.campo === "para") o.paraEditado = true;
    var st = estado(o), pill = document.querySelector('[data-pill="' + o.id + '"]');
    if (pill) { pill.className = "pill " + st.clase; pill.textContent = st.texto; }
    actualizarContador();
  });

  lista.addEventListener("click", function (e) {
    var b = e.target.closest("button");
    if (!b) return;
    var o = buscar(b.dataset.id);
    if (!o) return;

    if (b.dataset.act === "quitar") {
      oficios = oficios.filter(function (x) { return x !== o; });
      pintar();
    }

    if (b.dataset.act === "eml") {
      if (estado(o).clase !== "ok") { mensaje("A este oficio le falta " + estado(o).texto.replace("Falta ", "") + ".", "err"); return; }
      O.descargar(O.nombreEml(o) + ".eml", O.armarEml(o, adjuntar.checked), "message/rfc822");
    }

    if (b.dataset.act === "copiar") {
      var campo = $(b.dataset.campo + "-" + o.id);
      var listo = function () { b.textContent = "Copiado"; setTimeout(function () { b.textContent = "Copiar"; }, 1400); };
      var manual = function () { campo.select(); b.textContent = "Ctrl+C"; };
      try { navigator.clipboard.writeText(campo.value).then(listo, manual); } catch (err) { manual(); }
    }
  });

  /* ---------- Ajustes ---------- */
  excluir.addEventListener("input", function () {
    try { localStorage.setItem("oficios-excluir", excluir.value); } catch (e) {}
    oficios.forEach(function (o) {
      if (!o.paraEditado && !o.error) {
        o.para = o.correos.filter(function (c) { return !O.estaExcluido(c, excluir.value); }).join("; ");
      }
    });
    pintar();
  });
  adjuntar.addEventListener("change", pintar);

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

  /* ---------- Oficio de ejemplo (se quita al cargar los tuyos) ---------- */
  var EJEMPLO = [
    "Bogotá D.C., 9 de octubre de 2026",
    "Radicado: 20264100242441",
    "",
    "Señores",
    "ALCALDÍA MUNICIPAL DE GINEBRA (VALLE)",
    "Secretaría de Planeación",
    "planeacion@municipio-ejemplo.gov.co",
    "",
    "ASUNTO: Solicitud - Envío de certificación clasificación del uso del suelo según el POT vigente, para actualización de datos básicos ante la Oficina de Registro de Instrumentos Públicos - GINEBRA",
    "",
    "Respetados señores:",
    "",
    "De manera atenta solicitamos la expedición de la certificación de clasificación del uso del suelo de los predios relacionados en el anexo, según el Plan de Ordenamiento Territorial vigente, con el fin de adelantar la actualización de datos básicos ante la Oficina de Registro de Instrumentos Públicos, conforme a la Instrucción Administrativa 03 de 2019 y la Resolución 7448 de 2021.",
    "",
    "Agradecemos remitir la respuesta a este mismo correo.",
    "",
    "Cordialmente,",
    "",
    "Nombre del funcionario",
    "Cargo"
  ].join("\n");

  var ejemplo = crearOficio(EJEMPLO, "Oficio de ejemplo - Ginebra.txt", new TextEncoder().encode(EJEMPLO), "txt");
  ejemplo.ejemplo = true;
  oficios.push(ejemplo);
  pintar();
})(window.Oficios);
