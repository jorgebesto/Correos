/* docx-html.js — Lee un Word (.docx) y lo convierte a HTML conservando el formato:
   negrillas, cursivas, subrayado, letra, tamaño, color, alineación, sangrías,
   tablas con sus bordes y las imágenes incrustadas. */
window.Oficios = window.Oficios || {};

(function (O) {
  "use strict";

  /* ---------- Utilidades XML ---------- */
  function hijos(n, nombre) {
    var r = [];
    for (var c = n ? n.firstChild : null; c; c = c.nextSibling) {
      if (c.nodeType === 1 && (!nombre || c.nodeName === nombre)) r.push(c);
    }
    return r;
  }
  function primero(n, nombre) { return hijos(n, nombre)[0] || null; }
  function attr(n, a) { return n ? n.getAttribute(a) : null; }
  function val(n) { return attr(n, "w:val"); }
  function encendido(n) {               // <w:b/>, <w:b w:val="0"/>…
    if (!n) return undefined;
    var v = val(n);
    return !(v === "0" || v === "false" || v === "none");
  }
  function esc(s) { return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }
  function pt(twips) { return Math.round(twips / 20 * 100) / 100 + "pt"; }
  function mezclar() {
    var r = {};
    for (var i = 0; i < arguments.length; i++) {
      var o = arguments[i];
      if (o) for (var k in o) if (o[k] !== undefined) r[k] = o[k];
    }
    return r;
  }

  var RESALTADO = {
    yellow: "#FFFF00", green: "#00FF00", cyan: "#00FFFF", magenta: "#FF00FF", blue: "#0000FF", red: "#FF0000",
    darkBlue: "#000080", darkCyan: "#008080", darkGreen: "#008000", darkMagenta: "#800080", darkRed: "#800000",
    darkYellow: "#808000", darkGray: "#808080", lightGray: "#C0C0C0", black: "#000000", white: "#FFFFFF"
  };
  // Letra de todo el cuerpo del correo (reemplaza la del Word).
  O.FUENTE_CUERPO = "'Century Gothic',CenturyGothic,AppleGothic,sans-serif";

  var TIPO_IMG = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp" };

  /* ---------- Propiedades de texto y párrafo ---------- */
  function leerRPr(rPr) {
    if (!rPr) return {};
    var o = {};
    var f = primero(rPr, "w:rFonts");
    if (f) o.fuente = attr(f, "w:ascii") || attr(f, "w:hAnsi") || undefined;
    var sz = primero(rPr, "w:sz");
    if (sz) o.tam = +val(sz) / 2;
    o.b = encendido(primero(rPr, "w:b"));
    o.i = encendido(primero(rPr, "w:i"));
    var u = primero(rPr, "w:u");
    if (u) o.u = val(u) !== "none";
    o.tachado = encendido(primero(rPr, "w:strike"));
    o.mayus = encendido(primero(rPr, "w:caps"));
    var c = primero(rPr, "w:color");
    if (c) o.color = val(c) === "auto" ? "#000000" : "#" + val(c);
    var h = primero(rPr, "w:highlight");
    if (h) o.resaltado = RESALTADO[val(h)] || undefined;
    var shd = primero(rPr, "w:shd");
    if (shd) { var fill = attr(shd, "w:fill"); o.fondo = fill && fill !== "auto" && fill.toUpperCase() !== "FFFFFF" ? "#" + fill : undefined; }
    var va = primero(rPr, "w:vertAlign");
    if (va) o.vert = val(va);
    var rs = primero(rPr, "w:rStyle");
    if (rs) o.estiloCar = val(rs);
    return o;
  }

  function leerPPr(pPr) {
    if (!pPr) return {};
    var o = {};
    var jc = primero(pPr, "w:jc");
    if (jc) o.alin = { both: "justify", distribute: "justify", center: "center", right: "right", end: "right", left: "left", start: "left" }[val(jc)];
    var ind = primero(pPr, "w:ind");
    if (ind) {
      var l = attr(ind, "w:left") || attr(ind, "w:start"), r = attr(ind, "w:right") || attr(ind, "w:end");
      if (l !== null) o.izq = +l;
      if (r !== null) o.der = +r;
      if (attr(ind, "w:firstLine") !== null) o.primera = +attr(ind, "w:firstLine");
      if (attr(ind, "w:hanging") !== null) o.primera = -attr(ind, "w:hanging");
    }
    var sp = primero(pPr, "w:spacing");
    if (sp) {
      if (attr(sp, "w:before") !== null) o.antes = +attr(sp, "w:before");
      if (attr(sp, "w:after") !== null) o.despues = +attr(sp, "w:after");
      if (attr(sp, "w:line") !== null) { o.linea = +attr(sp, "w:line"); o.reglaLinea = attr(sp, "w:lineRule") || "auto"; }
    }
    var ps = primero(pPr, "w:pStyle");
    if (ps) o.estilo = val(ps);
    return o;
  }

  /* ---------- Estilos del documento ---------- */
  function leerEstilos(xml) {
    var est = { porId: {}, rDef: {}, pDef: {}, parrafoDef: null, tablaDef: null };
    if (!xml) return est;
    var dd = xml.getElementsByTagName("w:docDefaults")[0];
    if (dd) {
      var rd = dd.getElementsByTagName("w:rPrDefault")[0], pd = dd.getElementsByTagName("w:pPrDefault")[0];
      est.rDef = leerRPr(rd && primero(rd, "w:rPr"));
      est.pDef = leerPPr(pd && primero(pd, "w:pPr"));
    }
    var lista = xml.getElementsByTagName("w:style");
    for (var i = 0; i < lista.length; i++) {
      var s = lista[i], id = attr(s, "w:styleId"), tipo = attr(s, "w:type");
      est.porId[id] = {
        tipo: tipo,
        base: val(primero(s, "w:basedOn")),
        r: leerRPr(primero(s, "w:rPr")),
        p: leerPPr(primero(s, "w:pPr")),
        tblPr: primero(s, "w:tblPr"),
        tcPr: primero(s, "w:tcPr")
      };
      if (attr(s, "w:default") === "1" || attr(s, "w:default") === "true") {
        if (tipo === "paragraph") est.parrafoDef = id;
        if (tipo === "table") est.tablaDef = id;
      }
    }
    return est;
  }

  // Cadena de estilos (el estilo y los que hereda), del más general al más específico.
  function cadena(est, id) {
    var r = [], vistos = {};
    while (id && est.porId[id] && !vistos[id]) { vistos[id] = 1; r.unshift(est.porId[id]); id = est.porId[id].base; }
    return r;
  }

  /* ---------- Conversor ---------- */
  function Conversor(doc, rels, estilos, medios) {
    this.doc = doc; this.rels = rels; this.est = estilos; this.medios = medios;
    this.imagenes = []; this.cidPorRel = {};
  }

  Conversor.prototype.rPrEstiloCar = function (id) {
    return mezclar.apply(null, cadena(this.est, id).map(function (s) { return s.r; }));
  };

  Conversor.prototype.cssTexto = function (r) {
    var css = [];
    if (r.fuente) css.push("font-family:" + O.FUENTE_CUERPO);
    if (r.tam) css.push("font-size:" + r.tam + "pt");
    if (r.b) css.push("font-weight:bold");
    if (r.i) css.push("font-style:italic");
    var deco = [];
    if (r.u) deco.push("underline");
    if (r.tachado) deco.push("line-through");
    if (deco.length) css.push("text-decoration:" + deco.join(" "));
    if (r.color) css.push("color:" + r.color);
    if (r.resaltado || r.fondo) css.push("background:" + (r.resaltado || r.fondo));
    if (r.mayus) css.push("text-transform:uppercase");
    if (r.vert === "superscript") css.push("vertical-align:super;font-size:smaller");
    if (r.vert === "subscript") css.push("vertical-align:sub;font-size:smaller");
    return css.join(";");
  };

  Conversor.prototype.imagen = function (dibujo) {
    var blip = dibujo.getElementsByTagName("a:blip")[0];
    if (!blip) return "";
    var rid = attr(blip, "r:embed") || attr(blip, "r:link");
    var medio = this.medios[rid];
    if (!medio) return "";
    var ext = dibujo.getElementsByTagName("wp:extent")[0];
    var ancho = ext ? Math.round(+attr(ext, "cx") / 9525) : null;
    var alto = ext ? Math.round(+attr(ext, "cy") / 9525) : null;
    var cid = this.cidPorRel[rid];
    if (!cid) {
      cid = "img" + (this.imagenes.length + 1) + "." + Date.now().toString(36) + "@oficio";
      this.cidPorRel[rid] = cid;
      this.imagenes.push({ cid: cid, nombre: medio.nombre, tipo: medio.tipo, bytes: medio.bytes });
    }
    return '<img src="cid:' + cid + '"' + (ancho ? ' width="' + ancho + '" height="' + alto + '" style="width:' + ancho + "px;height:" + alto + 'px"' : "") + ' alt="">';
  };

  // Recorre el contenido de un párrafo (runs, hipervínculos, controles…).
  Conversor.prototype.contenido = function (nodo, base) {
    var html = "", self = this;
    hijos(nodo).forEach(function (n) {
      switch (n.nodeName) {
        case "w:r":
          html += self.run(n, base);
          break;
        case "w:hyperlink":
          var rid = attr(n, "r:id"), destino = rid && self.rels[rid] ? self.rels[rid].destino : null;
          var dentro = self.contenido(n, base);
          html += destino ? '<a href="' + esc(destino) + '" style="color:inherit">' + dentro + "</a>" : dentro;
          break;
        case "w:smartTag": case "w:ins": case "w:fldSimple": case "w:customXml":
          html += self.contenido(n, base);
          break;
        case "w:sdt":
          html += self.contenido(primero(n, "w:sdtContent"), base);
          break;
      }
    });
    return html;
  };

  Conversor.prototype.run = function (r, base) {
    var rp = leerRPr(primero(r, "w:rPr"));
    var props = mezclar(base, rp.estiloCar ? this.rPrEstiloCar(rp.estiloCar) : null, rp);
    var dentro = "", self = this;
    hijos(r).forEach(function (n) {
      switch (n.nodeName) {
        case "w:t":
          dentro += esc(n.textContent).replace(/ {2}/g, " &nbsp;").replace(/^ /, "&nbsp;");
          break;
        case "w:tab": dentro += "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"; break;
        case "w:br": if (attr(n, "w:type") !== "page") dentro += "<br>"; break;
        case "w:noBreakHyphen": dentro += "-"; break;
        case "w:drawing": dentro += self.imagen(n); break;
        case "mc:AlternateContent":
          var d = n.getElementsByTagName("w:drawing")[0];
          if (d) dentro += self.imagen(d);
          break;
      }
    });
    if (!dentro) return "";
    var css = this.cssTexto(props);
    return css ? '<span style="' + css + '">' + dentro + "</span>" : dentro;
  };

  Conversor.prototype.parrafo = function (p) {
    var pp = leerPPr(primero(p, "w:pPr"));
    var idEstilo = pp.estilo || this.est.parrafoDef;
    var cad = cadena(this.est, idEstilo);
    var propsP = mezclar.apply(null, [this.est.pDef].concat(cad.map(function (s) { return s.p; })).concat([pp]));
    var baseR = mezclar.apply(null, [this.est.rDef].concat(cad.map(function (s) { return s.r; })));

    var html = this.contenido(p, baseR);
    var css = ["margin:0"];
    if (propsP.antes) css.push("margin-top:" + pt(propsP.antes));
    if (propsP.despues) css.push("margin-bottom:" + pt(propsP.despues));
    if (propsP.alin) css.push("text-align:" + propsP.alin);
    if (propsP.izq) css.push("margin-left:" + pt(propsP.izq));
    if (propsP.der && propsP.der > 0) css.push("margin-right:" + pt(propsP.der));
    if (propsP.primera) css.push("text-indent:" + pt(propsP.primera));
    if (propsP.linea) {
      css.push(propsP.reglaLinea === "auto" ? "line-height:" + Math.round(propsP.linea / 240 * 100) + "%" : "line-height:" + pt(propsP.linea));
    }

    if (!html.replace(/<[^>]+>|&nbsp;|\s/g, "") && html.indexOf("<img") === -1) {
      // Párrafo vacío: conserva el alto de la línea en blanco.
      var marca = mezclar(baseR, leerRPr(primero(primero(p, "w:pPr"), "w:rPr")));
      var cssM = this.cssTexto({ fuente: marca.fuente, tam: marca.tam });
      return '<p style="' + css.join(";") + '"><span style="' + cssM + '">&nbsp;</span></p>';
    }
    // Un salto de línea al final del párrafo crea una línea en blanco en Word; en HTML hay que forzarla.
    if (/<br>(<\/span>|<\/a>)*$/.test(html)) html += "&nbsp;";
    var cssBase = this.cssTexto({ fuente: baseR.fuente, tam: baseR.tam });
    return '<p style="' + css.join(";") + (cssBase ? ";" + cssBase : "") + '">' + html + "</p>";
  };

  function borde(n) {
    if (!n) return null;
    var v = val(n);
    if (!v || v === "nil" || v === "none") return "none";
    var ancho = Math.max((+attr(n, "w:sz") || 4) / 8, 0.5);
    var color = attr(n, "w:color");
    color = !color || color === "auto" ? "#000000" : "#" + color;
    var estilo = v === "double" ? "double" : v === "dotted" ? "dotted" : /dash/.test(v) ? "dashed" : "solid";
    if (estilo === "double") ancho = Math.max(ancho, 2.25);
    return ancho + "pt " + estilo + " " + color;
  }
  function leerBordes(cont) {
    var o = {};
    if (!cont) return o;
    ["top", "left", "bottom", "right", "insideH", "insideV", "start", "end"].forEach(function (k) {
      var b = borde(primero(cont, "w:" + k));
      if (b) o[k === "start" ? "left" : k === "end" ? "right" : k] = b;
    });
    return o;
  }
  function leerMargenes(cont, o) {
    if (!cont) return o;
    ["top", "left", "bottom", "right", "start", "end"].forEach(function (k) {
      var m = primero(cont, "w:" + k);
      if (m) o[k === "start" ? "left" : k === "end" ? "right" : k] = +attr(m, "w:w");
    });
    return o;
  }

  Conversor.prototype.tabla = function (tbl) {
    var self = this;
    var tblPr = primero(tbl, "w:tblPr");
    var idEstilo = val(primero(tblPr, "w:tblStyle")) || this.est.tablaDef;
    var cad = cadena(this.est, idEstilo);

    var bordes = {}, margenes = { top: 0, bottom: 0, left: 108, right: 108 };
    cad.concat([{ tblPr: tblPr }]).forEach(function (s) {
      if (!s.tblPr) return;
      bordes = mezclar(bordes, leerBordes(primero(s.tblPr, "w:tblBorders")));
      leerMargenes(primero(s.tblPr, "w:tblCellMar"), margenes);
    });

    var css = ["border-collapse:collapse", "mso-table-lspace:0", "mso-table-rspace:0"];
    var w = primero(tblPr, "w:tblW");
    if (w && attr(w, "w:type") === "dxa" && +attr(w, "w:w") > 0) css.push("width:" + pt(+attr(w, "w:w")));
    if (w && attr(w, "w:type") === "pct") css.push("width:" + (+attr(w, "w:w") / 50) + "%");
    var ind = primero(tblPr, "w:tblInd");
    if (ind && +attr(ind, "w:w")) css.push("margin-left:" + pt(+attr(ind, "w:w")));
    var jc = primero(tblPr, "w:jc");
    if (jc && val(jc) === "center") css.push("margin-left:auto;margin-right:auto");

    var filas = hijos(tbl, "w:tr");
    // Mapa de celdas para calcular combinaciones verticales (vMerge).
    var mapa = filas.map(function (tr) {
      var col = 0;
      return hijos(tr, "w:tc").map(function (tc) {
        var tcPr = primero(tc, "w:tcPr");
        var span = +val(primero(tcPr, "w:gridSpan")) || 1;
        var vm = primero(tcPr, "w:vMerge");
        var c = { tc: tc, tcPr: tcPr, col: col, span: span, vm: vm ? (val(vm) === "restart" ? "inicio" : "sigue") : null, filas: 1 };
        col += span;
        return c;
      });
    });
    var totalCols = Math.max.apply(null, mapa.map(function (f) { return f.reduce(function (a, c) { return a + c.span; }, 0); }).concat([1]));
    mapa.forEach(function (fila, i) {
      fila.forEach(function (c) {
        if (c.vm !== "inicio") return;
        for (var j = i + 1; j < mapa.length; j++) {
          var abajo = mapa[j].filter(function (x) { return x.col === c.col; })[0];
          if (abajo && abajo.vm === "sigue") c.filas++; else break;
        }
      });
    });

    var html = '<table cellspacing="0" cellpadding="0" style="' + css.join(";") + '">';
    mapa.forEach(function (fila, i) {
      var trPr = primero(filas[i], "w:trPr"), alto = primero(trPr, "w:trHeight");
      html += "<tr" + (alto && +val(alto) > 0 ? ' style="height:' + pt(+val(alto)) + '"' : "") + ">";
      fila.forEach(function (c) {
        if (c.vm === "sigue") return;
        var ultimaFila = i + c.filas - 1 === mapa.length - 1;
        var b = mezclar({
          top: i === 0 ? bordes.top : bordes.insideH,
          bottom: ultimaFila ? bordes.bottom : bordes.insideH,
          left: c.col === 0 ? bordes.left : bordes.insideV,
          right: c.col + c.span >= totalCols ? bordes.right : bordes.insideV
        }, leerBordes(primero(c.tcPr, "w:tcBorders")));
        var m = leerMargenes(primero(c.tcPr, "w:tcMar"), mezclar(margenes));
        var cssC = [
          "border-top:" + (b.top || "none"), "border-bottom:" + (b.bottom || "none"),
          "border-left:" + (b.left || "none"), "border-right:" + (b.right || "none"),
          "padding:" + pt(m.top) + " " + pt(m.right) + " " + pt(m.bottom) + " " + pt(m.left)
        ];
        var tcW = primero(c.tcPr, "w:tcW");
        if (tcW && attr(tcW, "w:type") === "dxa" && +attr(tcW, "w:w") > 0) cssC.push("width:" + pt(+attr(tcW, "w:w")));
        var va = primero(c.tcPr, "w:vAlign");
        cssC.push("vertical-align:" + (va ? ({ center: "middle", bottom: "bottom" }[val(va)] || "top") : "top"));
        var shd = primero(c.tcPr, "w:shd");
        if (shd && attr(shd, "w:fill") && attr(shd, "w:fill") !== "auto") cssC.push("background:#" + attr(shd, "w:fill"));
        html += "<td" + (c.span > 1 ? ' colspan="' + c.span + '"' : "") + (c.filas > 1 ? ' rowspan="' + c.filas + '"' : "") +
                ' style="' + cssC.join(";") + '">' + self.bloques(hijos(c.tc)) + "</td>";
      });
      html += "</tr>";
    });
    return html + "</table>";
  };

  Conversor.prototype.bloques = function (lista) {
    var self = this;
    return lista.map(function (n) {
      if (n.nodeName === "w:p") return self.parrafo(n);
      if (n.nodeName === "w:tbl") return self.tabla(n);
      if (n.nodeName === "w:sdt") return self.bloques(hijos(primero(n, "w:sdtContent")));
      return "";
    }).join("");
  };

  /* ---------- Texto plano de un bloque ---------- */
  function texto(n) {
    var s = "";
    (function rec(x) {
      for (var c = x.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue;
        if (c.nodeName === "w:t") s += c.textContent;
        else if (c.nodeName === "w:tab") s += "\t";
        else if (c.nodeName === "w:br") s += "\n";
        else if (c.nodeName === "w:del" || c.nodeName === "w:instrText") continue;
        else { rec(c); if (c.nodeName === "w:p") s += "\n"; }
      }
    })(n);
    return n.nodeName === "w:p" ? s + "\n" : s;
  }

  /* ---------- Entrada principal ---------- */
  function rutaMedio(destino) {
    if (destino.indexOf("/") === 0) return destino.slice(1);
    var partes = ("word/" + destino).split("/"), r = [];
    partes.forEach(function (p) { if (p === "..") r.pop(); else if (p && p !== ".") r.push(p); });
    return r.join("/");
  }

  /**
   * Lee un .docx. Devuelve:
   *  - bloques: [{ texto }] uno por párrafo o tabla del cuerpo
   *  - convertir(desde, hasta): { html, imagenes } con el formato de esos bloques
   */
  O.leerDocx = async function (buffer) {
    var zip = await JSZip.loadAsync(buffer);
    var parser = new DOMParser();
    var leerXml = async function (ruta) {
      var f = zip.file(ruta);
      return f ? parser.parseFromString(await f.async("string"), "application/xml") : null;
    };

    var doc = await leerXml("word/document.xml");
    if (!doc) throw new Error("No parece un archivo de Word válido.");
    var estilos = leerEstilos(await leerXml("word/styles.xml"));

    var rels = {}, medios = {};
    var relsXml = await leerXml("word/_rels/document.xml.rels");
    if (relsXml) {
      var lista = relsXml.getElementsByTagName("Relationship");
      for (var i = 0; i < lista.length; i++) {
        var rel = lista[i], id = attr(rel, "Id"), tipo = attr(rel, "Type") || "", destino = attr(rel, "Target");
        rels[id] = { tipo: tipo, destino: destino };
        if (/\/image$/.test(tipo) && attr(rel, "TargetMode") !== "External") {
          var ruta = rutaMedio(destino), arch = zip.file(ruta);
          var ext = (ruta.split(".").pop() || "").toLowerCase();
          if (arch && TIPO_IMG[ext]) {
            medios[id] = { nombre: ruta.split("/").pop(), tipo: TIPO_IMG[ext], bytes: await arch.async("uint8array") };
          }
        }
      }
    }

    var cuerpo = doc.getElementsByTagName("w:body")[0];
    var nodos = [];
    hijos(cuerpo).forEach(function (n) {
      if (n.nodeName === "w:p" || n.nodeName === "w:tbl") nodos.push(n);
      else if (n.nodeName === "w:sdt") hijos(primero(n, "w:sdtContent")).forEach(function (x) { nodos.push(x); });
    });

    return {
      bloques: nodos.map(function (n) { return { texto: texto(n) }; }),
      convertir: function (desde, hasta) {
        var c = new Conversor(doc, rels, estilos, medios);
        var html = c.bloques(nodos.slice(desde, hasta + 1));
        return { html: html, imagenes: c.imagenes };
      }
    };
  };

  /** Cambia las imágenes cid: por imágenes incrustadas (para la vista previa y para copiar). */
  O.htmlConImagenes = function (html, imagenes) {
    imagenes.forEach(function (im) {
      var s = "", TROZO = 0x8000;
      for (var i = 0; i < im.bytes.length; i += TROZO) s += String.fromCharCode.apply(null, im.bytes.subarray(i, i + TROZO));
      html = html.split("cid:" + im.cid).join("data:" + im.tipo + ";base64," + btoa(s));
    });
    return html;
  };
})(window.Oficios);
