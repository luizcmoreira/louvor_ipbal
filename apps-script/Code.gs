/**
 * API do Ministério de Louvor — cola isto no Apps Script vinculado à planilha
 * (Extensões → Apps Script), substituindo o conteúdo do arquivo Code.gs.
 *
 * Abas esperadas na planilha, com cabeçalho na linha 1:
 *
 * Musicas:      title | originalKey | tom_Ricardo | tom_Kariny | tom_Luiz | chordpro | status | isNew | videoCompleta | videoCongregacional
 * Escala:       date | label | ministro | musicos | vocalBacking | repertorio
 *               (musicos: "Nome:Instrumento, Nome:Instrumento" — ex: "Luiz:Teclado, Caio:Guitarra")
 *               (vocalBacking / repertorio: lista separada por vírgula)
 * Participantes: name | vocal | instrumentos
 *               (vocal: ministro / ministra / backing / vazio)
 *               (instrumentos: lista separada por vírgula)
 */

function doGet(e) {
  const sheetName = (e.parameter.sheet || "Musicas");
  const data = readSheet(sheetName);
  return jsonOut(data);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (body.action === "addMusica") {
    ss.getSheetByName("Musicas").appendRow([
      body.title, body.originalKey || "",
      body.tom_Ricardo || "", body.tom_Kariny || "", body.tom_Luiz || "",
      body.chordpro || "", "nao", !!body.isNew,
      body.videoCompleta || "", body.videoCongregacional || "",
    ]);
  } else if (body.action === "updateMusica") {
    updateRowByMatch("Musicas", "title", body.matchTitle, body);
  } else if (body.action === "addEscala") {
    ss.getSheetByName("Escala").appendRow([
      body.date, body.label || "", body.ministro || "",
      body.musicos || "", body.vocalBacking || "", body.repertorio || "",
    ]);
  } else {
    return jsonOut({ ok: false, error: "ação desconhecida" });
  }
  return jsonOut({ ok: true });
}

function updateRowByMatch(sheetName, matchCol, matchVal, patch) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const mi = headers.indexOf(matchCol);
  for (let r = 1; r < values.length; r++) {
    if (values[r][mi] === matchVal) {
      headers.forEach((h, i) => {
        if (Object.prototype.hasOwnProperty.call(patch, h)) {
          sheet.getRange(r + 1, i + 1).setValue(patch[h]);
        }
      });
      break;
    }
  }
}

function readSheet(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const rows = values.slice(1).filter(r => r[0] !== "");
  const objs = rows.map(r => {
    const o = {};
    headers.forEach((h, i) => o[h] = r[i]);
    return o;
  });
  if (name === "Musicas") return objs.map(mapMusica);
  if (name === "Escala") return objs.map(mapEscala);
  if (name === "Participantes") return objs.map(mapParticipante);
  return objs;
}

function mapMusica(o) {
  const presets = {};
  if (o.tom_Ricardo) presets["Ricardo"] = o.tom_Ricardo;
  if (o.tom_Kariny) presets["Kariny"] = o.tom_Kariny;
  if (o.tom_Luiz) presets["Luiz"] = o.tom_Luiz;
  return {
    title: o.title,
    originalKey: o.originalKey || "",
    presets: presets,
    chordpro: (o.chordpro || "").toString(),
    status: o.status || "nao",
    isNew: o.isNew === true || o.isNew === "TRUE" || o.isNew === "true",
    videoCompleta: o.videoCompleta || "",
    videoCongregacional: o.videoCongregacional || "",
  };
}

function mapEscala(o) {
  return {
    date: formatDate(o.date),
    label: o.label || "",
    ministro: o.ministro || "",
    musicos: parsePairs(o.musicos),
    vocalBacking: parseList(o.vocalBacking),
    repertorio: parseList(o.repertorio),
  };
}

function mapParticipante(o) {
  return {
    name: o.name,
    vocal: o.vocal || null,
    instrumentos: parseList(o.instrumentos),
  };
}

function parsePairs(s) {
  if (!s) return [];
  return String(s).split(",").map(p => p.trim()).filter(Boolean).map(p => {
    const parts = p.split(":");
    return { name: (parts[0] || "").trim(), instrumento: (parts[1] || "").trim() };
  });
}

function parseList(s) {
  if (!s) return [];
  return String(s).split(",").map(x => x.trim()).filter(Boolean);
}

function formatDate(d) {
  if (Object.prototype.toString.call(d) === "[object Date]") {
    return Utilities.formatDate(d, "America/Sao_Paulo", "yyyy-MM-dd");
  }
  return d;
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
