/**
 * API do Ministério de Louvor — cola isto no Apps Script vinculado à planilha
 * (Extensões → Apps Script), substituindo o conteúdo do arquivo Code.gs.
 *
 * Abas esperadas na planilha, com cabeçalho na linha 1:
 *
 * Musicas:      title | originalKey | tom_Ricardo | tom_Kariny | tom_Luiz | chordpro | status | isNew | videoCompleta | videoCongregacional
 * Escala:       date | label | ministro | musicos | vocalBacking | repertorio | ensaioData | ensaioHorario | obsAntes | obsDepois
 *               (musicos: "Nome:Instrumento, Nome:Instrumento" — ex: "Luiz:Teclado, Caio:Guitarra")
 *               (vocalBacking / repertorio: lista separada por vírgula)
 *               (ensaioData: yyyy-mm-dd do ensaio prévio, separado da data do culto)
 *               (obsAntes/obsDepois: observações livres, antes e depois do ensaio)
 * Participantes: name | vocal | instrumentos
 *               (vocal: ministro / ministra / backing / vazio)
 *               (instrumentos: lista separada por vírgula)
 * Sugestoes:    title | artist
 */

function doGet(e) {
  // e vem undefined se alguém rodar doGet direto no editor do Apps Script
  // (botão "Executar"), sem passar por uma requisição HTTP de verdade.
  const params = (e && e.parameter) || {};
  const sheetName = params.sheet || "Musicas";
  const data = readSheet(sheetName);
  return jsonOut(data);
}

function doPost(e) {
  if (!e || !e.postData) return jsonOut({ ok: false, error: "sem dados (rodou fora de uma requisição de verdade?)" });
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
    // Se o título mudou, atualiza também toda referência a essa música na
    // Escala — senão o repertório de cultos já cadastrados fica apontando
    // pro nome antigo (a música "some" da tela de Ensaios, tom vira "?").
    if (body.title && body.title !== body.matchTitle) {
      renameInEscalaRepertorio(body.matchTitle, body.title);
    }
  } else if (body.action === "deleteMusica") {
    deleteRowByMatch("Musicas", { title: body.title });
  } else if (body.action === "addEscala") {
    ss.getSheetByName("Escala").appendRow([
      body.date, body.label || "", body.ministro || "",
      body.musicos || "", body.vocalBacking || "", body.repertorio || "",
      body.ensaioData || "", body.ensaioHorario || "", body.obsAntes || "", body.obsDepois || "",
    ]);
  } else if (body.action === "updateEscala") {
    updateRowByMultiMatch("Escala",
      { date: body.matchDate, label: body.matchLabel },
      { date: body.date, label: body.label, ministro: body.ministro, musicos: body.musicos, vocalBacking: body.vocalBacking, repertorio: body.repertorio,
        ensaioData: body.ensaioData, ensaioHorario: body.ensaioHorario, obsAntes: body.obsAntes, obsDepois: body.obsDepois });
  } else if (body.action === "deleteEscala") {
    deleteRowByMatch("Escala", { date: body.date, label: body.label });
  } else if (body.action === "addSugestao") {
    ss.getSheetByName("Sugestoes").appendRow([body.title || "", body.artist || ""]);
  } else if (body.action === "updateSugestao") {
    updateRowByMultiMatch("Sugestoes",
      { title: body.matchTitle, artist: body.matchArtist || "" },
      { title: body.title || "", artist: body.artist || "" });
  } else if (body.action === "deleteSugestao") {
    deleteRowByMatch("Sugestoes", { title: body.title, artist: body.artist || "" });
  } else {
    return jsonOut({ ok: false, error: "ação desconhecida" });
  }
  return jsonOut({ ok: true });
}

function renameInEscalaRepertorio(oldTitle, newTitle) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Escala");
  if (!sheet) return;
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const repIdx = headers.indexOf("repertorio");
  if (repIdx === -1) return;
  for (let r = 1; r < values.length; r++) {
    const cell = values[r][repIdx];
    if (!cell) continue;
    const titles = String(cell).split(",").map(t => t.trim());
    if (!titles.includes(oldTitle)) continue;
    const updated = titles.map(t => t === oldTitle ? newTitle : t).join(", ");
    sheet.getRange(r + 1, repIdx + 1).setValue(updated);
  }
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

// Igual updateRowByMatch, mas identifica a linha por VÁRIAS colunas ao mesmo
// tempo (ex: date + label, já que Escala não tem uma coluna única por si só).
function updateRowByMultiMatch(sheetName, matchCols, patch) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const cols = Object.keys(matchCols).map(k => ({ key: k, idx: headers.indexOf(k) }));
  for (let r = 1; r < values.length; r++) {
    const isMatch = cols.every(({ key, idx }) => {
      let cell = values[r][idx];
      if (key === "date") cell = formatDate(cell);
      return cell === matchCols[key];
    });
    if (isMatch) {
      headers.forEach((h, i) => {
        if (Object.prototype.hasOwnProperty.call(patch, h) && patch[h] !== undefined) {
          sheet.getRange(r + 1, i + 1).setValue(patch[h]);
        }
      });
      return true;
    }
  }
  return false;
}

// Apaga a PRIMEIRA linha cujas colunas em `matchCols` batem todas com o valor
// pedido (ex: { title: "Te agradeço" } ou { date: "2026-08-16", label: "..." }).
// Datas são comparadas já formatadas (yyyy-MM-dd), porque a planilha guarda a
// coluna "date" como objeto Date, não como texto puro.
function deleteRowByMatch(sheetName, matchCols) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const cols = Object.keys(matchCols).map(k => ({ key: k, idx: headers.indexOf(k) }));
  for (let r = 1; r < values.length; r++) {
    const isMatch = cols.every(({ key, idx }) => {
      let cell = values[r][idx];
      if (key === "date") cell = formatDate(cell);
      return cell === matchCols[key];
    });
    if (isMatch) {
      sheet.deleteRow(r + 1);
      return true;
    }
  }
  return false;
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
  if (name === "Sugestoes") return objs.map(mapSugestao);
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
    repertorio: parseRepertorio(o.repertorio),
    ensaioData: formatDate(o.ensaioData) || "",
    ensaioHorario: formatTime(o.ensaioHorario) || "",
    obsAntes: o.obsAntes || "",
    obsDepois: o.obsDepois || "",
  };
}

function mapParticipante(o) {
  return {
    name: o.name,
    vocal: o.vocal || null,
    instrumentos: parseList(o.instrumentos),
  };
}

function mapSugestao(o) {
  return {
    title: o.title || "",
    artist: o.artist || "",
  };
}

function parsePairs(s) {
  if (!s) return [];
  return String(s).split(",").map(p => p.trim()).filter(Boolean).map(p => {
    const parts = p.split(":");
    return { name: (parts[0] || "").trim(), instrumento: (parts[1] || "").trim() };
  });
}

// Repertório é uma lista ORDENADA (a ordem em que a música toca no culto).
// Um item que começa com "#" é um título de seção (ex: "#Prelúdio"), o resto
// é música de verdade — assim dá pra intercalar seções e músicas mantendo
// tudo numa coluna só, na ordem exata em que foi montado.
function parseRepertorio(s) {
  if (!s) return [];
  return String(s).split(",").map(t => t.trim()).filter(Boolean).map(t => {
    if (t.startsWith("#")) return { type: "header", text: t.slice(1).trim() };
    return { type: "song", text: t };
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

// A planilha converte um valor tipo "19:30" pra um Date interno dela mesma
// (ancorado em 30/12/1899) — sem isso, ensaioHorario voltava um timestamp
// completo em vez de só "HH:mm".
function formatTime(d) {
  if (Object.prototype.toString.call(d) === "[object Date]") {
    return Utilities.formatDate(d, "America/Sao_Paulo", "HH:mm");
  }
  return d;
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
