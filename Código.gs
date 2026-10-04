// ===== Configuração =====
const ID_PLANILHA = "COLE_O_ID_AQUI"; // o mesmo ID que você já usa
const ABAS = ["REMUNERAÇÃO", "DESPESAS GERAIS", "BASE DE DADOS", "COFRINHOS"];
const MESES = ["JANEIRO","FEVEREIRO","MARÇO","ABRIL","MAIO","JUNHO","JULHO","AGOSTO","SETEMBRO","OUTUBRO","NOVEMBRO","DEZEMBRO"];
const up = s => String(s == null ? "" : s).trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();

function autorizado(t) {
  const token = PropertiesService.getScriptProperties().getProperty("TOKEN");
  return !!token && !!t && t === token;
}
function saida(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ===== Leitura (app abre e atualiza) =====
function doGet(e) {
  if (!e || !e.parameter || !autorizado(e.parameter.token)) return saida({ erro: "Não autorizado" });
  const ss = SpreadsheetApp.openById(ID_PLANILHA), resultado = {};
  ABAS.forEach(nome => {
    const aba = ss.getSheetByName(nome);
    if (!aba) { resultado[nome] = { erro: "Aba não encontrada" }; return; }
    const v = aba.getDataRange().getDisplayValues();
    resultado[nome] = { cabecalho: v[0], linhas: v.slice(1) };
  });
  return saida(resultado);
}

// ===== Gravação (lançamentos feitos pelo app) =====
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const p = JSON.parse(e.postData.contents);
    if (!autorizado(p.token)) return saida({ erro: "Não autorizado" });
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    if (p.acao === "despesa") return saida(lancarDespesa(ss, p));
    if (p.acao === "receita") return saida(lancarReceita(ss, p));
    if (p.acao === "excluir") return saida(excluir(ss, p));
    return saida({ erro: "Ação inválida" });
  } catch (err) {
    return saida({ erro: "Falha ao gravar: " + err.message });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

// Lê grupos/descrições/formas de pagamento da BASE DE DADOS (grupo aparece só na 1ª linha de cada bloco)
function lerBase(ss) {
  const v = ss.getSheetByName("BASE DE DADOS").getDataRange().getValues();
  const h = v.findIndex(r => r.map(up).includes("GRUPO"));
  const c = v[h].map(up), ig = c.indexOf("GRUPO"), id = c.indexOf("DESCRICAO"), it = c.indexOf("TIPO DE PAGAMENTO");
  const out = { grupos: {}, tipos: [] };
  let g = "";
  v.slice(h + 1).forEach(r => {
    if (String(r[ig]).trim()) g = String(r[ig]).trim();
    const d = String(r[id]).trim();
    if (g && d) (out.grupos[g] = out.grupos[g] || []).push(d);
    const t = it >= 0 ? String(r[it]).trim() : "";
    if (t && out.tipos.indexOf(t) < 0) out.tipos.push(t);
  });
  return out;
}

function lancarDespesa(ss, p) {
  const v = Number(p.valor);
  if (!(v > 0 && v < 10000000)) return { erro: "Valor inválido" };
  const mi = MESES.map(up).indexOf(up(p.mes));
  if (mi < 0) return { erro: "Mês inválido" };
  const dt = /^\d{2}\/\d{2}\/\d{4}$/;
  if (!dt.test(p.data || "") || (p.fechamento && !dt.test(p.fechamento))) return { erro: "Data inválida" };
  const base = lerBase(ss);
  if (!base.grupos[p.grupo] || base.grupos[p.grupo].indexOf(p.descricao) < 0) return { erro: "Grupo ou descrição não existe na BASE DE DADOS" };
  if (base.tipos.length && base.tipos.indexOf(p.tipo) < 0) return { erro: "Forma de pagamento não existe na BASE DE DADOS" };

  const aba = ss.getSheetByName("DESPESAS GERAIS"), dados = aba.getDataRange().getValues();
  const h = dados.findIndex(r => { const c = r.map(up); return c.includes("VALOR") && c.includes("GRUPO"); });
  if (h < 0) return { erro: "Cabeçalho da aba DESPESAS GERAIS não encontrado" };
  const cab = dados[h].map(up), col = n => cab.indexOf(n);
  // primeira linha livre = primeira com VALOR vazio abaixo do cabeçalho
  let linha = h + 1;
  while (linha < dados.length && String(dados[linha][col("VALOR")]).trim() !== "") linha++;
  if (linha + 1 > aba.getMaxRows()) aba.insertRowsAfter(aba.getMaxRows(), 1);

  const campos = {
    "MES FATURA": MESES[mi], "FECHAMENTO DA FATURA": p.fechamento || "", "DATA DA COMPRA": p.data,
    "TIPO DE PAGAMENTO": p.tipo, "GRUPO": p.grupo, "DESCRICAO": p.descricao, "VALOR": v, "SITUACAO": String(p.situacao || "").slice(0, 40)
  };
  Object.keys(campos).forEach(n => { // só preenche o que veio; não toca em colunas com fórmula
    if (campos[n] !== "" && col(n) >= 0) aba.getRange(linha + 1, col(n) + 1).setValue(campos[n]);
  });
  return { ok: true, linha: linha + 1 };
}

function lancarReceita(ss, p) {
  const v = Number(p.valor);
  if (!isFinite(v) || v < 0 || v >= 10000000) return { erro: "Valor inválido" };
  const aba = ss.getSheetByName("REMUNERAÇÃO"), dados = aba.getDataRange().getValues();
  const h = dados.findIndex(r => up(r[1]) === "JANEIRO");
  if (h < 0) return { erro: "Linha dos meses não encontrada" };
  const c = dados[h].findIndex(x => up(x) === up(p.mes));
  const l = dados.findIndex((r, i) => i > h && up(r[0]) === up(p.evento));
  if (c < 1 || l < 0) return { erro: "Evento ou mês não encontrado" };
  const cel = aba.getRange(l + 1, c + 1);
  if (cel.getFormula()) return { erro: "Essa célula tem fórmula e não foi alterada" };
  const atual = Number(dados[l][c]) || 0;
  cel.setValue(p.modo === "somar" ? Math.round((atual + v) * 100) / 100 : v);
  return { ok: true };
}

// ===== Exclusão: limpa o conteúdo (não remove a linha, para não deslocar nada na planilha) =====
const nm = v => typeof v === "number" ? v : (parseFloat(String(v).replace(/[^\d,.-]/g, "").replace(/\./g, "").replace(",", ".")) || 0);

function excluir(ss, p) {
  if (p.aba === "receita") {
    const aba = ss.getSheetByName("REMUNERAÇÃO"), dados = aba.getDataRange().getValues();
    const h = dados.findIndex(r => up(r[1]) === "JANEIRO");
    const c = h < 0 ? -1 : dados[h].findIndex(x => up(x) === up(p.mes));
    const l = dados.findIndex((r, i) => i > h && up(r[0]) === up(p.evento));
    if (c < 1 || l < 0) return { erro: "Evento ou mês não encontrado" };
    if (Math.abs(nm(dados[l][c]) - Number(p.esperado)) > 0.005) return { erro: "A planilha mudou. Atualize o app e tente de novo." };
    const cel = aba.getRange(l + 1, c + 1);
    if (cel.getFormula()) return { erro: "Essa célula tem fórmula. Apague direto na planilha." };
    cel.clearContent();
    return { ok: true };
  }
  const desp = p.aba === "despesa";
  if (!desp && p.aba !== "cofrinho") return { erro: "Tipo inválido" };
  const aba = ss.getSheetByName(desp ? "DESPESAS GERAIS" : "COFRINHOS"), dados = aba.getDataRange().getValues();
  const chave = desp ? ["VALOR", "GRUPO"] : ["VALOR", "COFRINHO"];
  const h = dados.findIndex(r => { const c = r.map(up); return chave.every(k => c.includes(k)); });
  const linha = Number(p.linha);
  if (h < 0 || !(linha > h + 1 && linha <= dados.length)) return { erro: "Linha inválida" };
  const cab = dados[h].map(up), r = dados[linha - 1], col = n => cab.indexOf(n), e = p.esperado || {};
  // confere se a linha ainda é o mesmo lançamento que o app mostrou
  const confere = Math.abs(nm(r[col("VALOR")]) - Number(e.valor)) < 0.005 &&
    (desp ? up(r[col("GRUPO")]) === up(e.grupo) && up(r[col("DESCRICAO")]) === up(e.desc)
          : up(r[col("COFRINHO")]) === up(e.nome));
  if (!confere) return { erro: "A planilha mudou. Atualize o app e tente de novo." };
  if (aba.getRange(linha, col("VALOR") + 1).getFormula()) return { erro: "O valor é uma fórmula. Apague direto na planilha." };
  // despesa: apaga a linha toda; cofrinho: mantém mês e nome do cofrinho e limpa valor, situação e obs.
  const cols = desp ? ["MES FATURA", "FECHAMENTO DA FATURA", "DATA DA COMPRA", "TIPO DE PAGAMENTO", "GRUPO", "DESCRICAO", "VALOR", "SITUACAO"] : ["VALOR", "SITUACAO", "OBS."];
  cols.forEach(n => {
    if (col(n) < 0) return;
    const cel = aba.getRange(linha, col(n) + 1);
    if (!cel.getFormula()) cel.clearContent();
  });
  return { ok: true };
}
