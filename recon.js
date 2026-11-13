#!/usr/bin/env node
// Recon de contatos: lê uma lista de empresas (nome, "nome; cidade" ou CNPJ) e procura na web
// telefones, celulares e WhatsApp públicos. Gera um CSV que abre direto no Excel.
// Fontes: DuckDuckGo, site oficial da empresa (+ página de contato), guias de empresas e
// cadastro público do CNPJ (BrasilAPI). Sem dependências — só Node 18+.
//
// Uso:  node recon.js lista.txt [--cidade "Campinas/SP"] [--saida resultados.csv] [--novo]

import { readFileSync, writeFileSync, appendFileSync, existsSync } from "node:fs";
import { basename } from "node:path";

// ------------------------------------------------------------ util
const UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0 Safari/537.36";
const DDDS = new Set("11 12 13 14 15 16 17 18 19 21 22 24 27 28 31 32 33 34 35 37 38 41 42 43 44 45 46 47 48 49 51 53 54 55 61 62 63 64 65 66 67 68 69 71 73 74 75 77 79 81 82 83 84 85 86 87 88 89 91 92 93 94 95 96 97 98 99".split(" "));
// Não adianta abrir (exigem login, não são da empresa ou são marketplaces)
const NAO_ABRIR = /(facebook|instagram|linkedin|twitter|x\.com|tiktok|youtube|wikipedia|google|bing|duckduckgo|mercadolivre|magazineluiza|amazon|shopee|americanas|reclameaqui|glassdoor|indeed|infojobs|catho|olx|vivareal|zapimoveis)\./i;
const DIRETORIOS = /(guiamais|apontador|telelistas|solutudo|econodata|cnpj\.biz|cnpja|casadosdados|empresascnpj|consultacnpj|listamais|guiafacil|cylex|hagah|infobel|kompass|applocal|paginaamarela|eguias|guiadeempresas|empresaqui|telefone\.|reclamacoes\.net|adital|encontrasp|encontra[a-z]+\.com|achei|negocio\.site|guiato|portaldoconsumidor)/i;

const pausa = (ms) => new Promise((r) => setTimeout(r, ms));
const semAcento = (s) => String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
const decodificar = (s) => s.replace(/&amp;/g, "&").replace(/&quot;/g, '"').replace(/&#39;|&#x27;/g, "'").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&nbsp;/g, " ")
  .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(+n)).replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCharCode(parseInt(h, 16)));
const semTags = (h) => decodificar(h.replace(/<script[\s\S]*?<\/script>|<style[\s\S]*?<\/style>/gi, " ").replace(/<[^>]+>/g, " ")).replace(/\s+/g, " ");

async function baixar(url, ms = 12000) {
  const r = await fetch(url, { headers: { "User-Agent": UA, "Accept-Language": "pt-BR,pt;q=0.9", Accept: "text/html,*/*" }, redirect: "follow", signal: AbortSignal.timeout(ms) });
  if (!r.ok && r.status !== 202) throw new Error(`HTTP ${r.status}`);
  const tipo = r.headers.get("content-type") || "";
  if (tipo && !/html|text|json/.test(tipo)) throw new Error("não é página");
  return { status: r.status, texto: (await r.text()).slice(0, 1_500_000), url: r.url };
}

// ------------------------------------------------------------ telefones
function normalizarTelefone(bruto) {
  let d = String(bruto || "").replace(/\D/g, "");
  if (d.length >= 12 && d.startsWith("55")) d = d.slice(2);
  d = d.replace(/^0+/, "");
  if (d.length !== 10 && d.length !== 11) return null;
  if (!DDDS.has(d.slice(0, 2))) return null;
  if (d.length === 11 && d[2] !== "9") return null;
  if (d.length === 10 && !/[2-5]/.test(d[2])) return null; // fixo começa com 2–5
  if (/^(\d)\1+$/.test(d.slice(2))) return null;
  return d;
}
const formatarTelefone = (d) => d.length === 11 ? `(${d.slice(0, 2)}) ${d.slice(2, 7)}-${d.slice(7)}` : `(${d.slice(0, 2)}) ${d.slice(2, 6)}-${d.slice(6)}`;
const ehCelular = (d) => d.length === 11;

// Telefones escritos no texto: exige formatação (parênteses/espaço/hífen) para não confundir com CNPJ, CEP, códigos.
const RE_TEL = /(?<![\d/.-])(?:\+?55[\s.-]?)?\(?0?([1-9]{2})\)?[\s.-]{0,2}(9[\s.]?\d{4}|[2-5]\d{3})[\s.-](\d{4})(?![\d/-])/g;
function telefonesNoTexto(texto) {
  const achados = [];
  for (const m of texto.matchAll(RE_TEL)) {
    const d = normalizarTelefone(`${m[1]}${m[2]}${m[3]}`);
    if (!d) continue;
    const contexto = semAcento(texto.slice(Math.max(0, m.index - 70), m.index + m[0].length + 25));
    // "WhatsApp: (19) 9…" — o rótulo vem logo antes do número (e não é o rótulo do número anterior)
    const antes = semAcento(texto.slice(Math.max(0, m.index - 30), m.index)).split(/\d{4}[-\s]?\d{4}/).pop();
    achados.push({ d, bruto: m[0], whatsTexto: /whats|zap|wpp/.test(antes), contatoTexto: /tel|fone|contato|ligue|atendimento|celular|whats|zap|sac/.test(contexto) });
  }
  return achados;
}
function whatsappNoHtml(html) {
  const out = [];
  const re = /(?:wa\.me\/|api\.whatsapp\.com\/send\/?\?(?:[^"'\s>]*?&(?:amp;)?)?phone=|web\.whatsapp\.com\/send\?phone=|whatsapp:\/\/send\?phone=)\+?(\d[\d\s-]{9,16})/gi;
  for (const m of html.matchAll(re)) { const d = normalizarTelefone(m[1]); if (d) out.push(d); }
  return out;
}
const telLinks = (html) => [...html.matchAll(/href=["']tel:([^"']+)["']/gi)].map((m) => { try { return normalizarTelefone(decodeURIComponent(m[1])); } catch { return null; } }).filter(Boolean);
const cnpjNoTexto = (t) => t.match(/\b\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}\b/)?.[0].replace(/\D/g, "") || null;
const emailsNoTexto = (t) => [...new Set((t.match(/[\w.+-]+@[\w-]+\.[\w.-]*[a-z]/gi) || []).map((e) => e.toLowerCase())
  .filter((e) => !/\.(png|jpe?g|gif|webp|svg)$/.test(e) && !/sentry|example|wixpress|domain|seudominio|email\.com$/.test(e)))].slice(0, 3);
const instagramNoHtml = (h) => h.match(/instagram\.com\/([a-z0-9_.]{2,30})\/?["'?]/i)?.[1]?.replace(/^(p|reel|explore|accounts)$/i, "") || "";

// ------------------------------------------------------------ buscador (DuckDuckGo)
let ddgLiberadoEm = 0;
async function buscar(consulta) {
  const espera = ddgLiberadoEm - Date.now();
  if (espera > 0) await pausa(espera);
  ddgLiberadoEm = Date.now() + 4500 + Math.random() * 2000; // pausa entre buscas para não ser bloqueado
  const versoes = [
    { url: `https://lite.duckduckgo.com/lite/?q=${encodeURIComponent(consulta)}&kl=br-pt`, link: /<a([^>]*class=['"]result-link['"][^>]*)>([\s\S]*?)<\/a>/g, trecho: /<td[^>]*class=['"]result-snippet['"][^>]*>([\s\S]*?)<\/td>/g },
    { url: `https://html.duckduckgo.com/html/?q=${encodeURIComponent(consulta)}&kl=br-pt`, link: /<a([^>]*class="result__a"[^>]*)>([\s\S]*?)<\/a>/g, trecho: /<a[^>]*class="result__snippet"[^>]*>([\s\S]*?)<\/a>/g },
  ];
  for (const v of versoes) {
    const { status, texto } = await baixar(v.url);
    if (status === 202 || /anomaly|captcha/i.test(texto)) continue;
    const links = [...texto.matchAll(v.link)];
    const trechos = [...texto.matchAll(v.trecho)].map((m) => semTags(m[1]));
    return links.map((m, i) => {
      let href = decodificar(m[1].match(/href=['"]([^'"]+)['"]/)?.[1] || "");
      const uddg = href.match(/[?&]uddg=([^&]+)/)?.[1];
      if (uddg) href = decodeURIComponent(uddg);
      if (href.startsWith("//")) href = "https:" + href;
      return { url: href, titulo: semTags(m[2]), trecho: trechos[i] || "" };
    }).filter((r) => /^https?:\/\//.test(r.url) && !/duckduckgo\.com\/(y\.js|duckduckgo-help)/.test(r.url));
  }
  throw new Error("BLOQUEIO");
}

// ------------------------------------------------------------ fontes
const palavrasChave = (nome) => semAcento(nome)
  .replace(/\b(ltda|me|eireli|sa|s\/a|epp|comercio|comercial|de|da|do|das|dos|e|industria|servicos|loja|lojas|grupo|cia)\b/g, " ")
  .split(/[^a-z0-9]+/).filter((p) => p.length > 2);
// Palavras de ramo: sozinhas não identificam a empresa ("Ferragens Rocha" ≠ "Donna Ferragens")
const GENERICAS = new Set(("ferragens ferragem ferramentas ferramenta maquinas maquina materiais material construcao construcoes eletrica eletricos " +
  "eletrico hidraulica hidraulicos tintas tinta madeiras madeireira casa auto pecas center centro distribuidora distribuidor atacado " +
  "atacadista varejo brasil industrial store shop online parafusos equipamentos solucoes tecnologia agro rural imports importadora").split(" "));

function relevancia(nomes, r, soDominio = false) {
  let alvo;
  try { alvo = semAcento(soDominio ? new URL(r.url).hostname : `${decodeURIComponent(new URL(r.url).href)} ${r.titulo}`); } catch { return 0; }
  return Math.max(0, ...[].concat(nomes).filter(Boolean).map((nome) => {
    const ps = palavrasChave(nome);
    if (!ps.length) return 0;
    const distintas = ps.filter((p) => !GENERICAS.has(p));
    if (distintas.length && !distintas.every((p) => alvo.includes(p))) return 0;
    return ps.filter((p) => alvo.includes(p)).length / ps.length;
  }));
}
const paginaDeBloqueio = (url, html) => /perfdrive|captcha|challenge|cf-chl|validate\./i.test(url) || /<title>[^<]*(just a moment|attention required|access denied|captcha)/i.test(html);

// Tenta achar o site oficial pelo nome: "Casa do Construtor" -> casadoconstrutor.com.br
async function adivinharSite(nome) {
  const slug = semAcento(nome).replace(/\b(ltda|me|eireli|epp|s\/?a)\b/g, "").replace(/[^a-z0-9]/g, "");
  if (slug.length < 4 || slug.length > 40) return null;
  const distintas = palavrasChave(nome).filter((p) => !GENERICAS.has(p));
  for (const url of [`https://www.${slug}.com.br`, `https://${slug}.com.br`, `https://www.${slug}.com`]) {
    try {
      const pag = await baixar(url, 7000);
      if (!new URL(pag.url).hostname.includes(slug.slice(0, 6))) continue; // redirecionou para outro lugar
      const alvo = semAcento(semTags(pag.texto.slice(0, 200_000)));
      if (!paginaDeBloqueio(pag.url, pag.texto) && distintas.length && !distintas.every((p) => alvo.includes(p))) continue;
      return { url: pag.url, pagina: pag };
    } catch { /* tenta o próximo */ }
  }
  return null;
}

async function consultarCNPJ(cnpj) {
  const r = await fetch(`https://brasilapi.com.br/api/cnpj/v1/${cnpj}`, { headers: { "User-Agent": UA }, signal: AbortSignal.timeout(12000) });
  if (!r.ok) return null;
  const j = await r.json();
  return {
    razao: j.razao_social, fantasia: j.nome_fantasia, cidade: j.municipio && `${j.municipio}/${j.uf}`, situacao: j.descricao_situacao_cadastral,
    telefones: [j.ddd_telefone_1, j.ddd_telefone_2].map(normalizarTelefone).filter(Boolean), email: j.email || null,
  };
}

// ------------------------------------------------------------ pesquisa de uma empresa
async function pesquisarEmpresa({ nome: nomeEntrada, cidade: cidadeEntrada, cnpj: cnpjEntrada }) {
  const cnpjDireto = cnpjEntrada || null;
  const nome = nomeEntrada || cnpjDireto || "";
  const res = { empresa: nome, cidade: cidadeEntrada || "", site: "", instagram: "", cnpj: cnpjDireto || "", situacao: "", emails: [], fontes: new Set(), obs: [] };
  const tels = new Map();
  const add = (d, origem, peso, whats = false) => {
    const t = tels.get(d) || { whats: false, origens: new Set(), peso: 0 };
    t.whats ||= whats;
    if (!t.origens.has(origem)) t.peso += peso; // a mesma fonte repetida não soma de novo
    t.origens.add(origem);
    tels.set(d, t);
  };
  const usarCNPJ = (c, peso) => {
    c.telefones.forEach((d) => add(d, "Cadastro CNPJ", peso));
    if (c.email) res.emails.push(c.email.toLowerCase());
    res.situacao ||= c.situacao || "";
    res.fontes.add("CNPJ");
  };

  let nomeBusca = nome;
  let nomesRef = [nome];
  if (cnpjDireto) {
    const c = await consultarCNPJ(cnpjDireto).catch(() => null);
    if (c) {
      if (!nomeEntrada) res.empresa = c.fantasia ? `${c.fantasia} (${c.razao})` : c.razao || nome;
      res.cidade ||= c.cidade || "";
      usarCNPJ(c, 2);
      // razão social identifica melhor; nome fantasia costuma ser genérico ("Matriz")
      nomeBusca = nomeEntrada || (c.razao || c.fantasia || nome).replace(/\b(ltda|me|eireli|epp|s\.?\/?a\.?)\b\.?/gi, "").trim();
      nomesRef = [nomeEntrada, c.razao, c.fantasia];
    } else res.obs.push("CNPJ não encontrado no cadastro público");
  }

  const cidadeAlvo = res.cidade;
  const cidadeChave = semAcento(cidadeAlvo.split("/")[0].split(",")[0].trim());
  const pertoDaCidade = (texto, idx) => !cidadeChave || semAcento(texto.slice(Math.max(0, idx - 300), idx + 300)).includes(cidadeChave);

  // 1) site oficial pelo nome (rápido, sem buscador)
  const adivinhado = nomeBusca && !/^\d+$/.test(nomeBusca) ? await adivinharSite(nomeBusca) : null;

  // 2) buscador: trechos dos resultados que são mesmo dessa empresa
  let resultados = await buscar(`${nomeBusca} ${cidadeAlvo} whatsapp telefone`.trim());
  if (resultados.filter((r) => relevancia(nomesRef, r) >= 0.5).length < 2) {
    resultados = resultados.concat(await buscar(`"${nomeBusca}" ${cidadeAlvo} contato`.trim()));
  }
  res.fontes.add("Busca");
  const relevantes = resultados.map((r) => ({ ...r, rel: relevancia(nomesRef, r), oficial: relevancia(nomesRef, r, true) >= 0.5 && !NAO_ABRIR.test(r.url) && !DIRETORIOS.test(r.url) }))
    .filter((r) => r.rel >= 0.5);
  if (resultados.length && !relevantes.length) res.obs.push("Nenhum resultado com o nome dessa empresa");
  for (const r of relevantes) {
    for (const t of telefonesNoTexto(`${r.titulo} ${r.trecho}`)) {
      add(t.d, r.oficial ? "Site da empresa" : DIRETORIOS.test(r.url) ? "Guia de empresas" : "Resultado da busca", r.oficial ? 2 : 1, t.whatsTexto);
    }
    if (!res.instagram) res.instagram = r.url.match(/instagram\.com\/([a-z0-9_.]{2,30})/i)?.[1] || "";
  }

  // 3) abrir páginas: site oficial (+ página de contato) e até 2 guias de empresas
  const oficial = relevantes.filter((r) => r.oficial).sort((a, b) => b.rel - a.rel)[0] || (adivinhado && { url: adivinhado.url, pagina: adivinhado.pagina });
  const guias = relevantes.filter((r) => DIRETORIOS.test(r.url) && !NAO_ABRIR.test(r.url)).slice(0, 2);
  const abrir = [...(oficial ? [{ ...oficial, tipo: "site" }] : []), ...guias.map((g) => ({ ...g, tipo: "guia" }))];

  for (const p of abrir) {
    try {
      const pag = p.pagina || await baixar(p.url);
      if (paginaDeBloqueio(pag.url, pag.texto)) throw new Error("site bloqueia robôs");
      const htmls = [pag.texto];
      if (p.tipo === "site") {
        res.site = new URL(pag.url).origin;
        const contato = [...pag.texto.matchAll(/href=["']([^"'#]*(?:contato|fale-conosco|contact|atendimento|lojas)[^"'#]*)["']/gi)]
          .map((m) => { try { return new URL(decodificar(m[1]), pag.url).href; } catch { return null; } })
          .find((u) => u && u.startsWith(res.site) && u !== pag.url);
        if (contato) { try { htmls.push((await baixar(contato)).texto); } catch { /* sem página de contato */ } }
      }
      const origem = p.tipo === "site" ? "Site da empresa" : "Guia de empresas";
      const peso = p.tipo === "site" ? 1 : 0.5;
      for (const h of htmls) {
        whatsappNoHtml(h).forEach((d) => add(d, p.tipo === "site" ? "Link de WhatsApp no site" : "WhatsApp em guia", 5 * peso, true));
        telLinks(h).forEach((d) => add(d, origem, 3 * peso));
        const txt = semTags(h);
        // em guias aparecem telefones de outras empresas: só vale o que está perto de "telefone/whatsapp"
        telefonesNoTexto(txt).filter((t) => p.tipo === "site" || t.contatoTexto).slice(0, p.tipo === "site" ? 25 : 4)
          .forEach((t) => add(t.d, origem, (pertoDaCidade(txt, txt.indexOf(t.bruto)) ? 2 : 0.7) * peso, t.whatsTexto));
        if (p.tipo === "site") {
          res.emails.push(...emailsNoTexto(txt));
          res.cnpj ||= cnpjNoTexto(txt) || "";
          res.instagram ||= instagramNoHtml(h);
        }
      }
      res.fontes.add(p.tipo === "site" ? "Site" : "Guias");
    } catch (e) {
      if (p.tipo === "site") { res.site ||= new URL(p.url).origin; res.obs.push(`Site não abriu (${e.message})`); }
    }
  }

  // 4) CNPJ achado no site → telefone do cadastro
  if (res.cnpj && !cnpjDireto) {
    const c = await consultarCNPJ(res.cnpj).catch(() => null);
    if (c) usarCNPJ(c, 1.5);
  }

  // Classificação final
  const lista = [...tels.entries()].map(([d, v]) => ({
    d, numero: formatarTelefone(d),
    tipo: v.whats ? "WhatsApp" : ehCelular(d) ? "Celular" : "Fixo",
    confianca: v.peso >= 3 || v.origens.has("Link de WhatsApp no site") ? "alta" : v.peso >= 1.5 || v.origens.size >= 2 ? "média" : "baixa",
    origens: [...v.origens],
    ordem: v.peso + (v.whats ? 10 : 0) + (ehCelular(d) ? 1 : 0),
  })).sort((a, b) => b.ordem - a.ordem);

  res.whatsapp = lista.filter((t) => (t.tipo === "WhatsApp" && t.confianca !== "baixa") || (t.tipo === "Celular" && t.confianca === "alta")).slice(0, 4);
  res.celulares = lista.filter((t) => !res.whatsapp.includes(t) && ehCelular(t.d)).slice(0, 4);
  res.fixos = lista.filter((t) => !res.whatsapp.includes(t) && !ehCelular(t.d)).slice(0, 4);
  res.emails = [...new Set(res.emails)].slice(0, 3);
  res.fontes = [...res.fontes];
  if (!lista.length) res.obs.push("Nenhum telefone público encontrado");
  return res;
}

// ------------------------------------------------------------ entrada / saída
function lerArgs(argv) {
  const a = { arquivo: null, cidade: "", saida: null, novo: false };
  for (let i = 0; i < argv.length; i++) {
    const v = argv[i];
    if (v === "--cidade") a.cidade = argv[++i] || "";
    else if (v === "--saida") a.saida = argv[++i];
    else if (v === "--novo") a.novo = true;
    else if (v === "-h" || v === "--help") a.ajuda = true;
    else if (!a.arquivo) a.arquivo = v;
  }
  return a;
}

const soDigitos = (s) => String(s || "").replace(/\D/g, "");
const pareceCNPJ = (s) => soDigitos(s).length === 14 && /^[\d.\/\s-]+$/.test(String(s).trim());

// Aceita .txt (uma empresa por linha: "Nome", "Nome; Cidade" ou CNPJ) ou .csv com cabeçalho
// (colunas reconhecidas: empresa/nome/razão social/fantasia, cidade/município, uf, cnpj).
function lerLista(caminho, cidadePadrao) {
  const linhas = readFileSync(caminho, "utf8").replace(/^﻿/, "").split(/\r?\n/).map((l) => l.trim()).filter((l) => l && !l.startsWith("#"));
  if (!linhas.length) return [];
  const sep = [";", "\t", ","].find((s) => linhas[0].includes(s)) || ";";
  const dividir = (l) => l.split(sep).map((c) => c.trim().replace(/^"(.*)"$/, "$1").trim());
  const cab = dividir(linhas[0]).map(semAcento);
  const col = (re) => cab.findIndex((c) => re.test(c));
  const iNome = col(/^(empresa|nome|razao|fantasia|cliente)/), iCidade = col(/^(cidade|municipio)/), iUF = col(/^(uf|estado)$/), iCNPJ = col(/cnpj/);

  if (iNome >= 0 || iCNPJ >= 0) {
    return linhas.slice(1).map((l) => {
      const c = dividir(l);
      const cidade = [c[iCidade], c[iUF]].filter(Boolean).join("/") || cidadePadrao;
      const cnpj = iCNPJ >= 0 && soDigitos(c[iCNPJ]).length === 14 ? soDigitos(c[iCNPJ]) : "";
      return { entrada: l, nome: iNome >= 0 ? c[iNome] : "", cidade, cnpj };
    }).filter((e) => e.nome || e.cnpj);
  }
  return linhas.map((l) => {
    if (pareceCNPJ(l)) return { entrada: l, nome: "", cidade: cidadePadrao, cnpj: soDigitos(l) };
    const [nome, cidade] = l.split(/\s*[;|\t]\s*/);
    return { entrada: l, nome, cidade: (cidade || cidadePadrao || "").trim(), cnpj: "" };
  });
}

const COLUNAS = ["Entrada", "Empresa", "Cidade", "WhatsApp principal", "Link WhatsApp", "Outros WhatsApp", "Celulares", "Fixos", "E-mails", "Site", "Instagram", "CNPJ", "Situação CNPJ", "Confiança", "Origem do principal", "Fontes", "Observações"];
const celula = (v) => {
  const s = String(v ?? "").replace(/\s+/g, " ").trim();
  return /[";\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
};
const juntar = (ts) => ts.map((t) => t.numero).join(" | ");

function linhaCSV(e, r) {
  const p = r.whatsapp?.[0];
  return [
    e.entrada, r.empresa, r.cidade,
    p?.numero || "", p ? `https://wa.me/55${p.d}` : "",
    juntar(r.whatsapp?.slice(1) || []), juntar(r.celulares || []), juntar(r.fixos || []),
    (r.emails || []).join(" | "), r.site, r.instagram ? `https://instagram.com/${r.instagram}` : "",
    r.cnpj, r.situacao, p?.confianca || "", p?.origens.join(", ") || "", (r.fontes || []).join(", "), (r.obs || []).join(" · "),
  ].map(celula).join(";");
}

// Lê a coluna "Entrada" do CSV já existente para pular o que já foi pesquisado (retomar de onde parou).
function jaFeitas(caminho) {
  if (!existsSync(caminho)) return new Set();
  const feitas = new Set();
  for (const l of readFileSync(caminho, "utf8").replace(/^﻿/, "").split(/\r?\n/).slice(1)) {
    if (!l) continue;
    const m = l.match(/^"((?:[^"]|"")*)"|^([^;]*)/);
    feitas.add((m[1] ?? m[2]).replace(/""/g, '"'));
  }
  return feitas;
}

// ------------------------------------------------------------ terminal
const cor = (c, s) => process.stdout.isTTY ? `\x1b[${c}m${s}\x1b[0m` : s;
const verde = (s) => cor(32, s), amarelo = (s) => cor(33, s), cinza = (s) => cor(90, s), negrito = (s) => cor(1, s), ciano = (s) => cor(36, s);

async function main() {
  const args = lerArgs(process.argv.slice(2));
  if (args.ajuda || !args.arquivo) {
    console.log(`
${negrito("Recon de Contatos")} — acha WhatsApp, celular e telefone de uma lista de empresas

  node recon.js <lista.txt | lista.csv> [opções]

  ${ciano("--cidade")} "Campinas/SP"   cidade padrão para quem não tiver cidade na lista
  ${ciano("--saida")}  arquivo.csv     nome do arquivo de resultado (padrão: resultados_<lista>.csv)
  ${ciano("--novo")}                   começa do zero (por padrão retoma de onde parou)

  Formato da lista (.txt), uma por linha:
    Casa do Construtor
    Ferragens Rocha; Campinas/SP
    12.345.678/0001-90
`);
    process.exit(args.ajuda ? 0 : 1);
  }
  if (!existsSync(args.arquivo)) { console.error(`Arquivo não encontrado: ${args.arquivo}`); process.exit(1); }

  const empresas = lerLista(args.arquivo, args.cidade);
  const saida = args.saida || `resultados_${basename(args.arquivo).replace(/\.[^.]+$/, "")}.csv`;
  if (args.novo || !existsSync(saida)) writeFileSync(saida, "﻿" + COLUNAS.join(";") + "\r\n", "utf8");
  const feitas = jaFeitas(saida);
  const pendentes = empresas.filter((e) => !feitas.has(e.entrada));

  console.log(`\n${negrito("Recon de Contatos")}  ${cinza(`${empresas.length} empresas na lista · ${empresas.length - pendentes.length} já feitas · ${pendentes.length} para pesquisar`)}`);
  console.log(cinza(`Resultado: ${saida}   (Ctrl+C para parar — rodando de novo, continua de onde parou)\n`));

  let comWhats = 0, comTel = 0, esperas = 0;
  const inicio = Date.now();
  for (let i = 0; i < pendentes.length; i++) {
    const e = pendentes[i];
    const pos = `[${String(empresas.length - pendentes.length + i + 1).padStart(String(empresas.length).length)}/${empresas.length}]`;
    process.stdout.write(`${cinza(pos)} ${negrito(e.nome || e.cnpj)}${e.cidade ? cinza(` · ${e.cidade}`) : ""} … `);
    let r;
    try {
      r = await pesquisarEmpresa(e);
      esperas = 0;
    } catch (err) {
      if (err.message === "BLOQUEIO" && esperas < 4) {
        esperas++;
        const s = 90 * esperas;
        console.log(amarelo(`\n   o buscador pediu uma pausa — retomando em ${s}s…`));
        await pausa(s * 1000);
        i--; // tenta a mesma empresa de novo
        continue;
      }
      r = { empresa: e.nome || e.cnpj, cidade: e.cidade, cnpj: e.cnpj, obs: [`Erro: ${err.message}`] };
    }
    appendFileSync(saida, linhaCSV(e, r) + "\r\n", "utf8");

    const p = r.whatsapp?.[0];
    const outros = [...(r.celulares || []), ...(r.fixos || [])];
    if (p) { comWhats++; comTel++; console.log(`${verde("✔ " + p.numero)} ${cinza(`(${p.tipo}, confiança ${p.confianca})`)}${outros.length ? cinza(` +${outros.length} tel`) : ""}`); }
    else if (outros.length) { comTel++; console.log(`${amarelo("☎ " + outros[0].numero)} ${cinza(`(${outros[0].tipo}, sem WhatsApp confirmado)`)}${outros.length > 1 ? cinza(` +${outros.length - 1}`) : ""}`); }
    else console.log(cinza(`— nada encontrado${r.obs?.length ? ` (${r.obs[0]})` : ""}`));
  }

  const min = ((Date.now() - inicio) / 60000).toFixed(1);
  console.log(`\n${negrito("Pronto!")} ${pendentes.length} pesquisadas em ${min} min · ${verde(`${comWhats} com WhatsApp/celular`)} · ${comTel} com algum telefone`);
  console.log(`Abra ${ciano(saida)} no Excel.\n`);
}

main().catch((e) => { console.error(e); process.exit(1); });
