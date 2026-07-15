// Helpers compartilhados entre admin.html e index.html

function fmtInt(n) {
  return Number(n || 0).toLocaleString('pt-BR');
}

function fmtDec(n) {
  return Number(n || 0).toLocaleString('pt-BR', { maximumFractionDigits: 1, minimumFractionDigits: 1 });
}

// Converte serial de data do Excel (dias desde 1899-12-30) para 'YYYY-MM-DD'
function excelSerialToISODate(serial) {
  if (serial === null || serial === undefined || serial === '') return null;
  const n = Number(serial);
  if (Number.isNaN(n)) return null;
  const utcDays = Math.floor(n - 25569);
  const utcMs = utcDays * 86400 * 1000;
  const d = new Date(utcMs);
  if (Number.isNaN(d.getTime())) return null;
  return d.toISOString().slice(0, 10);
}

// Gera o path SVG de um "selo" (rosetta) — círculo com borda serrilhada,
// remete a um selo de certificado. r = raio médio, wobble = variação da onda.
function sealPath(cx, cy, r, wobble, points) {
  const pts = [];
  const step = (Math.PI * 2) / points;
  for (let i = 0; i < points; i++) {
    const angle = i * step;
    const rr = r + (i % 2 === 0 ? wobble : -wobble * 0.4);
    pts.push([cx + rr * Math.cos(angle), cy + rr * Math.sin(angle)]);
  }
  let d = `M ${pts[0][0]} ${pts[0][1]} `;
  for (let i = 1; i <= pts.length; i++) {
    const p = pts[i % pts.length];
    d += `L ${p[0]} ${p[1]} `;
  }
  return d + 'Z';
}

function sealSVG(rank, size, uid) {
  const ns = uid || 'x';
  const tierColor = rank === 1 ? '#fb2e0a' : rank === 2 ? '#c74a34' : '#a8674f';
  const tierColorDeep = rank === 1 ? '#560303' : rank === 2 ? '#7a2418' : '#5c3527';
  const r = size / 2 - size * 0.06;
  const cx = size / 2, cy = size / 2;
  const path = sealPath(cx, cy, r, size * 0.045, 22);
  const gradId = `sealGrad-${ns}-${rank}`;
  return `
    <svg width="${size}" height="${size}" viewBox="0 0 ${size} ${size}">
      <defs>
        <linearGradient id="${gradId}" x1="0" y1="0" x2="1" y2="1">
          <stop offset="0%" stop-color="${tierColor}" />
          <stop offset="100%" stop-color="${tierColorDeep}" />
        </linearGradient>
      </defs>
      <path d="${path}" fill="url(#${gradId})" />
      <circle cx="${cx}" cy="${cy}" r="${r * 0.72}" fill="none" stroke="rgba(255,255,255,0.55)" stroke-width="1.5" stroke-dasharray="2 3" />
    </svg>
  `;
}

// Unidades com valor de Filial inválido/placeholder — tratadas como Matriz/Sem Unidade
const FILIAL_INVALIDA = new Set(['não tem', 'nao tem', 'n/a', 'na', '-', '31']);

// Mínimo de investidores ATIVOS pra uma unidade concorrer ao pódio (evita que 1 pessoa
// muito ativa domine o ranking proporcional por pura amostra pequena).
const MIN_ATIVOS_PODIO = 3;

// Normaliza um ID pra string, garantindo que o cruzamento entre bases
// (uma pode vir com ID numérico do Excel, outra como texto do CSV) sempre bate.
function keyOf(id) {
  return id === null || id === undefined ? '' : String(id).trim();
}

// O Supabase limita cada consulta a 1000 linhas por padrão — com milhares
// de certificados/investidores isso truncaria os dados silenciosamente.
// Essa função pagina automaticamente até trazer tudo.
async function fetchAllRows(queryBuilderFn, pageSize) {
  pageSize = pageSize || 1000;
  let all = [];
  let from = 0;
  while (true) {
    const { data, error } = await queryBuilderFn().range(from, from + pageSize - 1);
    if (error) throw error;
    all = all.concat(data || []);
    if (!data || data.length < pageSize) break;
    from += pageSize;
  }
  return all;
}

// --- Normalização de nomes de unidade -----------------------------------
// Junta variações de grafia da mesma unidade (V4_Lab / V4 Lab, Kuri & Co /
// Kuri&Co - Franchise Store, etc.) e só aceita como unidade válida quem
// tiver "V4" em pelo menos uma das grafias encontradas.

function stripAccents(s) {
  return s.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

const FILIAL_STOPWORDS = new Set([
  'v4', 'company', 'associados', 'associado', 'assoc', 'co', 'ltda', 'franchise', 'store', 'e',
]);

// Reduz um nome de unidade ao seu "núcleo" comparável, ignorando prefixos/
// sufixos corporativos, pontuação, acentos e maiúsculas/minúsculas.
function normalizeFilialKey(raw) {
  if (!raw) return '';
  let s = stripAccents(String(raw).toLowerCase());
  s = s.replace(/&/g, ' e ');
  s = s.replace(/[_\-.,]/g, ' ');
  s = s.replace(/[^a-z0-9\s]/g, ' ');
  const words = s.split(/\s+/).filter(w => w && !FILIAL_STOPWORDS.has(w));
  return words.join(' ').trim();
}

// Unidades legítimas que não têm "V4" no nome, mas são reais mesmo assim —
// adicionadas manualmente como exceção à regra.
const FILIAL_EXCECOES_SEM_V4 = [
  'ROSOLEM VERONEZE & CORDEIRO ASSESSORIA DE MARKETING LTDA',
];
const FILIAL_EXCECOES_KEYS = new Set(FILIAL_EXCECOES_SEM_V4.map(normalizeFilialKey));

// A partir de todas as grafias brutas encontradas, monta o mapa
// núcleo -> { canonical, hasV4 }. Só vira unidade válida quem tiver "v4"
// em pelo menos uma grafia do grupo (ou estiver na lista de exceções).
function buildFilialCanonicalMap(rawNames) {
  const groups = new Map();
  rawNames.forEach(raw => {
    const key = normalizeFilialKey(raw);
    if (!key) return;
    if (!groups.has(key)) groups.set(key, new Map());
    const variants = groups.get(key);
    variants.set(raw, (variants.get(raw) || 0) + 1);
  });

  const canonicalByKey = new Map();
  groups.forEach((variants, key) => {
    const entries = Array.from(variants.entries());
    const v4Entries = entries.filter(([v]) => /v4/i.test(v));
    const hasV4 = v4Entries.length > 0 || FILIAL_EXCECOES_KEYS.has(key);
    const pool = v4Entries.length ? v4Entries : entries;
    pool.sort((a, b) => b[1] - a[1] || b[0].length - a[0].length);
    canonicalByKey.set(key, { canonical: pool[0][0], hasV4 });
  });
  return canonicalByKey;
}

// Resolve o nome canônico de uma unidade, ou null se o grupo inteiro
// nunca teve "V4" em nenhuma grafia (nesse caso, cai em Matriz/Sem Unidade).
function resolveCanonicalFilial(raw, canonicalByKey) {
  const key = normalizeFilialKey(raw);
  const entry = canonicalByKey.get(key);
  if (entry && entry.hasV4) return entry.canonical;
  return null;
}

// Constrói um índice de busca de investidores ativos por ID e por e-mail
// (e-mail funciona como respaldo quando o ID não bate entre as duas bases).
// Também normaliza o nome da unidade de cada pessoa (grafias variantes
// viram um único nome canônico; unidades sem "V4" em nenhuma grafia caem
// em "Matriz / Sem Unidade").
// Retorna { lookup, roster } — lookup resolve qualquer linha de certificado
// à pessoa ativa correta; roster é a lista canônica de ativos (1 por pessoa).
function buildActiveLookup(ativosRows) {
  const rows = ativosRows || [];
  const canonicalByKey = buildFilialCanonicalMap(rows.map(r => r.filial));

  const lookup = new Map();
  const roster = new Map(); // id canônico -> info
  rows.forEach(r => {
    const id = keyOf(r.id_usuario);
    const email = (r.email || '').toLowerCase().trim();
    const canonicalId = id || ('email:' + email);
    const filial = resolveCanonicalFilial(r.filial, canonicalByKey) || 'Matriz / Sem Unidade';
    const info = { id: canonicalId, nome: r.nome, filial, cargo: r.cargo, email: r.email };
    if (id) lookup.set('id:' + id, info);
    if (email) lookup.set('email:' + email, info);
    roster.set(canonicalId, info);
  });
  return { lookup, roster };
}

// Resolve a qual investidor ativo uma linha de certificado pertence,
// tentando primeiro por ID e depois por e-mail.
function resolveInvestor(row, lookup) {
  const id = keyOf(row.id_usuario);
  if (id && lookup.has('id:' + id)) return lookup.get('id:' + id);
  const email = (row.email || '').toLowerCase().trim();
  if (email && lookup.has('email:' + email)) return lookup.get('email:' + email);
  return null;
}

// Constrói o ranking de unidades usando SÓ investidores ativos como universo:
// tanto o denominador (quantos ativos tem a unidade) quanto o numerador
// (quantos certificados eles geraram) vêm exclusivamente de quem está ativo hoje.
// Cada linha do arquivo de certificados vale 1, independente de ter ID preenchido.
// Classifica a unidade por porte (nº de investidores ativos), pra competir
// só dentro da própria faixa — não faz sentido uma unidade de 5 pessoas
// disputar com uma de 300.
function getTamanhoTier(ativos) {
  if (ativos <= 50) return { id: 'pequena', label: 'Até 50 pessoas' };
  if (ativos <= 100) return { id: 'media', label: '50 a 100 pessoas' };
  return { id: 'grande', label: 'Mais de 100 pessoas' };
}

function buildRankingUnidades(certRows, activeIndex) {
  const { lookup, roster } = activeIndex || { lookup: new Map(), roster: new Map() };
  const byFilial = new Map();

  const normFilial = (f) => {
    let filial = f && f.trim() ? f.trim() : 'Matriz / Sem Unidade';
    if (FILIAL_INVALIDA.has(filial.toLowerCase())) filial = 'Matriz / Sem Unidade';
    return filial;
  };

  for (const [id, info] of roster.entries()) {
    const filial = normFilial(info.filial);
    if (!byFilial.has(filial)) byFilial.set(filial, { certificados: 0, comCertificado: new Map(), ativos: new Set() });
    byFilial.get(filial).ativos.add(id);
  }

  for (const r of certRows) {
    const investor = resolveInvestor(r, lookup);
    if (!investor) continue;
    const filial = normFilial(investor.filial);
    const entry = byFilial.get(filial);
    entry.certificados += 1;
    entry.comCertificado.set(investor.id, { nome: investor.nome, email: investor.email });
  }

  const list = [];
  for (const [filial, v] of byFilial.entries()) {
    const ativos = v.ativos.size;
    const media = ativos > 0 ? v.certificados / ativos : 0;
    const engajamento = ativos > 0 ? (v.comCertificado.size / ativos) * 100 : 0;
    const pessoasCertificadas = Array.from(v.comCertificado.values())
      .sort((a, b) => a.nome.localeCompare(b.nome))
      .map(p => `${p.nome} — ${p.email}`);
    list.push({
      nome: filial,
      certificados: v.certificados,
      comCertificado: v.comCertificado.size,
      ativos,
      temDadosAtivos: ativos > 0,
      media,
      engajamento,
      elegivelPodio: ativos >= MIN_ATIVOS_PODIO,
      pessoasCertificadas,
      tamanho: getTamanhoTier(ativos),
    });
  }
  list.sort((a, b) => b.media - a.media || b.engajamento - a.engajamento);
  return list;
}

// Constrói o ranking de investidores — só investidores ATIVOS entram,
// mesmo que ainda tenham zero certificados (aparecem no fim da lista).
// Cada investidor carrega a lista de nomes dos certificados que gerou.
function buildRankingInvestidores(certRows, activeIndex) {
  const { lookup, roster } = activeIndex || { lookup: new Map(), roster: new Map() };
  const byUser = new Map();

  for (const [id, info] of roster.entries()) {
    byUser.set(id, {
      nome: info.nome || id,
      filial: info.filial && info.filial.trim() ? info.filial.trim() : 'Matriz / Sem Unidade',
      cargo: info.cargo && info.cargo.trim() ? info.cargo.trim() : 'Não informado',
      email: info.email || '',
      certificados: 0,
      certificadosNomes: [],
    });
  }

  for (const r of certRows) {
    const investor = resolveInvestor(r, lookup);
    if (!investor) continue;
    const u = byUser.get(investor.id);
    u.certificados += 1;
    if (r.conteudo) u.certificadosNomes.push(r.conteudo);
  }

  const list = Array.from(byUser.values());
  list.sort((a, b) => b.certificados - a.certificados);
  return list;
}

// Decodifica entidades HTML (&amp; &#39; etc.) — o export de usuários do Growth Learning
// vem com nomes escapados, e &amp; contém um ';' que quebraria o parse do CSV.
function decodeHtmlEntities(str) {
  const el = document.createElement('textarea');
  el.innerHTML = str;
  return el.value;
}

// Parser simples de CSV delimitado por ';' (sem aspas/quoting no arquivo de origem).
// Decodifica entidades HTML primeiro pra remover ';' escondidos em nomes.
function parseSemicolonCSV(text) {
  const clean = decodeHtmlEntities(text);
  const lines = clean.split(/\r?\n/).filter(l => l.trim().length > 0);
  const headers = lines[0].split(';').map(h => h.trim());
  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    const cells = lines[i].split(';');
    const obj = {};
    headers.forEach((h, idx) => { obj[h] = cells[idx] !== undefined ? cells[idx].trim() : ''; });
    rows.push(obj);
  }
  return rows;
}
