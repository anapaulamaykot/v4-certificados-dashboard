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

function sealSVG(rank, size) {
  const tierColor = rank === 1 ? '#fb2e0a' : rank === 2 ? '#c74a34' : '#a8674f';
  const tierColorDeep = rank === 1 ? '#560303' : rank === 2 ? '#7a2418' : '#5c3527';
  const r = size / 2 - size * 0.06;
  const cx = size / 2, cy = size / 2;
  const path = sealPath(cx, cy, r, size * 0.045, 22);
  return `
    <svg width="${size}" height="${size}" viewBox="0 0 ${size} ${size}">
      <defs>
        <linearGradient id="sealGrad${rank}" x1="0" y1="0" x2="1" y2="1">
          <stop offset="0%" stop-color="${tierColor}" />
          <stop offset="100%" stop-color="${tierColorDeep}" />
        </linearGradient>
      </defs>
      <path d="${path}" fill="url(#sealGrad${rank})" />
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

// Constrói o ranking de unidades usando SÓ investidores ativos como universo:
// tanto o denominador (quantos ativos tem a unidade) quanto o numerador
// (quantos certificados eles geraram) vêm exclusivamente de quem está ativo hoje.
// activeInvestorsMap: Map(id_usuario -> { filial, cargo, nome })
function buildRankingUnidades(certRows, activeInvestorsMap) {
  activeInvestorsMap = activeInvestorsMap || new Map();
  const byFilial = new Map();

  const normFilial = (f) => {
    let filial = f && f.trim() ? f.trim() : 'Matriz / Sem Unidade';
    if (FILIAL_INVALIDA.has(filial.toLowerCase())) filial = 'Matriz / Sem Unidade';
    return filial;
  };

  // Primeiro registra todo mundo ativo na sua unidade (garante o denominador
  // certo mesmo pra quem ainda não tirou nenhum certificado)
  for (const [id, info] of activeInvestorsMap.entries()) {
    const filial = normFilial(info.filial);
    if (!byFilial.has(filial)) byFilial.set(filial, { certificados: 0, comCertificado: new Set(), ativos: new Set() });
    byFilial.get(filial).ativos.add(id);
  }

  // Depois soma só os certificados de gente que está nesse mapa de ativos
  for (const r of certRows) {
    const id = keyOf(r.id_usuario);
    if (!activeInvestorsMap.has(id)) continue;
    const filial = normFilial(activeInvestorsMap.get(id).filial);
    const entry = byFilial.get(filial);
    entry.certificados += 1;
    entry.comCertificado.add(id);
  }

  const list = [];
  for (const [filial, v] of byFilial.entries()) {
    const ativos = v.ativos.size;
    const media = ativos > 0 ? v.certificados / ativos : 0;
    const engajamento = ativos > 0 ? (v.comCertificado.size / ativos) * 100 : 0;
    list.push({
      nome: filial,
      certificados: v.certificados,
      comCertificado: v.comCertificado.size,
      ativos,
      temDadosAtivos: ativos > 0,
      media,
      engajamento,
      elegivelPodio: ativos >= MIN_ATIVOS_PODIO,
    });
  }
  list.sort((a, b) => b.media - a.media || b.certificados - a.certificados);
  return list;
}

// Constrói o ranking de investidores — só investidores ATIVOS entram,
// mesmo que ainda tenham zero certificados (aparecem no fim da lista).
// Cada investidor carrega a lista de nomes dos certificados que gerou.
function buildRankingInvestidores(certRows, activeInvestorsMap) {
  activeInvestorsMap = activeInvestorsMap || new Map();
  const byUser = new Map();

  for (const [id, info] of activeInvestorsMap.entries()) {
    byUser.set(keyOf(id), {
      nome: info.nome || id,
      filial: info.filial && info.filial.trim() ? info.filial.trim() : 'Matriz / Sem Unidade',
      cargo: info.cargo && info.cargo.trim() ? info.cargo.trim() : 'Não informado',
      email: info.email || '',
      certificados: 0,
      certificadosNomes: [],
    });
  }

  for (const r of certRows) {
    const id = keyOf(r.id_usuario);
    if (!byUser.has(id)) continue;
    const u = byUser.get(id);
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
