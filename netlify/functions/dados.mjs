/**
 * HEMU • Consulta da planilha Excel pública do Google Drive.
 * Implantação: Netlify Functions (sem npm, sem package.json).
 * Endpoint: /.netlify/functions/dados
 * A função extrai somente as quatro abas necessárias do arquivo XLSX.
 * IMPORTANTE: a planilha de origem deve ter acesso de leitura por link.
 */
import { inflateRawSync } from 'node:zlib';
import { createHash } from 'node:crypto';

const DRIVE_FILE_ID = '1QQJNc1gmoy3UxI2t9IKv1bD92ud0dfow';
const CACHE_MS = 10_000;
let cache = null;
let cacheEm = 0;
let cacheHash = null;

function entradaZip(arquivo, nome) {
  const b = Buffer.isBuffer(arquivo) ? arquivo : Buffer.from(arquivo);
  let fim = -1;
  for (let i = b.length - 22; i >= Math.max(0, b.length - 65557); i--) {
    if (b.readUInt32LE(i) === 0x06054b50) { fim = i; break; }
  }
  if (fim < 0) throw new Error('A origem não retornou um arquivo XLSX válido (ZIP).');
  const total = b.readUInt16LE(fim + 10);
  let pos = b.readUInt32LE(fim + 16);
  for (let i = 0; i < total; i++) {
    if (b.readUInt32LE(pos) !== 0x02014b50) throw new Error('Índice ZIP danificado.');
    const metodo = b.readUInt16LE(pos + 10);
    const comp = b.readUInt32LE(pos + 20);
    const nomeTam = b.readUInt16LE(pos + 28);
    const extraTam = b.readUInt16LE(pos + 30);
    const comentarioTam = b.readUInt16LE(pos + 32);
    const offset = b.readUInt32LE(pos + 42);
    const arquivoNome = b.toString('utf8', pos + 46, pos + 46 + nomeTam);
    pos += 46 + nomeTam + extraTam + comentarioTam;
    if (arquivoNome !== nome) continue;
    if (b.readUInt32LE(offset) !== 0x04034b50) throw new Error('Entrada ZIP danificada.');
    const inicio = offset + 30 + b.readUInt16LE(offset + 26) + b.readUInt16LE(offset + 28);
    const dados = b.subarray(inicio, inicio + comp);
    if (metodo === 0) return dados.toString('utf8');
    if (metodo === 8) return inflateRawSync(dados).toString('utf8');
    throw new Error('Método de compressão do XLSX não suportado.');
  }
  throw new Error(`O XLSX não contém ${nome}.`);
}

function desescapar(x) {
  return String(x ?? '').replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (_m, k) => {
    if (k.startsWith('#x')) return String.fromCodePoint(parseInt(k.slice(2), 16));
    if (k.startsWith('#')) return String.fromCodePoint(parseInt(k.slice(1), 10));
    return {amp:'&',lt:'<',gt:'>',quot:'"',apos:"'"}[k.toLowerCase()] ?? _m;
  });
}
function textoXML(x) { return desescapar([...String(x).matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map(m => m[1]).join('')); }
function atr(x, nome) { return x.match(new RegExp(`(?:^|\\s)${nome}="([^"]*)"`))?.[1] ?? ''; }
function planilhaRows(xml, shared, limite = 200) {
  const out = [];
  const re = /<row\b([^>]*)>([\s\S]*?)<\/row>/g;
  let r;
  while ((r = re.exec(xml))) {
    const numero = Number(atr(r[1], 'r'));
    if (numero > limite) break;
    const celulas = {};
    for (const c of r[2].matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const ref = atr(c[1], 'r');
      const col = ref.match(/^[A-Z]+/)?.[0];
      if (!col) continue;
      const tipo = atr(c[1], 't');
      const interior = c[2] || '';
      let v = interior.match(/<v(?:\s[^>]*)?>([\s\S]*?)<\/v>/)?.[1];
      if (tipo === 's') v = shared[Number(v)] ?? '';
      else if (tipo === 'inlineStr') v = textoXML(interior);
      else if (v === undefined && interior.includes('<is')) v = textoXML(interior);
      else v = desescapar(v ?? '');
      if (v !== '') celulas[col] = v;
    }
    out.push({numero, c:celulas});
  }
  return out;
}
function normalizarStatus(x) { return String(x ?? '').trim() || 'Não informado'; }
function numerico(x) {
  const s = String(x ?? '').trim();
  return /^\d+(?:\.0+)?$/.test(s) ? Number(s) : null;
}
function dataExcel(x) {
  const s = String(x ?? '').trim();
  if (!s) return '';
  const n = Number(s);
  if (Number.isFinite(n) && n > 20000 && n < 100000) {
    return new Date(Date.UTC(1899, 11, 30) + Math.floor(n) * 86400000).toISOString().slice(0, 10);
  }
  const pt = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (pt) return `${pt[3]}-${pt[2].padStart(2,'0')}-${pt[1].padStart(2,'0')}`;
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}
function obterAbas(arquivo) {
  const workbook = entradaZip(arquivo, 'xl/workbook.xml');
  const rels = entradaZip(arquivo, 'xl/_rels/workbook.xml.rels');
  const ids = {};
  for (const m of rels.matchAll(/<Relationship\b([^>]*?)\/>/g)) {
    ids[atr(m[1], 'Id')] = atr(m[1], 'Target');
  }
  const abas = {};
  for (const m of workbook.matchAll(/<sheet\b([^>]*?)\/>/g)) {
    const nome = desescapar(atr(m[1], 'name'));
    const rid = atr(m[1], 'r:id');
    const target = ids[rid];
    if (target) abas[nome] = target.replace(/^\//, '').replace(/^xl\//, 'xl/').replace(/^worksheets\//, 'xl/worksheets/');
  }
  return abas;
}

export function parseWorkbook(buffer) {
  const arquivo = Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer);
  const sharedXml = entradaZip(arquivo, 'xl/sharedStrings.xml');
  const shared = [...sharedXml.matchAll(/<si(?:\s[^>]*)?>([\s\S]*?)<\/si>/g)].map(m => textoXML(m[1]));
  const abas = obterAbas(arquivo);
  function ler(nome, limite) {
    if (!abas[nome]) throw new Error(`A aba “${nome}” não foi encontrada no Excel.`);
    return planilhaRows(entradaZip(arquivo, abas[nome]), shared, limite);
  }

  const campos = ['Mapa','CCF','Matriz de Risco','Matriz de Registro','LEG','LMDE','POP','PRS','Indicadores'];
  const colunas = ['E','F','G','H','I','J','K','L','M'];
  const Controle_Setor = [];
  const Controle_Ferramentas = [];
  for (const {c} of ler('CONTROLE', 250)) {
    const id = numerico(c.B);
    const setor = String(c.C ?? '').trim();
    if (!id || !setor) continue;
    const info = { ID_Setor: id, Setor: setor, Sigla: String(c.D ?? '').trim() };
    let aplicaveis = 0, concluidos = 0;
    campos.forEach((campo, i) => {
      const status = normalizarStatus(c[colunas[i]]);
      info[campo] = status;
      Controle_Ferramentas.push({ID_Setor:id,Setor:setor,Sigla:info.Sigla,Ferramenta:campo,Status:status});
      if (status !== 'N.A' && status !== 'Não informado') {
        aplicaveis++;
        if (status === 'Concluído') concluidos++;
      }
    });
    info.Percentual_Controle = aplicaveis ? concluidos / aplicaveis : 0;
    Controle_Setor.push(info);
  }

  const Cronograma_Mapa_CCF = [];
  for (const {c} of ler('CRONOGRAMA MAPA E CCF', 250)) {
    const id = numerico(c.B), setor = String(c.C ?? '').trim();
    if (!id || !setor) continue;
    const origemData = String(c.E ?? '').trim();
    Cronograma_Mapa_CCF.push({
      ID_Setor:id, Setor:setor, Sigla:String(c.D??'').trim(),
      Data_Prevista:dataExcel(origemData), Data_Texto_Origem:origemData,
      Horario:String(c.F??'').trim(), Responsavel:String(c.H??'').trim(),
      Status:normalizarStatus(c.G)
    });
  }

  const Planejamento_SGQ = [];
  let modulo = '';
  for (const {numero,c} of ler('PLANEJAMENTO SGQ', 150)) {
    const b = String(c.B??'').trim();
    if (/^MÓDULO\s+\d+/i.test(b)) modulo = b;
    const id = numerico(b);
    const etapa = String(c.C??'').trim();
    if (!id || !etapa) continue;
    Planejamento_SGQ.push({
      Projeto:'SGQ',Modulo:modulo,ID_Etapa:id,Etapa:etapa,
      Responsavel:String(c.D??'').trim(),Status:normalizarStatus(c.E),
      Observacoes:String(c.F??'').trim(),Periodo_Estimado:String(c.G??'').trim(),
      Linha_Origem:numero
    });
  }

  const Politicas = [];
  for (const {c} of ler('POLÍTICAS', 100)) {
    const id = numerico(c.B), politica=String(c.C??'').trim();
    if (!id || !politica) continue;
    Politicas.push({
      ID_Politica:id,Politica:politica,
      Setor_Responsavel:String(c.D??'').trim(),Status:normalizarStatus(c.E)
    });
  }
  if (!Controle_Setor.length || !Cronograma_Mapa_CCF.length || !Planejamento_SGQ.length || !Politicas.length) {
    throw new Error('A leitura das abas retornou registros incompletos. Confira a estrutura do Excel.');
  }
  return {Controle_Setor,Controle_Ferramentas,Cronograma_Mapa_CCF,Planejamento_SGQ,Politicas};
}

async function baixarExcel() {
  const id = process.env.DRIVE_FILE_ID || DRIVE_FILE_ID;
  const versao = Date.now();
  const urls = process.env.EXCEL_DOWNLOAD_URL ? [process.env.EXCEL_DOWNLOAD_URL] : [
    `https://drive.usercontent.google.com/download?id=${encodeURIComponent(id)}&export=download&confirm=t&_=${versao}`,
    `https://drive.google.com/uc?export=download&id=${encodeURIComponent(id)}&confirm=t&_=${versao}`,
    `https://docs.google.com/spreadsheets/d/${encodeURIComponent(id)}/export?format=xlsx&_=${versao}`
  ];
  const erros = [];
  for (const url of urls) {
    try {
      const r = await fetch(url, {signal:AbortSignal.timeout(15000),redirect:'follow',cache:'no-store',headers:{'User-Agent':'Mozilla/5.0 HEMU-Qualidade-SGQ/1.0','Cache-Control':'no-cache','Pragma':'no-cache'}});
      if (!r.ok) { erros.push(`Origem respondeu HTTP ${r.status}`); continue; }
      const bytes = Buffer.from(await r.arrayBuffer());
      if (bytes.length > 30_000_000) throw new Error('Arquivo excede o limite de segurança de 30 MB.');
      if (bytes.length < 4 || bytes.readUInt32LE(0) !== 0x04034b50) {
        erros.push('A origem retornou uma página HTML/autorização em vez do XLSX');
        continue;
      }
      return bytes;
    } catch(e) { erros.push(e.message); }
  }
  throw new Error('Não foi possível baixar o Excel do Google Drive. ' + erros.slice(0,3).join(' | '));
}

export default async function handler(request) {
  if (request.method !== 'GET') return new Response(JSON.stringify({error:'Método não permitido.'}), {status:405,headers:{'Content-Type':'application/json; charset=utf-8'}});
  try {
    if (!cache || Date.now() - cacheEm > CACHE_MS) {
      const arquivo = await baixarExcel();
      const novos = parseWorkbook(arquivo);
      cache = novos;
      cacheHash = createHash('sha256').update(JSON.stringify(novos)).digest('hex');
      cacheEm = Date.now();
    }
    return new Response(JSON.stringify({ok:true, origem:'Google Drive - Excel XLSX', sincronizadoEm:new Date(cacheEm).toISOString(), fonteHash:cacheHash, totais:{setores:cache.Controle_Setor.length,ferramentas:cache.Controle_Ferramentas.length,cronograma:cache.Cronograma_Mapa_CCF.length,sgq:cache.Planejamento_SGQ.length,politicas:cache.Politicas.length}, data:cache}), {
      status:200,headers:{'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store'}
    });
  } catch(e) {
    console.error('Erro ao carregar planilha:',e);
    return new Response(JSON.stringify({ok:false,error:e.message}),{
      status:502,headers:{'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store'}
    });
  }
}