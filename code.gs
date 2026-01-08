/************************
 * CONFIGURAÇÕES
 ************************/
const PLANILHA_ID    = ''; o id da planilha
const ID_TEMPLATE    = ''; id do word para que saia em pdf preenchido
const FOLDER_PDFS_ID = ''; id da pasta onde é armazenado os pdfs
const TZ             = 'America/Sao_Paulo';

const RESPOSTAS_SHEET_NAME = 'Respostas';
const CPF_HEADER           = 'CPF';

// Cabeçalhos compartilhados
const BAIXAS_SHEET_NAME = 'Baixas';
const ENTREGUE_HEADER   = 'Entregue';
const ENTREGUE_EM_HDR   = 'Entregue em';
const ENTREGUE_POR_HDR  = 'Entregue por';
const ENTREGUE_UNID_HDR = 'Unidade de Entrega';
const ENTREGUE_OBS_HDR  = 'Obs Entrega';
const PROTOCOLO_HDR     = 'Protocolo';

// E-mail
const OUTBOX_SHEET_NAME = 'FilaEmails';
const OUTBOX_MAX_TRIES  = 5;
const COPIA_EMAIL       = ''; // opcional

/************************
 * ROUTER / VIEWS (AJUSTADO)
 ************************/
function doGet(e) {
  try { ensureOutboxSheet_(); ensureQueueTrigger_(); } catch (err) { Logger.log('[BOOT] ' + (err.message || err)); }

  const viewRaw = (e && e.parameter && e.parameter.view) ? String(e.parameter.view) : 'login';
  const view = (viewRaw || '').trim().toLowerCase();

  const map = {
    login:'login',
    hub:'hub',
    formulario:'formulario',
    baixa:'baixa',
    central:'central',
    admin:'admin',
    analista:'analista',
    ti:'ti'
  };

  const file = map[view] || 'login';

  try {
    // ✅ serve como TEMPLATE (evaluate) — resolve include(), variáveis, etc.
    return renderHtmlFile_(file, { view })
      .setTitle('Sistema SEMFAS')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (err) {
    return HtmlService
      .createHtmlOutput(errorHtml_('View não encontrada: ' + file + '.html', err, true))
      .setTitle('Sistema SEMFAS')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * Renderiza arquivo HTML como Template (com evaluate).
 * Use isso SEMPRE quando seu HTML tiver <?!= include() ?> ou variáveis.
 */
function renderHtmlFile_(file, data){
  const t = HtmlService.createTemplateFromFile(file);
  t.view = (data && data.view) ? data.view : '';
  t.ctx  = (data && data.ctx)  ? data.ctx  : {};
  return t.evaluate();
}

/**
 * Retorna HTML em STRING (para quando você usa google.script.run e troca a tela no front).
 * Também usando TEMPLATE (evaluate) pra não dar null/erro.
 */
function renderView(view, ctx){
  const map = {
    login:'login',
    hub:'hub',
    formulario:'formulario',
    baixa:'baixa',
    central:'central',
    admin:'admin',
    analista:'analista',
    ti:'ti'
  };
  const file = map[(view||'').toString().trim().toLowerCase()] || 'login';

  try{
    return renderHtmlFile_(file, { view:file, ctx: ctx || {} }).getContent();
  } catch (err){
    return errorHtml_('Erro ao renderizar: ' + file + '.html', err, false);
  }
}

function include(file){
  // usado dentro do HTML: <?!= include('arquivo') ?>
  return HtmlService.createHtmlOutputFromFile(file).getContent();
}

/** ✅ Função genérica para o front pedir qualquer tela */
function getTelaHtml(view, ctx){
  return renderView(view, ctx);
}

/** Compatíveis com seu front atual */
function getFormularioHtml(){ return renderView('formulario'); }
function getAdminHtml()     { return renderView('admin'); }
function getCentralHtml()   { return renderView('central'); }
function getAnalistaHtml()  { return renderView('analista'); }
function getHubHtml()       { return renderView('hub'); }
function getBaixaHtml()     { return renderView('baixa'); }
function getTiHtml()        { return renderView('ti'); }

/** Página de erro (pra nunca ficar “tela branca”) */
function errorHtml_(title, err, autoBackToLogin){
  const msg = (err && (err.stack || err.message)) ? (err.stack || err.message) : String(err);
  return `
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>${escapeHtml_(title)}</title>
  <style>
    body{font-family:Arial,system-ui;padding:18px;background:#f8fafc;color:#0f172a}
    .card{max-width:980px;margin:0 auto;background:#fff;border:1px solid #e2e8f0;border-radius:14px;padding:16px}
    pre{white-space:pre-wrap;background:#0b1220;color:#e5e7eb;padding:12px;border-radius:12px;overflow:auto}
    a{color:#2563eb;font-weight:700}
  </style>
</head>
<body>
  <div class="card">
    <h2>${escapeHtml_(title)}</h2>
    <p>Copie o erro abaixo e me mande aqui.</p>
    <pre>${escapeHtml_(msg)}</pre>
    ${autoBackToLogin ? `<p><a href="?view=login">Voltar ao login</a></p>` : ``}
  </div>
  ${autoBackToLogin ? `<script>setTimeout(()=>location.search='?view=login', 1500)</script>` : ``}
</body>
</html>`;
}

function escapeHtml_(s){
  return (s ?? '').toString()
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;')
    .replace(/'/g,'&#039;');
}

/************************
 * LOGIN
 ************************/
function canonicalSectorKey_(s){
  return (s || '').toString()
    .replace(/\u00A0/g,' ')
    .replace(/[\u200B-\u200D\uFEFF]/g,'')
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/\s+/g,' ')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g,'');
}

function cleanSectorLabel_(s){
  return (s || '').toString()
    .replace(/\u00A0/g,' ')
    .replace(/[\u200B-\u200D\uFEFF]/g,'')
    .replace(/\s+/g,' ')
    .trim();
}

function listarSetores() {
  const ss  = SpreadsheetApp.openById(PLANILHA_ID);
  const sh  = ss.getSheetByName('Login');
  if (!sh) throw new Error('Aba "Login" não encontrada.');

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) throw new Error('Aba "Login" sem dados.');

  const headers = sh.getRange(1,1,1,lastCol).getValues()[0];
  let colIdx = findColIdx_(headers, 'Setor','Unidade','Setor/Unidade','Setor (Unidade)');
  if (colIdx < 0) {
    const row2 = sh.getRange(2,1,1,lastCol).getValues()[0];
    colIdx = row2.findIndex(v => String(v||'').trim() !== '');
    if (colIdx < 0) colIdx = 0;
  }

  const valores = sh.getRange(2, colIdx+1, lastRow-1, 1).getValues().map(r=>r[0]);
  const mapa = new Map();
  valores.forEach(v => {
    const label = cleanSectorLabel_(v);
    if (!label) return;
    mapa.set(canonicalSectorKey_(label), label);
  });

  const setores = Array.from(mapa.values()).sort((a,b)=>a.localeCompare(b,'pt-BR',{sensitivity:'base'}));
  if (!setores.length) throw new Error('Nenhum setor encontrado na aba "Login".');
  return setores;
}

function isHubSector_(label){
  const s = cleanSectorLabel_(label).toUpperCase();
  return s.includes('CRAS') || s.includes('CREAS') || s.includes('CRAM') ||
         s.includes('SEDE') || s.includes('CENTROPOP') || s.includes('CENTRO POP') || s.includes('CENTRO-POP');
}

function getLoginIdx_(headers){
  return {
    setor: findColIdx_(headers, 'Setor','Unidade','Setor/Unidade'),
    user:  findColIdx_(headers, 'Usuário','Usuario','User','Login'),
    pass:  findColIdx_(headers, 'Senha','Password'),
    email: findColIdx_(headers, 'E-mail','Email'),
    role:  findColIdx_(headers, 'Perfil','Role','Permissão','Permissao')
  };
}

/** ✅ (FALTAVA) Mapeia perfil -> view direta */
function getDirectViewForRole_(role){
  const r = String(role || 'usuario').trim().toLowerCase();
  if (!r) return 'formulario';

  if (r.includes('admin'))    return 'admin';
  if (r.includes('central'))  return 'central';
  if (r.includes('analista')) return 'analista';
  if (r === 'ti' || r.includes('ti')) return 'ti';
  if (r.includes('baixa'))    return 'baixa';

  // padrão: usuário comum cai no formulário/hub
  return 'formulario';
}

function verificarLogin(setor, usuario, senha) {
  const aba = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName('Login');
  if (!aba) return { ok:false };

  const lastRow = aba.getLastRow();
  const lastCol = aba.getLastColumn();
  if (lastRow < 2) return { ok:false };

  const headers = aba.getRange(1,1,1,lastCol).getValues()[0] || [];
  const idx = getLoginIdx_(headers);

  // fallback se não achar cabeçalho
  if (idx.setor < 0) idx.setor = 0;
  if (idx.user  < 0) idx.user  = 1;
  if (idx.pass  < 0) idx.pass  = 2;
  if (idx.role  < 0) idx.role  = 4;

  const dados = aba.getRange(2,1,lastRow-1,lastCol).getValues();

  const setorKey = canonicalSectorKey_(setor);
  const userKey  = canonicalSectorKey_(usuario);
  const pass     = String(senha || '').trim();

  for (let i=0; i<dados.length; i++) {
    const row  = dados[i];
    const set  = cleanSectorLabel_(row[idx.setor] || '');
    const usr  = String(row[idx.user] || '');
    const pwd  = String(row[idx.pass] || '');
    const role = String(row[idx.role] || 'usuario');

    if (canonicalSectorKey_(set) === setorKey &&
        canonicalSectorKey_(usr) === userKey &&
        pwd.trim() === pass) {
      return { ok:true, role:(role.trim() || 'usuario'), setorLabel:set };
    }
  }
  return { ok:false };
}

function autenticarERetornarTela(setor, usuario, senha){
  const res = verificarLogin(setor, usuario, senha);
  if (!res || !res.ok) return { ok:false };

  try {
    PropertiesService.getUserProperties()
      .setProperties({ semfas_setor:setor, semfas_usuario:usuario }, true);
  } catch (_){}

  const direct = getDirectViewForRole_(res.role);
  const view = (direct !== 'formulario')
    ? direct
    : (isHubSector_(res.setorLabel || setor) ? 'hub' : 'formulario');

  return {
    ok: true,
    view,
    role: (res.role || 'usuario'),
    setorLabel: (res.setorLabel || setor),
    forceRedirect: true,
    redirectUrl: '?view=' + encodeURIComponent(view)
  };
}

/************************
 * HELPERs
 ************************/
function normText_(s){
  return String(s||'')
    .normalize('NFD').replace(/[\u00B7\u0300-\u036f]/g,'')
    .replace(/\s+/g,' ')
    .trim()
    .toLowerCase();
}

function withRetry(fn, tentativas=5, baseMs=250){
  for (let i=0;i<tentativas;i++){
    try { return fn(); }
    catch(e){
      if (i===tentativas-1) throw e;
      Utilities.sleep(baseMs * Math.pow(2,i));
    }
  }
}

function parseAnyDate(s){
  if (!s) return null;
  if (s instanceof Date) return s;

  if (typeof s === 'number'){
    if (s > 10*365*24*3600*1000) return new Date(s); // epoch ms
    return new Date(Math.round((s - 25569) * 86400 * 1000)); // serial Sheets
  }

  if (typeof s !== 'string') return null;

  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(+m[1], +m[2]-1, +m[3]);

  m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) return new Date(+m[3], +m[2]-1, +m[1]);

  const d = new Date(s);
  return isNaN(d) ? null : d;
}

function coerceSheetDate(v){
  if (!v && v!==0) return null;
  if (v instanceof Date) return startOfDay(v);
  if (typeof v === 'number')
    return startOfDay(new Date(Math.round((v - 25569) * 86400 * 1000)));
  if (typeof v === 'string'){
    const d = parseAnyDate(v);
    return d ? startOfDay(d) : null;
  }
  return null;
}

function startOfDay(d){
  const x=new Date(d);
  x.setHours(0,0,0,0);
  return x;
}

function endOfDay(d){
  const x=new Date(d);
  x.setHours(23,59,59,999);
  return x;
}

function firstDayNextMonth_(d){
  const x=new Date(d.getFullYear(), d.getMonth()+1, 1);
  x.setHours(0,0,0,0);
  return x;
}

function toISODate_(d){
  const y = d.getFullYear();
  const m = ('0'+(d.getMonth()+1)).slice(-2);
  const da= ('0'+d.getDate()).slice(-2);
  return `${y}-${m}-${da}`;
}

/** CPF sempre tratado como string */
function normalizeCPF(cpf){
  return String(cpf == null ? '' : cpf).replace(/\D/g,'');
}

function formatCPF(cpf){
  const num = normalizeCPF(cpf);
  return num.length===11
    ? num.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/,'$1.$2.$3-$4')
    : cpf;
}

function formatDateBR(d){
  if (!(d instanceof Date)) return '';
  return Utilities.formatDate(d, TZ, 'dd/MM/yyyy');
}

function parseISODateSafe(str){
  if (!str) return new Date();
  if (str instanceof Date) return str;
  if (typeof str === 'number') return new Date(str);

  let m = String(str).match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) return new Date(+m[3], +m[2]-1, +m[1]);

  m = String(str).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(+m[1], +m[2]-1, +m[3]);

  const d = new Date(str);
  return isNaN(d) ? new Date() : d;
}

/** ✔︎ Caixinhas de PARECER no PDF */
const BOX_CHECKED   = '☑';
const BOX_UNCHECKED = '☐';

function normalizeParecerOpcao_(v){
  const s = (v||'').toString()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .toLowerCase().trim();
  if (!s) return '';
  if (/(favoravel|aprovado|deferido|sim|favor)/.test(s)) return 'FAVORAVEL';
  if (/(desfavoravel|reprovado|indeferido|nao|não|contra)/.test(s)) return 'DESFAVORAVEL';
  return '';
}

/************************
 * DRIVE / PDFs
 ************************/
function safeFolderName_(s){
  const label = (s || 'Sem Setor').toString().trim();
  return label
    .replace(/[\\\/<>:"|?*\u0000-\u001F]+/g,'-')
    .replace(/\s+/g,' ')
    .trim();
}

function ensureSubfolder_(parent, name){
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function getPdfDestinationFolder_(setorLabel, dateObj){
  const tz   = TZ;
  const root = DriveApp.getFolderById(FOLDER_PDFS_ID);
  const setor= ensureSubfolder_(root, safeFolderName_(setorLabel));
  const year = ensureSubfolder_(setor, Utilities.formatDate(dateObj,tz,'yyyy'));
  const month= ensureSubfolder_(year,  Utilities.formatDate(dateObj,tz,'MM'));
  const day  = ensureSubfolder_(month, Utilities.formatDate(dateObj,tz,'dd'));
  return day;
}

/**
 * Pasta específica do caso:
 *  FOLDER_PDFS_ID / Setor / Ano / Mês / Dia / "[PROTO] - NOME"
 */
function getCaseFolder_(setorLabel, dateObj, protocolo, nomeBeneficiario){
  const baseDay = getPdfDestinationFolder_(setorLabel, dateObj);
  const proto = (protocolo || '').toString().trim();
  const nome  = safeFolderName_(nomeBeneficiario || 'Beneficiário');
  let folderName = proto ? `${proto} - ${nome}` : nome;
  if (folderName.length > 120) folderName = folderName.substring(0,120);
  return ensureSubfolder_(baseDay, folderName);
}

function setPublicSharing_(file){
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(e){
    try {
      file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(_){}
  }
  return file;
}

function buildDriveLinks_(id){
  return {
    preview:`https://drive.google.com/file/d/${id}/preview`,
    download:`https://drive.google.com/uc?export=download&id=${id}`
  };
}

/** ***********************
 *  ASSINATURA DIGITAL
 **************************/
function dataUrlToBlob_(dataUrl, defaultMime) {
  if (!dataUrl) return null;
  const str = String(dataUrl);
  let mime = defaultMime || MimeType.PNG;
  let b64  = str;

  const m = str.match(/^data:([\w\/\-\+\.]+);base64,(.+)$/);
  if (m) {
    mime = m[1];
    b64  = m[2];
  }

  const bytes = Utilities.base64Decode(b64);
  return Utilities.newBlob(bytes, mime, 'assinatura.png');
}

function insertSignatureImageAtPlaceholder_(body, placeholderKey, blob) {
  if (!body || !blob || !placeholderKey) return;

  const pattern = '\\{\\{?\\s*' + placeholderKey + '\\s*\\}?\\}';
  const result = body.findText(pattern);
  if (!result) return;

  const el   = result.getElement();
  const text = el.asText();
  const start = result.getStartOffset();
  const end   = result.getEndOffsetInclusive();

  text.deleteText(start, end);

  let parent = text.getParent();
  while (parent && parent.getType() !== DocumentApp.ElementType.PARAGRAPH && parent.getParent) {
    parent = parent.getParent();
  }
  if (!parent || parent.getType() !== DocumentApp.ElementType.PARAGRAPH) return;

  const para = parent.asParagraph();
  const img = para.insertInlineImage(para.getNumChildren(), blob);
  try { if (img.getWidth() > 140) img.setWidth(140); } catch (_){}
}

/************************
 * DESCOBERTA AUTOMÁTICA
 ************************/
function getSheet_(name){
  return SpreadsheetApp.openById(PLANILHA_ID).getSheetByName(name);
}

function autoDetectDataSheet_(){
  const ss = SpreadsheetApp.openById(PLANILHA_ID);

  const pref = ss.getSheetByName(RESPOSTAS_SHEET_NAME);
  if (pref && pref.getLastRow() >= 2) return pref;

  const sheets = ss.getSheets();
  let best = null, bestScore = -1;

  sheets.forEach(sh => {
    const lr = sh.getLastRow();
    const lc = sh.getLastColumn();
    if (lr < 2 || lc < 3) return;

    const headers = sh.getRange(1,1,1,lc).getValues()[0];

    const iCPF   = findColIdx_(headers,
      CPF_HEADER,'cpf','cpf do beneficiário','cpf do beneficiario',
      'documento (cpf)','cpf beneficiário'
    );
    const iNome  = findColIdx_(headers,
      'nome','nome do beneficiário','nome do beneficiario',
      'beneficiário','beneficiario'
    );
    const iUnid  = findColIdx_(headers,
      'via de entrada','unidade','unidade / setor','setor',
      'cras','creas','cram'
    );
    const iBenef = findColIdx_(headers,
      'benefício','beneficio','demanda apresentada',
      'tipo de benefício','tipo de beneficio'
    );
    const iData  = findColIdx_(headers,
      'data da solicitação','data','timestamp',
      'solicitado em','data de cadastro'
    );

    const sampleRows = Math.min(400, lr-1);
    const rng = sh.getRange(2,1,sampleRows,lc).getValues();

    let cpfOk = 0;
    if (iCPF >= 0){
      for (let r=0;r<rng.length;r++){
        const n = String(rng[r][iCPF]||'').replace(/\D/g,'');
        if (n.length === 11) cpfOk++;
      }
    }

    let dataOk = 0;
    if (iData >= 0){
      for (let r=0;r<rng.length;r++){
        const v = rng[r][iData];
        if (v instanceof Date || typeof v === 'number'){ dataOk++; continue; }
        const s = String(v||'').trim();
        if (/^\d{2}\/\d{2}\/\d{4}$/.test(s) || /^\d{4}-\d{2}-\d{2}$/.test(s)) dataOk++;
      }
    }

    let unSet = new Set();
    if (iUnid >= 0){
      for (let r=0;r<rng.length;r++){
        const s = String(rng[r][iUnid]||'').trim();
        if (s) unSet.add(s);
      }
    }

    let headerScore = 0;
    [iCPF,iNome,iUnid,iBenef,iData].forEach(i => { if (i>=0) headerScore+=15; });

    const bulk  = Math.min(lr-1, 5000);
    const score = headerScore + cpfOk*2 + dataOk*1.5 + Math.min(unSet.size,50) + bulk/5;

    if (score > bestScore){ best = sh; bestScore = score; }
  });

  if (!best) throw new Error('Nenhuma aba de dados compatível encontrada.');
  return best;
}

function findColIdx_(headers, ...cands){
  const H = (headers||[]).map(h=>String(h||'').trim().toLowerCase());
  const flat = cands.flat();

  for (const cand of flat){
    const i = H.indexOf(String(cand||'').trim().toLowerCase());
    if (i >= 0) return i;
  }

  const patterns = flat.map(c =>
    new RegExp(String(c).replace(/[.*+?^${}()|[\]\\]/g,'\\$&'), 'i')
  );
  for (let i=0;i<(headers||[]).length;i++){
    const h = String(headers[i]||'');
    if (patterns.some(rx => rx.test(h))) return i;
  }
  return -1;
}

function ensureColumn_(sheet, header){
  const range   = sheet.getRange(1,1,1,Math.max(1,sheet.getLastColumn()));
  const headers = range.getValues()[0] || [];
  const idx     = findColIdx_(headers, header);
  if (idx === -1){
    sheet.insertColumnAfter(headers.length || 1);
    sheet.getRange(1, headers.length+1).setValue(header);
    return headers.length;
  }
  return idx;
}

function guessCpfColumn_(sh, startRow, lastRow, lastCol){
  const rows = Math.min(300, lastRow - startRow + 1);
  if (rows <= 0) return -1;

  let bestIdx = -1, bestScore = 0;
  for (let c=1;c<=lastCol;c++){
    const vals = sh.getRange(startRow, c, rows, 1).getValues();
    let score=0;
    for (let i=0;i<vals.length;i++){
      const n = String(vals[i][0]||'').replace(/\D/g,'');
      if (n.length === 11) score++;
    }
    if (score > bestScore){ bestScore = score; bestIdx = c-1; }
  }
  return bestScore >= 3 ? bestIdx : -1;
}

function guessDateColumn_(sh, startRow, lastRow, lastCol){
  const rows = Math.min(300, lastRow - startRow + 1);
  if (rows <= 0) return -1;

  let bestIdx = -1, bestScore = 0;
  for (let c=1;c<=lastCol;c++){
    const vals = sh.getRange(startRow, c, rows, 1).getValues();
    let score=0;
    for (let i=0;i<vals.length;i++){
      const v = vals[i][0];
      if (v instanceof Date || typeof v === 'number'){ score++; continue; }
      const s = String(v||'').trim();
      if (/^\d{2}\/\d{2}\/\d{4}$/.test(s) || /^\d{4}-\d{2}-\d{2}$/.test(s)) score++;
    }
    if (score > bestScore){ bestScore = score; bestIdx = c-1; }
  }
  return bestScore >= 3 ? bestIdx : -1;
}

function ensureDocumentosColumn_(sheet) {
  return ensureColumn_(sheet, 'Documentos');
}

/** ✅ (AJUSTADO) Recarrega cabeçalhos após inserir colunas */
function getRespostasIndexMap_(){
  const sh = autoDetectDataSheet_();

  let lastCol = sh.getLastColumn();
  let headers = sh.getRange(1,1,1,lastCol).getValues()[0] || [];

  const idx = {
    unidade:   findColIdx_(headers,
      'via de entrada','unidade','unidade / setor','setor','cras','creas','cram',
      'unidade/setor','setor/unidade','unidade/setor (cras/creas)'
    ),
    beneficio: findColIdx_(headers,
      'benefício','beneficio','benefício social',
      'demanda apresentada','tipo de benefício','tipo de beneficio','beneficio social'
    ),
    data:      findColIdx_(headers,
      'data da solicitação','data','timestamp',
      'data de cadastro','solicitado em','solicitação'
    ),
    nome:      findColIdx_(headers,
      'nome do beneficiário','beneficiário','beneficiario',
      'nome','nome do beneficiario','nome beneficiário'
    ),
    cpf:       findColIdx_(headers,
      CPF_HEADER,'cpf','cpf do beneficiário','cpf do beneficiario',
      'documento (cpf)','cpf beneficiário'
    ),
    status:    findColIdx_(headers,'status','situação','situacao'),
    pdf:       findColIdx_(headers,'pdf','link do pdf','link','arquivo','drive','url pdf','arquivo pdf'),
    docs:      findColIdx_(headers,'documentos','docs','anexos','arquivos anexos','documentos/anexos'),
    entregueFlag: -1
  };

  if (idx.cpf  < 0) idx.cpf  = guessCpfColumn_(sh, 2, Math.max(2, sh.getLastRow()), lastCol);
  if (idx.data < 0) idx.data = guessDateColumn_(sh, 2, Math.max(2, sh.getLastRow()), lastCol);

  let changed = false;
  if (idx.status < 0){ idx.status = ensureColumn_(sh, 'Status'); changed = true; }
  idx.entregueFlag = ensureColumn_(sh, ENTREGUE_HEADER); if (idx.entregueFlag >= headers.length) changed = true;
  if (idx.docs < 0){ idx.docs = ensureDocumentosColumn_(sh); changed = true; }

  if (changed){
    lastCol = sh.getLastColumn();
    headers = sh.getRange(1,1,1,lastCol).getValues()[0] || [];
  }

  if (idx.unidade  < 0) idx.unidade   = 1;
  if (idx.beneficio< 0) idx.beneficio = 2;
  if (idx.data     < 0) idx.data      = 3;
  if (idx.nome     < 0) idx.nome      = 4;
  if (idx.cpf      < 0) idx.cpf       = 12;

  return { sh, headers, idx };
}

/************************
 * AUX — ENTREGUE robusto
 ************************/
function isRowEntregue_(row, idx){
  const flag = String(row[idx.entregueFlag]||'').trim().toUpperCase();
  const st   = idx.status>=0 ? String(row[idx.status]||'').trim().toUpperCase() : '';
  return (
    flag === 'ENTREGUE' || flag === 'SIM' || flag === 'TRUE' || flag === 'VERDADEIRO' ||
    st === 'ENTREGUE' || st.includes('ENTREG')
  );
}

/************************
 * CPF: existência
 ************************/
function checkCpfExistence(cpf) {
  cpf = normalizeCPF(cpf);
  if (!cpf) return { exists:false };

  const { sh, idx } = getRespostasIndexMap_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { exists:false };

  const values = sh.getRange(2, idx.cpf+1, lastRow-1, 1).getValues();
  const exists = values.some(r => normalizeCPF(r[0]) === cpf);
  return { exists };
}

/************************
 * BENEFÍCIOS ILIMITADOS
 ************************/
function normalizeBenefit_(s){
  return (s || '').toString()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/[()\-]/g,' ')
    .replace(/[^a-zA-Z0-9 ]/g,'')
    .trim()
    .toLowerCase()
    .replace(/\s+/g,'');
}

const BENEFICIOS_ILIMITADOS = new Set([
  normalizeBenefit_('AUXÍLIO POR MORTE'),
  normalizeBenefit_('MAIS ACONCHEGO (NATALIDADE)'),
  normalizeBenefit_('PASSAGEM INTERESTADUAL (VIAGEM)'),
  normalizeBenefit_('Outro')
]);

/************************
 * 🔹 ANEXOS – helpers
 ************************/
function salvarAnexosNoDrive_(anexos, pastaDestino, protocolo) {
  if (!anexos || !anexos.length || !pastaDestino) return '';

  const linhas = [];
  anexos.forEach(function (item, idx) {
    try {
      if (!item) return;

      const nomeOriginal = (item.nome || item.filename || ('arquivo_' + (idx+1))).toString();
      const rotulo       = (item.nomeDocumento || '').toString().trim(); // ex.: RG, CPF, Comprovante
      const mime         = (item.mimeType || item.mimetype || MimeType.PDF);

      let base64 = item.base64 || item.conteudoBase64 || item.data || item.dataUrl || '';
      if (!base64) return;

      const m = String(base64).match(/^data:([\w\/\-\+\.]+);base64,(.+)$/);
      let mimeUse = mime;
      if (m) {
        mimeUse = m[1] || mimeUse;
        base64  = m[2] || '';
      }

      const bytes = Utilities.base64Decode(base64);
      const blob  = Utilities.newBlob(bytes, mimeUse, nomeOriginal);

      const nomeFinal = (protocolo ? protocolo + ' - ' : '') + (rotulo ? (rotulo + ' - ') : '') + nomeOriginal;

      const file = withRetry(function(){
        return pastaDestino.createFile(blob).setName(nomeFinal);
      }, 3, 300);

      setPublicSharing_(file);
      const links = buildDriveLinks_(file.getId());

      const etiqueta = rotulo || nomeOriginal;
      linhas.push(etiqueta + ': ' + links.preview);

    } catch (e) {
      Logger.log('[ANEXO] Erro ao salvar anexo: ' + e.message);
    }
  });

  return linhas.join('\n');
}

/************************
 * SALVAR + PDF + E-MAIL
 ************************/
function salvarFormulario(dados) {
  dados = dados || {};

  // Datas & formatos
  const dataSolicitacao   = parseISODateSafe(dados.data_solicitacao);
  const dataSolicitacaoBR = formatDateBR(dataSolicitacao);
  dados.data_solicitacao  = dataSolicitacaoBR;

  if (dados.nasc) {
    const nascDate = parseISODateSafe(dados.nasc);
    dados.nasc = formatDateBR(nascDate);
  }
  dados.cpf = formatCPF(dados.cpf);

  // Se não vier "cadastrador" do front, tenta o usuário logado
  if (!dados.cadastrador) {
    try {
      const up = PropertiesService.getUserProperties().getProperties();
      dados.cadastrador = up.semfas_usuario || Session.getActiveUser().getEmail() || '';
    } catch (_){
      dados.cadastrador = '';
    }
  }

  // 🔹 ASSINATURA DIGITAL (base64 vindo do formulário)
  const assinaturaB64 =
    dados.assinatura_base64 ||
    dados.assinaturaDigitalBase64 ||
    dados.assinatura_digital_base64 ||
    dados.assinaturaDigital ||
    dados.assinatura_img ||
    dados.assinaturaImagem ||
    dados.assinatura_digital ||
    '';

  let assinaturaBlob = null;
  if (assinaturaB64) {
    try {
      assinaturaBlob = dataUrlToBlob_(assinaturaB64, MimeType.PNG);
    } catch (e) {
      Logger.log('[ASSINATURA] Erro ao converter base64: ' + e.message);
    }
  }

  // 🔹 anexos (array vindo do front; se não vier, fica vazio)
  const anexos = Array.isArray(dados.anexos) ? dados.anexos : [];

  const shPref = getSheet_(RESPOSTAS_SHEET_NAME);
  const sh = (shPref && shPref.getLastRow() >= 1) ? shPref : autoDetectDataSheet_();

  let newRow, protocoloGerado = '';
  const lock1 = LockService.getScriptLock(); lock1.waitLock(10000);

  try {
    // ✅ (AJUSTADO) duplicidade usando índices detectados (não fixos)
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    const headers = sh.getRange(1,1,1,lastCol).getValues()[0] || [];

    let iCpf  = findColIdx_(headers, CPF_HEADER,'cpf','cpf do beneficiário','cpf do beneficiario','cpf beneficiário');
    let iBen  = findColIdx_(headers, 'demanda apresentada','benefício','beneficio','tipo de benefício','tipo de beneficio','benefício social','beneficio social');
    let iData = findColIdx_(headers, 'data da solicitação','data','timestamp','solicitado em','data de cadastro');

    if (iCpf  < 0) iCpf  = guessCpfColumn_(sh, 2, Math.max(2,lastRow), lastCol);
    if (iData < 0) iData = guessDateColumn_(sh, 2, Math.max(2,lastRow), lastCol);

    const registros = (lastRow >= 2)
      ? sh.getRange(2,1,lastRow-1,lastCol).getValues()
      : [];

    const cpfNovo       = normalizeCPF(dados.cpf);
    const beneficioNovo = (dados.demanda || '').toString().trim();
    const benKeyNovo    = normalizeBenefit_(beneficioNovo);
    const mesNovo       = dataSolicitacao.getMonth();
    const anoNovo       = dataSolicitacao.getFullYear();
    const isIlimitado   = BENEFICIOS_ILIMITADOS.has(benKeyNovo);

    let lastMatchDate = null;
    for (let i=0; i<registros.length; i++) {
      const row = registros[i];

      const cpfExistente = (iCpf>=0 && iCpf<row.length) ? normalizeCPF(row[iCpf]) : '';
      const beneficioExistente = (iBen>=0 && iBen<row.length) ? String(row[iBen]||'').trim() : '';
      const benKeyExistente = normalizeBenefit_(beneficioExistente);

      const dataExistente = (iData>=0 && iData<row.length) ? parseAnyDate(row[iData]) : null;

      if (cpfExistente === cpfNovo &&
          benKeyExistente === benKeyNovo &&
          dataExistente &&
          dataExistente.getMonth() === mesNovo &&
          dataExistente.getFullYear() === anoNovo) {
        if (!lastMatchDate || dataExistente > lastMatchDate) lastMatchDate = dataExistente;
      }
    }

    if (!isIlimitado && lastMatchDate) {
      const liberadoEm = firstDayNextMonth_(lastMatchDate);
      return {
        sucesso:false,
        mensagem:'Já existe um cadastro para este CPF com este benefício neste mês.',
        ultima_data:  formatDateBR(lastMatchDate),
        proxima_data: formatDateBR(liberadoEm)
      };
    }

    // Escreve a linha (estrutura padrão; se a planilha tiver menos colunas, ela expande)
    const linha = [
      new Date(), dados.via_entrada, dados.demanda, dataSolicitacaoBR,
      dados.nomeb, dados.nasc, dados.endereco, dados.bairro, dados.referencia,
      dados.telefone, dados.rg, dados.ssp, dados.cpf, (dados.nis || ''),
      dados.cras_ref, dados.tem_beneficio, dados.qual_beneficio, dados.membros,
      dados.renda, dados.nomes, dados.telefones, dados.enderecos, dados.rgs,
      dados.cpfs, dados.vinculo, dados.situacao,
      'PENDENTE', '', dados.cadastrador
    ];

    withRetry(()=>sh.appendRow(linha), 5, 250);
    newRow = sh.getLastRow();

    // Protocolo
    protocoloGerado = ensureProtocoloForRow_(sh, newRow);

    // PARECER TÉCNICO -> grava em colunas dedicadas
    const iPar   = ensureColumn_(sh, 'PARECER TÉCNICO');
    const iNomT  = ensureColumn_(sh, 'NOME TÉCNICO');
    const iMun   = ensureColumn_(sh, 'MUNICÍPIO PARECER');
    const iDataP = ensureColumn_(sh, 'DATA PARECER');
    const iAssB  = ensureColumn_(sh, 'ASSINATURA BENEFICIÁRIO');
    const iAssT  = ensureColumn_(sh, 'ASSINATURA/CARIMBO');

    const dataParecerBR = dados.data_parecer
      ? formatDateBR(parseISODateSafe(dados.data_parecer))
      : '';

    const assinaturaSheetValue = assinaturaB64 || dados.assinatura_carimbo || '';

    sh.getRange(newRow, iPar+1 ).setValue(dados.parecer_tecnico || dados.parecer || '');
    sh.getRange(newRow, iNomT+1).setValue(dados.nome_tecnico || '');
    sh.getRange(newRow, iMun+1 ).setValue(dados.municipio_parecer || 'Nossa Senhora do Socorro');
    sh.getRange(newRow, iDataP+1).setValue(dataParecerBR);
    sh.getRange(newRow, iAssB+1).setValue(dados.assinatura_beneficiario || '');
    sh.getRange(newRow, iAssT+1).setValue(assinaturaSheetValue);

  } finally {
    lock1.releaseLock();
  }

  // Pasta específica do caso
  const setorLabel = dados.via_entrada || 'Sem Setor';
  const dataRef    = parseISODateSafe(dados.data_solicitacao || new Date());
  const pastaCaso  = getCaseFolder_(setorLabel, dataRef, protocoloGerado, dados.nomeb);

  // Gera PDF
  const pdf = gerarFichaPDF(Object.assign({}, dados, {
    protocolo: protocoloGerado,
    __assinaturaBlob__: assinaturaBlob
  }), pastaCaso);

  // Atualiza célula do PDF + Documentos
  const lock2 = LockService.getScriptLock(); lock2.waitLock(10000);
  try {
    const iPdf  = ensureColumn_(sh, 'PDF');
    const iDocs = ensureDocumentosColumn_(sh);

    withRetry(()=>sh.getRange(newRow, iPdf+1).setValue(pdf.downloadUrl), 5, 250);

    // 🔹 ANEXOS → mesma pasta do caso
    if (anexos && anexos.length) {
      const textoDocs = salvarAnexosNoDrive_(anexos, pastaCaso, protocoloGerado);
      if (textoDocs) {
        const anterior = sh.getRange(newRow, iDocs+1).getValue() || '';
        const novo     = anterior ? (anterior + '\n' + textoDocs) : textoDocs;
        withRetry(() => sh.getRange(newRow, iDocs+1).setValue(novo), 5, 250);
      }
    }
  } finally {
    lock2.releaseLock();
  }

  // E-mail ao setor (se configurado)
  enviarEmailSetor(dados.via_entrada, pdf, dados.nomeb);

  return {
    sucesso:true,
    mensagem:'Cadastro realizado, PDF e anexos salvos em pasta organizada, e e-mail enviado ao setor!',
    link: pdf.downloadUrl,
    protocolo: protocoloGerado,
    concluirTexto:'Concluído'
  };
}

/************************
 * GERAÇÃO DO PDF (template + assinatura)
 ************************/
function gerarFichaPDF(dados, destFolderOpt) {
  const dSolic = coerceSheetDate(dados.data_solicitacao || dados.data || dados.dataSolicitacao);
  const dNasc  = coerceSheetDate(dados.nasc || dados.data_nascimento || dados.dataNascimento);

  const assinaturaBlob = dados.__assinaturaBlob__ || null;
  const parecerOp = normalizeParecerOpcao_(dados.parecer_opcao || dados.parecer_tecnico || dados.parecer);

  const payload = Object.assign({}, dados, {
    data_solicitacao: dSolic ? formatDateBR(dSolic) : (dados.data_solicitacao || ''),
    nasc:             dNasc  ? formatDateBR(dNasc)  : (dados.nasc || ''),
    cpf:              formatCPF(dados.cpf || ''),
    cadastrador:      dados.cadastrador || '',

    parecer:                 dados.parecer_tecnico || dados.parecer || '',
    nome_tecnico:            dados.nome_tecnico || '',
    municipio_parecer:       dados.municipio_parecer || 'Nossa Senhora do Socorro',
    data_parecer:            dados.data_parecer ? formatDateBR(parseISODateSafe(dados.data_parecer)) : '',
    assinatura_beneficiario: dados.assinatura_beneficiario || '',
    assinatura_carimbo:      dados.assinatura_carimbo || '',
    protocolo:               dados.protocolo || '',

    box_favoravel:    (parecerOp === 'FAVORAVEL')    ? BOX_CHECKED   : BOX_UNCHECKED,
    box_desfavoravel: (parecerOp === 'DESFAVORAVEL') ? BOX_CHECKED   : BOX_UNCHECKED,
    box_indef:        (!parecerOp)                   ? BOX_CHECKED   : BOX_UNCHECKED
  });

  delete payload.__assinaturaBlob__;

  const modelo = DriveApp.getFileById(ID_TEMPLATE);
  const copia  = withRetry(
    () => modelo.makeCopy('Ficha - ' + (payload.nomeb || 'Beneficiário')),
    5, 250
  );
  const doc   = DocumentApp.openById(copia.getId());
  const corpo = doc.getBody();

  try {
    corpo
      .setMarginTop(36)
      .setMarginBottom(36)
      .setMarginLeft(36)
      .setMarginRight(36);
  } catch (_){}

  function replacePlaceholder_(key, value) {
    const val = (value == null ? '' : String(value)).replace(/\r\n/g,'\n');
    const pattern = '\\{\\{\\s*' + key + '\\s*\\}\\}|\\{\\s*' + key + '\\s*\\}';
    try { corpo.replaceText(pattern, val); }
    catch (e) { Logger.log('[PDF] replaceText ' + key + ': ' + e.message); }
  }

  Object.keys(payload).forEach(k => {
    const isAssinaturaKey =
      k === 'assinatura_carimbo' ||
      k === 'assinatura_img' ||
      k === 'assinatura_digital' ||
      k === 'assinaturaImagem';

    if (assinaturaBlob && isAssinaturaKey) return;
    replacePlaceholder_(k, payload[k]);
  });

  if (assinaturaBlob) {
    try { insertSignatureImageAtPlaceholder_(corpo, 'assinatura_carimbo', assinaturaBlob); }
    catch (e) { Logger.log('[ASSINATURA] Erro ao inserir imagem: ' + e.message); }
  }

  try { corpo.replaceText('\\{\\s*([^{}]+)\\s*\\}', '$1'); } catch (e) { Logger.log('[PDF] limpar chaves simples: ' + e.message); }
  try { corpo.replaceText('\\{\\{\\s*[^}]+\\s*\\}\\}', ''); } catch (e) { Logger.log('[PDF] limpar placeholders {{}}: ' + e.message); }

  doc.saveAndClose();
  Utilities.sleep(300);

  const nomeArquivo = `Ficha - ${(payload.nomeb || 'Beneficiário')} - ${(payload.cpf || '').toString()}.pdf`;

  const pdfBlob = withRetry(
    () => DriveApp.getFileById(copia.getId()).getAs(MimeType.PDF).copyBlob().setName(nomeArquivo),
    5, 250
  );

  const setor   = safeFolderName_(payload.via_entrada || 'Sem Setor');
  const dataRef = dSolic || new Date();

  const pastaDestino = (destFolderOpt && destFolderOpt.createFile)
    ? destFolderOpt
    : getPdfDestinationFolder_(setor, dataRef);

  const pdfFile = withRetry(() => pastaDestino.createFile(pdfBlob), 5, 250);
  setPublicSharing_(pdfFile);

  const fileId = pdfFile.getId();
  const links  = buildDriveLinks_(fileId);

  try { DriveApp.getFileById(copia.getId()).setTrashed(true); } catch (_){}

  return {
    blob: pdfBlob,
    file: pdfFile,
    fileId,
    previewUrl:  links.preview,
    downloadUrl: links.download
  };
}

/************************
 * GERAR PDF A PARTIR DE UMA LINHA EXISTENTE
 ************************/
function gerarPdfDeRegistro(linha) {
  const shPref = getSheet_(RESPOSTAS_SHEET_NAME);
  const sh = (shPref && shPref.getLastRow() >= 1) ? shPref : autoDetectDataSheet_();

  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0] || [];
  const dados   = sh.getRange(linha,1,1,sh.getLastColumn()).getValues()[0];

  const dSolic = coerceSheetDate(dados[3]);
  const dNasc  = coerceSheetDate(dados[5]);

  const iProt = findColIdx_(headers,
    PROTOCOLO_HDR,'protocolo','nº protocolo','n° protocolo',
    'numero do protocolo','id','id protocolo'
  );
  const protocoloLido = iProt >= 0 ? (dados[iProt] || '') : '';

  const iPar   = findColIdx_(headers, 'PARECER TÉCNICO');
  const iNomT  = findColIdx_(headers, 'NOME TÉCNICO');
  const iMun   = findColIdx_(headers, 'MUNICÍPIO PARECER');
  const iData  = findColIdx_(headers, 'DATA PARECER');
  const iAssB  = findColIdx_(headers, 'ASSINATURA BENEFICIÁRIO','ASSINATURA BENEFICIARIO');
  const iAssT  = findColIdx_(headers,
    'ASSINATURA/CARIMBO',
    'ASSINATURA TÉCNICO','ASSINATURA TECNICO',
    'ASSINATURA TEC','ASSINATURA TEC.','ASSINATURA'
  );

  let assinaturaBlob = null;
  if (iAssT >= 0) {
    const sigCell = dados[iAssT];
    if (sigCell) {
      const s = String(sigCell).trim();
      try {
        if (/^data:image\//i.test(s) || (/^[A-Za-z0-9+/=]+$/.test(s) && s.length > 100)) {
          assinaturaBlob = dataUrlToBlob_(s, MimeType.PNG);
        }
      } catch (e) {
        Logger.log('[ASSINATURA] erro ao recriar blob da planilha: ' + e.message);
      }
    }
  }

  const registro = {
    via_entrada: dados[1],
    demanda: dados[2],
    data_solicitacao: dSolic ? formatDateBR(dSolic) : '',
    nomeb: dados[4],
    nasc: dNasc ? formatDateBR(dNasc) : '',
    endereco: dados[6],
    bairro: dados[7],
    referencia: dados[8],
    telefone: dados[9],
    rg: dados[10],
    ssp: dados[11],
    cpf: formatCPF(dados[12] || ''),
    nis: dados[13],
    cras_ref: dados[14],
    tem_beneficio: dados[15],
    qual_beneficio: dados[16],
    membros: dados[17],
    renda: dados[18],
    nomes: dados[19],
    telefones: dados[20],
    enderecos: dados[21],
    rgs: dados[22],
    cpfs: dados[23],
    vinculo: dados[24],
    situacao: dados[25],
    cadastrador: dados[29] || '',
    protocolo: protocoloLido || '',
    parecer: iPar>=0 ? (dados[iPar]||'') : '',
    nome_tecnico: iNomT>=0 ? (dados[iNomT]||'') : '',
    municipio_parecer: iMun>=0 ? (dados[iMun]||'') : '',
    data_parecer: iData>=0 ? (dados[iData]||'') : '',
    assinatura_beneficiario: iAssB>=0 ? (dados[iAssB]||'') : '',
    assinatura_carimbo:      iAssT>=0 ? (dados[iAssT]||'') : ''
  };

  const setorLabel = registro.via_entrada || 'Sem Setor';
  const dataRefReg = dSolic || new Date();
  const pastaCaso  = getCaseFolder_(setorLabel, dataRefReg, protocoloLido, registro.nomeb);

  const pdf = gerarFichaPDF(Object.assign({}, registro, {
    __assinaturaBlob__: assinaturaBlob
  }), pastaCaso);

  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const iPdf = ensureColumn_(sh, 'PDF');
    sh.getRange(linha, iPdf+1).setValue(pdf.downloadUrl);
  } finally {
    lock.releaseLock();
  }
  return { link: pdf.downloadUrl };
}

/************************
 * CENTRAL (consulta)
 ************************/
function buscarBeneficios(filtroCPF = "", dataInicio = "", dataFim = "", beneficio = "", unidade = "") {
  const { sh, headers, idx } = getRespostasIndexMap_();
  const dados = sh.getDataRange().getValues();

  const listaUnidades   = new Set();
  const listaBeneficios = new Set();
  for (let i=1; i<dados.length; i++){
    const r = dados[i];
    if (r[idx.unidade])   listaUnidades.add(String(r[idx.unidade]).trim());
    if (r[idx.beneficio]) listaBeneficios.add(String(r[idx.beneficio]).trim());
  }
  unidade   = resolveFiltroOuVazio_(Array.from(listaUnidades),   unidade);
  beneficio = resolveFiltroOuVazio_(Array.from(listaBeneficios), beneficio);

  const registros = [];
  const cpfFiltroNum = normalizeCPF(filtroCPF);
  const dIni = parseAnyDate(dataInicio);
  const dFim = parseAnyDate(dataFim);
  const benFiltro = (beneficio || '').toString().trim();
  const uniFiltro = (unidade   || '').toString().trim();

  const iProt = findColIdx_(headers, PROTOCOLO_HDR,'protocolo','nº protocolo','n° protocolo','numero do protocolo','id','id protocolo');

  for (let i=1; i<dados.length; i++) {
    const row = dados[i];

    const cpfStr   = (row[idx.cpf] || '').toString().trim();
    const cpfNum   = normalizeCPF(cpfStr);
    const status   = isRowEntregue_(row, idx) ? 'ENTREGUE' : (String(row[idx.status] || 'PENDENTE').toUpperCase());
    const data     = coerceSheetDate(row[idx.data]);
    const demanda  = (row[idx.beneficio] || '').toString().trim();
    const unidadeV = (row[idx.unidade] || '').toString().trim();
    const nome     = (row[idx.nome] || '').toString().trim();
    const protocolo= iProt>=0 ? String(row[iProt]||'').trim() : '';

    let okData = true;
    if (dIni && data && data < startOfDay(dIni)) okData = false;
    if (dFim && data && data > endOfDay(dFim)) okData = false;

    if (benFiltro && !matchesFilter_(demanda, benFiltro)) continue;
    if (uniFiltro && !matchesFilter_(unidadeV, uniFiltro)) continue;

    if ((!cpfFiltroNum || cpfNum === cpfFiltroNum) && okData) {
      registros.push({
        linha: i+1,
        data: data ? formatDateBR(data) : '',
        unidade: unidadeV,
        demanda,
        nome,
        cpf: formatCPF(cpfStr),
        protocolo,
        status,
        linkPdf: (idx.pdf>=0 ? (row[idx.pdf] || '') : ''),
        documentos: (idx.docs>=0 ? (row[idx.docs] || '') : '')
      });
    }
  }
  return { registros };
}

function matchesFilter_(value, filterTxt){
  const f = normText_(filterTxt);
  if (!f) return true;
  return normText_(value).includes(f);
}

function resolveFiltroOuVazio_(lista, valor){
  const v = normText_(valor||'');
  if (!v) return '';
  for (const item of lista){
    const n = normText_(item);
    if (n.includes(v) || v.includes(n)) return valor;
  }
  return '';
}

/************************
 * BAIXA — Opções / Lista / Busca / Entrega
 ************************/
function ensureBaixasSheet_(){
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  let sh = ss.getSheetByName(BAIXAS_SHEET_NAME);
  if (!sh){
    sh = ss.insertSheet(BAIXAS_SHEET_NAME);
    sh.appendRow([
      'Carimbo','Protocolo','CPF','Nome','Benefício','Status Antes','Status Depois',
      'Entregue em','Unidade','Entregue por','Observação','RowRef'
    ]);
  }
  return sh;
}

function sortList_(arr, order){
  const o = String(order||'data_desc').toLowerCase();
  const coll = (a,b)=>String(a||'').localeCompare(String(b||''),'pt-BR',{sensitivity:'base'});
  const getT = v => { const d = coerceSheetDate(v); return d ? d.getTime() : 0; };
  const key = (r, k)=>{
    switch(k){
      case 'data':   return getT(r.solicitadoEm || r.data || r.dataBR || r.dataISO);
      case 'nome':   return normText_(r.nome);
      case 'benef':  return normText_(r.beneficio);
      case 'uni':    return normText_(r.unidade);
      case 'status': return normText_(r.status);
      case 'prot':   return normText_(r.protocolo);
      case 'cpf':    return (String(r.cpf||'').replace(/\D/g,'')) || '';
      default:       return '';
    }
  };
  if (o === 'data_asc')  return arr.sort((a,b)=> key(a,'data') - key(b,'data'));
  if (o === 'data_desc') return arr.sort((a,b)=> key(b,'data') - key(a,'data'));
  if (o === 'nome_asc')  return arr.sort((a,b)=> coll(key(a,'nome'),  key(b,'nome')));
  if (o === 'nome_desc') return arr.sort((a,b)=> coll(key(b,'nome'),  key(a,'nome')));
  if (o === 'benef_asc') return arr.sort((a,b)=> coll(key(a,'benef'), key(b,'benef')));
  if (o === 'benef_desc')return arr.sort((a,b)=> coll(key(b,'benef'), key(a,'benef')));
  if (o === 'uni_asc')   return arr.sort((a,b)=> coll(key(a,'uni'),   key(b,'uni')));
  if (o === 'uni_desc')  return arr.sort((a,b)=> coll(key(b,'uni'),   key(a,'uni')));
  if (o === 'prot_asc')  return arr.sort((a,b)=> coll(key(a,'prot'),  key(b,'prot')));
  if (o === 'prot_desc') return arr.sort((a,b)=> coll(key(b,'prot'),  key(a,'prot')));
  if (o === 'cpf_asc')   return arr.sort((a,b)=> coll(key(a,'cpf'),   key(b,'cpf')));
  if (o === 'cpf_desc')  return arr.sort((a,b)=> coll(key(b,'cpf'),   key(a,'cpf')));
  return arr;
}

function listarUnidadesDashboard(){
  const { sh, idx } = getRespostasIndexMap_();
  const dados = sh.getDataRange().getValues().slice(1);
  const set = new Set();
  dados.forEach(l => {
    const u = (l[idx.unidade] || '').toString().trim();
    if (u) set.add(u);
  });
  return Array.from(set).sort();
}

function listarBeneficiosDashboard(){
  const { sh, idx } = getRespostasIndexMap_();
  const dados = sh.getDataRange().getValues().slice(1);
  const set = new Set();
  dados.forEach(l => {
    const b = (l[idx.beneficio] || '').toString().trim();
    if (b) set.add(b);
  });
  return Array.from(set).sort((a,b)=>a.localeCompare(b,'pt-BR',{sensitivity:'base'}));
}

function getOpcoesFiltros(){
  return {
    unidades: listarUnidadesDashboard(),
    beneficios: listarBeneficiosDashboard()
  };
}

function baixa_listarOpcoes(){
  return {
    unidades: listarUnidadesDashboard(),
    beneficios: listarBeneficiosDashboard()
  };
}

function baixa_list(statusFilter='todos', unidadeFiltro='', beneficioFiltro='', order='data_desc', pdfOnly=false, hasProto=false, dataInicio='', dataFim=''){
  const { sh, headers, idx } = getRespostasIndexMap_();
  const values = sh.getRange(2,1, Math.max(0, sh.getLastRow()-1), sh.getLastColumn()).getValues();

  const setUnidades = new Set(), setBenef = new Set();
  for (let r=0;r<values.length;r++){
    const row = values[r];
    if (row[idx.unidade])   setUnidades.add(String(row[idx.unidade]).trim());
    if (row[idx.beneficio]) setBenef.add(String(row[idx.beneficio]).trim());
  }
  unidadeFiltro   = resolveFiltroOuVazio_(Array.from(setUnidades), unidadeFiltro);
  beneficioFiltro = resolveFiltroOuVazio_(Array.from(setBenef),    beneficioFiltro);

  const wantPend = String(statusFilter).toLowerCase() === 'pendentes';
  const wantEnt  = String(statusFilter).toLowerCase() === 'entregues';
  const iProt = findColIdx_(headers, PROTOCOLO_HDR,'protocolo','nº protocolo','n° protocolo','numero do protocolo','id','id protocolo');

  const dIni = parseAnyDate(dataInicio);
  const dFim = parseAnyDate(dataFim);

  const out = [];
  for (let r=0; r<values.length; r++){
    const row = values[r];

    const unidade = (row[idx.unidade] || '').toString().trim();
    const benef   = (row[idx.beneficio] || '').toString().trim();
    if (unidadeFiltro && !matchesFilter_(unidade, unidadeFiltro)) continue;
    if (beneficioFiltro && !matchesFilter_(benef, beneficioFiltro)) continue;

    const isEnt   = isRowEntregue_(row, idx);
    if (wantPend && isEnt) continue;
    if (wantEnt  && !isEnt) continue;

    const stCell  = idx.status>=0 ? String(row[idx.status]||'').toUpperCase() : '';
    const status  = isEnt ? 'ENTREGUE' : (stCell || 'PENDENTE');
    const dataObj = coerceSheetDate(row[idx.data]);

    if (dIni && dataObj && dataObj < startOfDay(dIni)) continue;
    if (dFim && dataObj && dataObj > endOfDay(dFim))   continue;

    const linkPdf = idx.pdf>=0 ? (row[idx.pdf] || '') : '';
    if (pdfOnly && !linkPdf) continue;

    const protStr = iProt>=0 ? String(row[iProt]||'') : '';
    if (hasProto && !protStr) continue;

    out.push({
      rowRef:      r+2,
      protocolo:   protStr || '',
      nome:        idx.nome>=0 ? String(row[idx.nome]||'').trim() : '',
      cpf:         idx.cpf>=0 ? formatCPF(String(row[idx.cpf]||'')) : '',
      beneficio:   benef,
      unidade,
      solicitadoEm: dataObj ? formatDateBR(dataObj) : '',
      status,
      pdf:         linkPdf,
      entregue:    isEnt,
      documentos:  idx.docs>=0 ? (row[idx.docs] || '') : ''
    });
  }

  sortList_(out, order);
  return { registros: out.slice(0,2000) };
}

function baixa_search(q, statusFilter='todos', unidadeFiltro='', beneficioFiltro='', order='data_desc', pdfOnly=false, hasProto=false, dataInicio='', dataFim=''){
  q = String(q || '').trim();
  if (!q) return { registros: [] };

  const { sh, headers, idx } = getRespostasIndexMap_();
  const values = sh.getRange(2,1, Math.max(0, sh.getLastRow()-1), sh.getLastColumn()).getValues();

  const setUnidades = new Set(), setBenef = new Set();
  for (let r=0;r<values.length;r++){
    const row = values[r];
    if (row[idx.unidade])   setUnidades.add(String(row[idx.unidade]).trim());
    if (row[idx.beneficio]) setBenef.add(String(row[idx.beneficio]).trim());
  }
  unidadeFiltro   = resolveFiltroOuVazio_(Array.from(setUnidades), unidadeFiltro);
  beneficioFiltro = resolveFiltroOuVazio_(Array.from(setBenef),    beneficioFiltro);

  const wantPend = String(statusFilter).toLowerCase() === 'pendentes';
  const wantEnt  = String(statusFilter).toLowerCase() === 'entregues';
  const iProt = findColIdx_(headers, PROTOCOLO_HDR,'protocolo','nº protocolo','n° protocolo','numero do protocolo','id','id protocolo');

  const qLower = q.toLowerCase();
  const qNum   = q.replace(/\D/g,'');
  const qNorm  = normText_(q);
  const dIni   = parseAnyDate(dataInicio);
  const dFim   = parseAnyDate(dataFim);

  const out = [];
  for (let r=0; r<values.length; r++){
    const row = values[r];

    const unidade = (row[idx.unidade] || '').toString().trim();
    const benef   = (row[idx.beneficio] || '').toString().trim();
    if (unidadeFiltro && !matchesFilter_(unidade, unidadeFiltro)) continue;
    if (beneficioFiltro && !matchesFilter_(benef, beneficioFiltro)) continue;

    const cpfStr   = idx.cpf>=0 ? String(row[idx.cpf]||'') : '';
    const cpfNum   = cpfStr.replace(/\D/g,'');
    const nomeStr  = idx.nome>=0 ? String(row[idx.nome]||'') : '';
    const nomeNorm = normText_(nomeStr);
    const protStr  = iProt>=0 ? String(row[iProt]||'') : '';

    const match =
      (qNum && (cpfNum.includes(qNum) || protStr.replace(/\D/g,'').includes(qNum))) ||
      (qNorm && nomeNorm.includes(qNorm)) ||
      (iProt>=0 && protStr.toLowerCase().includes(qLower));
    if (!match) continue;

    const isEnt   = isRowEntregue_(row, idx);
    if (wantPend && isEnt) continue;
    if (wantEnt  && !isEnt) continue;

    const stCell  = idx.status>=0 ? String(row[idx.status]||'').toUpperCase() : '';
    const status  = isEnt ? 'ENTREGUE' : (stCell || 'PENDENTE');
    const dataObj = coerceSheetDate(row[idx.data]);

    if (dIni && dataObj && dataObj < startOfDay(dIni)) continue;
    if (dFim && dataObj && dataObj > endOfDay(dFim))   continue;

    const linkPdf = idx.pdf>=0 ? (row[idx.pdf] || '') : '';
    if (pdfOnly && !linkPdf) continue;
    if (hasProto && !protStr) continue;

    out.push({
      rowRef:      r+2,
      protocolo:   protStr,
      nome:        nomeStr,
      cpf:         formatCPF(cpfStr),
      beneficio:   benef,
      unidade,
      solicitadoEm: dataObj ? formatDateBR(dataObj) : '',
      status,
      pdf:         linkPdf,
      entregue:    isEnt,
      documentos:  idx.docs>=0 ? (row[idx.docs] || '') : ''
    });
  }

  sortList_(out, order);
  return { registros: out };
}

function baixa_marcarEntrega(payload){
  const { rowRef, dataEntrega, setor, usuario, obs } = payload || {};
  if (!rowRef) throw new Error('RowRef não informado');

  const { sh, headers, idx } = getRespostasIndexMap_();

  const iEnt    = idx.entregueFlag;
  const iEntEm  = ensureColumn_(sh, ENTREGUE_EM_HDR);
  const iEntPor = ensureColumn_(sh, ENTREGUE_POR_HDR);
  const iEntUni = ensureColumn_(sh, ENTREGUE_UNID_HDR);
  const iEntObs = ensureColumn_(sh, ENTREGUE_OBS_HDR);

  let iStatus = idx.status;
  if (iStatus < 0) iStatus = ensureColumn_(sh, 'Status');

  const rowVals = sh.getRange(rowRef, 1, 1, sh.getLastColumn()).getValues()[0];
  const antes = isRowEntregue_(rowVals, idx)
    ? 'ENTREGUE'
    : String(rowVals[iStatus] || 'PENDENTE').toUpperCase();

  const agora = new Date();
  let dt = agora;

  if (dataEntrega) {
    const d = parseISODateSafe(dataEntrega);
    if (d instanceof Date && !isNaN(d.getTime())) {
      d.setHours(agora.getHours(), agora.getMinutes(), agora.getSeconds(), 0);
      dt = d;
    }
  }

  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    sh.getRange(rowRef, iEnt+1).setValue('ENTREGUE');
    sh.getRange(rowRef, iEntEm+1).setValue(dt);
    sh.getRange(rowRef, iEntPor+1).setValue(usuario || '');
    sh.getRange(rowRef, iEntUni+1).setValue(setor   || '');
    sh.getRange(rowRef, iEntObs+1).setValue(obs     || '');
    sh.getRange(rowRef, iStatus+1).setValue('ENTREGUE');
  } finally {
    lock.releaseLock();
  }

  const log     = ensureBaixasSheet_();
  const idxCPF  = idx.cpf;
  const idxNome = idx.nome;
  const idxBen  = idx.beneficio;
  const iProt   = findColIdx_(headers,
    PROTOCOLO_HDR,'protocolo','nº protocolo','n° protocolo',
    'numero do protocolo','id','id protocolo'
  );

  log.appendRow([
    new Date(),
    iProt   >=0 ? rowVals[iProt]   : '',
    idxCPF  >=0 ? rowVals[idxCPF]  : '',
    idxNome >=0 ? rowVals[idxNome] : '',
    idxBen  >=0 ? rowVals[idxBen]  : '',
    antes, 'ENTREGUE', dt, setor||'', usuario||'', obs||'', rowRef
  ]);

  return true;
}

/************************
 * DASHBOARD
 ************************/
function getDashboardCompleto(inicio, fim, unidade = "") {
  const { sh, idx } = getRespostasIndexMap_();
  const rows = sh.getDataRange().getValues();
  if (rows.length <= 1) return payloadDashboardVazio_();

  const dados = rows.slice(1);
  const dIni = parseAnyDate(inicio);
  const dFim = parseAnyDate(fim);

  const statusTotais = { SOLICITADO:0, APROVADO:0, PENDENTE:0, RECUSADO:0, ENTREGUE:0 };
  const tipos = {};
  const setores = {};
  const meses = {};
  const byMonthStatus = {};
  const dow = [0,0,0,0,0,0,0];
  const registros = [];

  dados.forEach(linha => {
    const data = coerceSheetDate(linha[idx.data]);
    if (!data) return;
    if (dIni && data < startOfDay(dIni)) return;
    if (dFim && data > endOfDay(dFim)) return;

    const setor  = (linha[idx.unidade] || 'Indefinido').toString().trim();
    if (unidade && !matchesFilter_(setor, unidade)) return;

    const isEnt  = isRowEntregue_(linha, idx);
    const status = isEnt ? 'ENTREGUE' : (String(linha[idx.status] || 'PENDENTE').toUpperCase());
    const tipo   = (linha[idx.beneficio] || 'Indefinido').toString().trim();
    const cpfStr = (linha[idx.cpf] || '').toString().trim();
    const nome   = (linha[idx.nome] || '').toString().trim();

    statusTotais.SOLICITADO++;
    if (statusTotais.hasOwnProperty(status)) statusTotais[status]++;

    tipos[tipo] = (tipos[tipo] || 0) + 1;

    if (!setores[setor]) setores[setor] = { SOLICITADO:0, APROVADO:0, PENDENTE:0, RECUSADO:0, ENTREGUE:0 };
    setores[setor].SOLICITADO++;
    if (setores[setor].hasOwnProperty(status)) setores[setor][status]++;

    const chave = (('0'+(data.getMonth()+1)).slice(-2)) + '/' + data.getFullYear();
    meses[chave] = (meses[chave] || 0) + 1;

    if (!byMonthStatus[chave]) byMonthStatus[chave] = { SOLICITADO:0, APROVADO:0, PENDENTE:0, RECUSADO:0, ENTREGUE:0 };
    byMonthStatus[chave][status] = (byMonthStatus[chave][status] || 0) + 1;

    dow[data.getDay()]++;

    registros.push({
      nome,
      cpf: formatCPF(cpfStr),
      unidade:setor,
      beneficio:tipo,
      demanda:tipo,
      status,
      dataISO: toISODate_(data),
      dataBR: formatDateBR(data)
    });
  });

  const mesesLabels = Object.keys(meses).sort((a,b)=>{
    const [ma,aa] = a.split('/').map(Number);
    const [mb,ab] = b.split('/').map(Number);
    return new Date(aa,ma-1,1) - new Date(ab,mb-1,1);
  });
  const mesesData = mesesLabels.map(k=>meses[k]||0);

  const series = { SOLICITADO:[], APROVADO:[], PENDENTE:[], RECUSADO:[], ENTREGUE:[] };
  mesesLabels.forEach(lbl=>{
    const pack = byMonthStatus[lbl] || {};
    series.SOLICITADO.push((pack.SOLICITADO||0));
    series.APROVADO  .push((pack.APROVADO  ||0));
    series.PENDENTE  .push((pack.PENDENTE  ||0));
    series.RECUSADO  .push((pack.RECUSADO  ||0));
    series.ENTREGUE  .push((pack.ENTREGUE  ||0));
  });

  return {
    status: statusTotais,
    tipos,
    setores,
    meses, mesesLabels, mesesData,
    byMonthStatusLabels: mesesLabels,
    byMonthStatusSeries: series,
    dowLabels: ['Dom','Seg','Ter','Qua','Qui','Sex','Sáb'],
    dowData: dow,
    registros
  };
}

function payloadDashboardVazio_(){
  return {
    status:{ SOLICITADO:0, APROVADO:0, PENDENTE:0, RECUSADO:0, ENTREGUE:0 },
    tipos:{}, setores:{}, meses:{},
    mesesLabels:[], mesesData:[],
    byMonthStatusLabels:[],
    byMonthStatusSeries:{ SOLICITADO:[], APROVADO:[], PENDENTE:[], RECUSADO:[], ENTREGUE:[] },
    dowLabels:['Dom','Seg','Ter','Qua','Qui','Sex','Sáb'],
    dowData:[0,0,0,0,0,0,0],
    registros:[]
  };
}

/************************
 * E-MAIL (fila/retry)
 ************************/
function enviarEmailSetor(setor, pdf, nomeBeneficiario) {
  try { ensureOutboxSheet_(); ensureQueueTrigger_(); } catch (_){}

  const toSetor = obterEmailDoSetor_(setor);
  if (!toSetor) {
    Logger.log('[EMAIL] Setor sem e-mail: ' + setor);
    return;
  }

  const assunto = 'Nova Ficha de Benefício - ' + (nomeBeneficiario || '');
  const corpo   = 'Segue em anexo a ficha preenchida do beneficiário: ' + (nomeBeneficiario || '');

  let fileId = null;
  try {
    if (pdf?.file?.getId) fileId = pdf.file.getId();
    else if (pdf?.fileId) fileId = pdf.fileId;
  } catch(_){}

  if (!fileId && pdf && pdf.blob){
    try {
      const f = DriveApp.createFile(pdf.blob);
      fileId = f.getId();
      setPublicSharing_(f);
    } catch(e){
      Logger.log('[EMAIL] blob->file: '+e.message);
    }
  }
  if (!fileId) {
    Logger.log('[EMAIL] sem fileId');
    return;
  }

  const okSetor = tentarEnviarAgora_(toSetor, assunto, corpo, fileId, pdf && pdf.blob);
  if (!okSetor) enqueueEmail_(toSetor, assunto, corpo, fileId, setor, nomeBeneficiario);

  if (COPIA_EMAIL && validarEmail_(COPIA_EMAIL)) {
    const okCopia = tentarEnviarAgora_(COPIA_EMAIL, assunto + ' (cópia)', corpo, fileId, null);
    if (!okCopia) enqueueEmail_(COPIA_EMAIL, assunto + ' (cópia)', corpo, fileId, setor, nomeBeneficiario);
  }
}

function validarEmail_(e){
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(e||'').trim());
}

/** ✅ (AJUSTADO) pega e-mail usando cabeçalho, não posição fixa */
function obterEmailDoSetor_(setor){
  try{
    const planilha = SpreadsheetApp.openById(PLANILHA_ID);
    const aba = planilha.getSheetByName('Login');
    if (!aba) return '';

    const lastRow = aba.getLastRow();
    const lastCol = aba.getLastColumn();
    if (lastRow < 2) return '';

    const headers = aba.getRange(1,1,1,lastCol).getValues()[0] || [];
    const idx = getLoginIdx_(headers);

    const iSet = idx.setor >= 0 ? idx.setor : 0;
    const iEm  = idx.email >= 0 ? idx.email : 3;

    const dados = aba.getRange(2,1,lastRow-1,lastCol).getValues();
    const wantedKey = canonicalSectorKey_(setor);

    for (let i=0;i<dados.length;i++){
      const setLabel = cleanSectorLabel_(dados[i][iSet] || '');
      if (canonicalSectorKey_(setLabel) !== wantedKey) continue;

      const em = (dados[i][iEm] || '').toString().trim();
      if (validarEmail_(em)) return em;
    }
  }catch(e){
    Logger.log('[EMAIL] obter email setor: ' + e.message);
  }
  return '';
}

function tentarEnviarAgora_(to, subject, body, fileId, blobOpt){
  if (!validarEmail_(to)) {
    Logger.log('[EMAIL] inválido: ' + to);
    return false;
  }
  const quota = MailApp.getRemainingDailyQuota();
  if (quota <= 0) {
    Logger.log('[EMAIL] Sem quota diária');
    return false;
  }

  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e){
    Logger.log('[EMAIL] lock falhou: ' + e.message);
  }

  try {
    for (let tent=1; tent<=2; tent++){
      try{
        const attachments = [];
        if (blobOpt) attachments.push(blobOpt);
        else attachments.push(DriveApp.getFileById(fileId).getAs(MimeType.PDF));

        if (tent>1) Utilities.sleep(600 * tent);

        try{
          MailApp.sendEmail({ to, subject, body, attachments, name:'SEMFAS Sistema' });
          Logger.log('[EMAIL] MailApp OK -> ' + to);
          return true;
        }catch(e1){
          Logger.log('[EMAIL] MailApp falhou: ' + e1.message);
          GmailApp.sendEmail(to, subject, body, { attachments, name:'SEMFAS Sistema' });
          Logger.log('[EMAIL] GmailApp OK -> ' + to);
          return true;
        }
      }catch(e){
        Logger.log('[EMAIL] tentativa ' + tent + ': ' + e.message);
        if (tent===2) return false;
      }
    }
  } finally {
    try { lock.releaseLock(); } catch(_){}
  }
  return false;
}

function ensureOutboxSheet_(){
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  let sh = ss.getSheetByName(OUTBOX_SHEET_NAME);
  if (!sh){
    sh = ss.insertSheet(OUTBOX_SHEET_NAME);
    sh.getRange(1,1,1,10).setValues([[
      'Timestamp','To','Assunto','Corpo','FileId','Tentativas','Status','UltimoErro','Setor','Beneficiario'
    ]]);
  }
  return sh;
}

function enqueueEmail_(to, subject, body, fileId, setor, beneficiario){
  const sh = ensureOutboxSheet_();
  sh.appendRow([
    new Date(), to, subject, body, fileId, 0, 'PENDING', '', setor||'', beneficiario||''
  ]);
  ensureQueueTrigger_();
  Logger.log('[EMAIL] enfileirado -> ' + to);
}

function processEmailQueue(){
  const sh = ensureOutboxSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const rng = sh.getRange(2,1,lastRow-1,10).getValues();
  const out = [];
  const quota = MailApp.getRemainingDailyQuota();
  if (quota <= 0) {
    Logger.log('[EMAIL] Sem quota diária');
    return;
  }

  for (let i=0;i<rng.length;i++){
    let [ts, to, subject, body, fileId, tries, status, lastErr, setor, benef] = rng[i];
    if (status === 'SENT') {
      out.push(rng[i]);
      continue;
    }
    if (tries >= OUTBOX_MAX_TRIES) {
      out.push([ts,to,subject,body,fileId,tries,'ERROR',lastErr,setor,benef]);
      continue;
    }

    let ok=false, errMsg='';
    try{
      const file = DriveApp.getFileById(fileId);
      const attachments = [ file.getAs(MimeType.PDF) ];
      try {
        MailApp.sendEmail({ to, subject, body, attachments, name:'SEMFAS Sistema' });
        ok=true;
      } catch(e1){
        GmailApp.sendEmail(to, subject, body, { attachments, name:'SEMFAS Sistema' });
        ok=true;
      }
    }catch(e){
      ok=false;
      errMsg = e.message || String(e);
    }

    if (ok) out.push([ts,to,subject,body,fileId,tries+1,'SENT','',setor,benef]);
    else    out.push([ts,to,subject,body,fileId,tries+1,'PENDING',errMsg,setor,benef]);
  }
  sh.getRange(2,1,out.length,10).setValues(out);
}

function ensureQueueTrigger_(){
  const fn = 'processEmailQueue';
  const triggers = ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===fn);
  if (triggers.length===0)
    ScriptApp.newTrigger(fn).timeBased().everyMinutes(5).create();
}

function extractDriveIdFromUrl_(url){
  if (!url) return '';
  const m1 = String(url).match(/\/d\/([a-zA-Z0-9_-]{10,})/);
  if (m1 && m1[1]) return m1[1];
  const m2 = String(url).match(/[?&]id=([a-zA-Z0-9_-]{10,})/);
  if (m2 && m2[1]) return m2[1];
  return '';
}

/************************
 * DIAGNÓSTICO RÁPIDO
 ************************/
function baixa_ping(){
  const { sh, idx } = getRespostasIndexMap_();
  return {
    sheet: sh.getName(),
    rows: sh.getLastRow() - 1,
    cols: sh.getLastColumn(),
    idx: idx
  };
}

/** =========================
 *  PROTOCOLO AUTOMÁTICO
 *  =========================*/
function nextProtocolo_(){
  const year = new Date().getFullYear();
  const key = 'PROTO_SEQ_' + year;
  const sp = PropertiesService.getScriptProperties();
  let n = parseInt(sp.getProperty(key) || '0', 10);
  n++;
  sp.setProperty(key, String(n));
  return `BEV-${year}-${String(n).padStart(6,'0')}`;
}

function ensureProtocoloForRow_(sh, rowIndex){
  const iProt = ensureColumn_(sh, PROTOCOLO_HDR);
  const val = sh.getRange(rowIndex, iProt+1).getValue();
  if (!String(val||'').trim()){
    const proto = nextProtocolo_();
    sh.getRange(rowIndex, iProt+1).setValue(proto);
    return proto;
  }
  return val;
}

function backfillProtocolos(){
  const { sh } = getRespostasIndexMap_();
  const iProt = ensureColumn_(sh, PROTOCOLO_HDR);
  const last = sh.getLastRow();
  if (last < 2) return { preenchidos: 0 };
  const rng = sh.getRange(2, iProt+1, last-1, 1);
  const vals = rng.getValues();
  let count = 0;
  for (let i=0;i<vals.length;i++){
    if (!String(vals[i][0]||'').trim()){
      vals[i][0] = nextProtocolo_();
      count++;
    }
  }
  rng.setValues(vals);
  return { preenchidos: count };
}

/************************
 * DETALHES (Drawer Central)
 ************************/
function getRegistroCompleto(linha){
  if (!linha) throw new Error('Linha não informada.');
  const { sh, headers, idx } = getRespostasIndexMap_();
  const lastCol = sh.getLastColumn();
  const row = sh.getRange(linha, 1, 1, lastCol).getValues()[0];

  const toStr = v => {
    if (v instanceof Date) return formatDateBR(v);
    if (typeof v === 'number') return String(v);
    return v == null ? '' : String(v);
  };

  const cols = [];
  for (let i=0;i<lastCol;i++){
    const label = headers[i] ? String(headers[i]) : ('Coluna ' + (i+1));
    cols.push({ label, value: toStr(row[i]) });
  }

  const iProt = findColIdx_(headers,
    PROTOCOLO_HDR,'protocolo','nº protocolo','n° protocolo',
    'numero do protocolo','id','id protocolo'
  );
  const protocolo = iProt >= 0 ? toStr(row[iProt]) : '';

  const status = isRowEntregue_(row, idx)
    ? 'ENTREGUE'
    : (String(idx.status>=0 ? row[idx.status] : 'PENDENTE').toUpperCase() || 'PENDENTE');

  const resumo = {
    linha,
    data: (function(){
      const d = coerceSheetDate(row[idx.data]); return d ? formatDateBR(d) : '';
    })(),
    unidade:  idx.unidade  >= 0 ? toStr(row[idx.unidade])  : '',
    demanda:  idx.beneficio>= 0 ? toStr(row[idx.beneficio]): '',
    nome:     idx.nome     >= 0 ? toStr(row[idx.nome])     : '',
    cpf:      idx.cpf      >= 0 ? formatCPF(toStr(row[idx.cpf])) : '',
    status,
    protocolo
  };

  return { linha, resumo, cols };
}

/************************
 * ATUALIZAR STATUS (Central)
 ************************/
function atualizarStatus(linha, status){
  status = String(status||'').trim().toUpperCase();
  if (!linha || !status) throw new Error('Parâmetros inválidos.');

  const { sh, idx } = getRespostasIndexMap_();

  let iStatus = idx.status;
  if (iStatus < 0) iStatus = ensureColumn_(sh, 'Status');

  const iEnt    = idx.entregueFlag;
  const iEntEm  = ensureColumn_(sh, ENTREGUE_EM_HDR);
  const iEntPor = ensureColumn_(sh, ENTREGUE_POR_HDR);
  const iEntUni = ensureColumn_(sh, ENTREGUE_UNID_HDR);
  const iEntObs = ensureColumn_(sh, ENTREGUE_OBS_HDR);

  const rowVals = sh.getRange(linha, 1, 1, sh.getLastColumn()).getValues()[0];

  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try{
    sh.getRange(linha, iStatus+1).setValue(status);
    if (status === 'ENTREGUE'){
      sh.getRange(linha, iEnt+1).setValue('ENTREGUE');
      if (!rowVals[iEntEm]) sh.getRange(linha, iEntEm+1).setValue(new Date());
    } else {
      sh.getRange(linha, iEnt+1).setValue('');
    }
  } finally { lock.releaseLock(); }
  return true;
}

/************************
 * PARECER TÉCNICO — atualizar/ler
 ************************/
function atualizarParecer(linha, payload){
  if (!linha) throw new Error('Linha não informada.');
  payload = payload || {};
  const { sh } = getRespostasIndexMap_();

  const iPar  = ensureColumn_(sh, 'PARECER TÉCNICO');
  const iNomT = ensureColumn_(sh, 'NOME TÉCNICO');
  const iMun  = ensureColumn_(sh, 'MUNICÍPIO PARECER');
  const iData = ensureColumn_(sh, 'DATA PARECER');
  const iAssB = ensureColumn_(sh, 'ASSINATURA BENEFICIÁRIO');
  const iAssT = ensureColumn_(sh, 'ASSINATURA/CARIMBO');

  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    if (payload.parecer != null)
      sh.getRange(linha, iPar+1).setValue(String(payload.parecer||''));
    if (payload.nome_tecnico != null)
      sh.getRange(linha, iNomT+1).setValue(String(payload.nome_tecnico||''));
    if (payload.municipio_parecer != null)
      sh.getRange(linha, iMun+1).setValue(String(payload.municipio_parecer||''));
    if (payload.data_parecer != null){
      const d = parseISODateSafe(payload.data_parecer);
      sh.getRange(linha, iData+1).setValue(d ? formatDateBR(d) : '');
    }
    if (payload.assinatura_beneficiario != null)
      sh.getRange(linha, iAssB+1).setValue(String(payload.assinatura_beneficiario||''));
    if (payload.assinatura_carimbo != null)
      sh.getRange(linha, iAssT+1).setValue(String(payload.assinatura_carimbo||''));
  } finally {
    lock.releaseLock();
  }
  return true;
}

function lerParecer(linha){
  if (!linha) throw new Error('Linha não informada.');
  const { sh } = getRespostasIndexMap_();

  const iPar  = ensureColumn_(sh, 'PARECER TÉCNICO');
  const iNomT = ensureColumn_(sh, 'NOME TÉCNICO');
  const iMun  = ensureColumn_(sh, 'MUNICÍPIO PARECER');
  const iData = ensureColumn_(sh, 'DATA PARECER');
  const iAssB = ensureColumn_(sh, 'ASSINATURA BENEFICIÁRIO');
  const iAssT = ensureColumn_(sh, 'ASSINATURA/CARIMBO');

  const vals = sh.getRange(linha, 1, 1, sh.getLastColumn()).getValues()[0];

  return {
    parecer:                 vals[iPar]  || '',
    nome_tecnico:            vals[iNomT] || '',
    municipio_parecer:       vals[iMun]  || '',
    data_parecer:            vals[iData] || '',
    assinatura_beneficiario: vals[iAssB] || '',
    assinatura_carimbo:      vals[iAssT] || ''
  };
}

/************************
 * PROTOCOLO / PDF — utilidades
 ************************/
function obterLinkPdf(linha){
  const { sh } = getRespostasIndexMap_();
  if (!linha) throw new Error('Linha não informada.');
  const iPdf = ensureColumn_(sh, 'PDF');
  const link = sh.getRange(linha, iPdf+1).getValue();
  return { link: link || '' };
}

function regerarPdf(linha){
  if (!linha) throw new Error('Linha não informada.');
  return gerarPdfDeRegistro(linha);
}

function garantirProtocolosPreenchidos(){
  return backfillProtocolos();
}

/************************
 * ENTREGA — desfazer/ajustar
 ************************/
function desfazerEntrega(linha){
  if (!linha) throw new Error('Linha não informada.');
  const { sh, idx } = getRespostasIndexMap_();

  const iEnt   = idx.entregueFlag;
  const iEntEm = ensureColumn_(sh, ENTREGUE_EM_HDR);
  const iEntPor= ensureColumn_(sh, ENTREGUE_POR_HDR);
  const iEntUni= ensureColumn_(sh, ENTREGUE_UNID_HDR);
  const iEntObs= ensureColumn_(sh, ENTREGUE_OBS_HDR);

  let iStatus  = idx.status;
  if (iStatus < 0) iStatus = ensureColumn_(sh, 'Status');

  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try{
    sh.getRange(linha, iEnt+1).setValue('');
    sh.getRange(linha, iEntEm+1).setValue('');
    sh.getRange(linha, iEntPor+1).setValue('');
    sh.getRange(linha, iEntUni+1).setValue('');
    sh.getRange(linha, iEntObs+1).setValue('');
    sh.getRange(linha, iStatus+1).setValue('PENDENTE');
  } finally {
    lock.releaseLock();
  }
  return true;
}

/************************
 * HISTÓRICO por CPF
 ************************/
function historicoPorCPF(cpf){
  const docpf = normalizeCPF(cpf);
  if (!docpf) return { registros: [] };

  const { sh, headers, idx } = getRespostasIndexMap_();
  const vals = sh.getDataRange().getValues();

  const iProt   = findColIdx_(headers,
    PROTOCOLO_HDR,'protocolo','nº protocolo','n° protocolo',
    'numero do protocolo','id','id protocolo'
  );
  const iEntEm  = ensureColumn_(sh, ENTREGUE_EM_HDR);
  const iEntPor = ensureColumn_(sh, ENTREGUE_POR_HDR);
  const iEntUni = ensureColumn_(sh, ENTREGUE_UNID_HDR);
  const iEntObs = ensureColumn_(sh, ENTREGUE_OBS_HDR);
  const iPdf    = idx.pdf >= 0 ? idx.pdf : ensureColumn_(sh, 'PDF');
  const iDocs   = idx.docs;

  const out = [];
  for (let i = 1; i < vals.length; i++){
    const row = vals[i];

    const cpfCell = (idx.cpf >= 0 && idx.cpf < row.length) ? row[idx.cpf] : '';
    const cpfRow  = normalizeCPF(cpfCell);
    if (!cpfRow || cpfRow !== docpf) continue;

    const dSolic  =
      (idx.data >= 0 && idx.data < row.length)
        ? coerceSheetDate(row[idx.data])
        : null;

    const dEntRaw =
      (iEntEm >= 0 && iEntEm < row.length)
        ? row[iEntEm]
        : null;
    const dEnt = dEntRaw ? parseAnyDate(dEntRaw) : null;

    const statusCell =
      (idx.status >= 0 && idx.status < row.length)
        ? row[idx.status]
        : 'PENDENTE';

    const status = isRowEntregue_(row, idx)
      ? 'ENTREGUE'
      : String(statusCell || 'PENDENTE').toUpperCase();

    out.push({
      linha: i+1,
      protocolo:       (iProt >=0   && iProt   < row.length) ? (row[iProt]      || '') : '',
      nome:            (idx.nome>=0 && idx.nome< row.length) ? (row[idx.nome]   || '') : '',
      unidade:         (idx.unidade>=0 && idx.unidade<row.length) ? (row[idx.unidade] || '') : '',
      beneficio:       (idx.beneficio>=0 && idx.beneficio<row.length) ? (row[idx.beneficio] || '') : '',
      data:            dSolic ? formatDateBR(dSolic) : '',
      status,
      entregue_em:     dEnt ? Utilities.formatDate(dEnt, TZ, 'dd/MM/yyyy HH:mm') : '',
      unidade_entrega: (iEntUni>=0 && iEntUni<row.length) ? (row[iEntUni] || '') : '',
      entregue_por:    (iEntPor>=0 && iEntPor<row.length) ? (row[iEntPor] || '') : '',
      obs_entrega:     (iEntObs>=0 && iEntObs<row.length) ? (row[iEntObs] || '') : '',
      pdf:             (iPdf   >=0 && iPdf   <row.length) ? (row[iPdf]    || '') : '',
      docs:            (iDocs  >=0 && iDocs  <row.length) ? (row[iDocs]   || '') : ''
    });
  }

  sortList_(out, 'data_desc');
  return { registros: out };
}

/************************
 * EXPORTAÇÃO CSV (link Drive)
 ************************/
function exportarCsvDashboard(params){
  params = params || {};
  const statusFilter   = params.status || 'todos';
  const unidadeFiltro  = params.unidade || '';
  const beneficioFiltro= params.beneficio || '';
  const order          = params.order || 'data_desc';
  const pdfOnly        = !!params.pdfOnly;
  const hasProto       = !!params.hasProto;
  const dataInicio     = params.dataInicio || '';
  const dataFim        = params.dataFim || '';

  const pack = baixa_list(statusFilter, unidadeFiltro, beneficioFiltro, order, pdfOnly, hasProto, dataInicio, dataFim);
  const registros = pack.registros || [];
  const sep = ';';

  const header = [
    'Linha','Protocolo','Nome','CPF','Benefício','Unidade','Solicitado em','Status','PDF'
  ];
  const linhas = [header.join(sep)];
  registros.forEach(r=>{
    linhas.push([
      r.rowRef, r.protocolo, r.nome, r.cpf, r.beneficio, r.unidade, r.solicitadoEm, r.status, r.pdf
    ].map(x => (String(x||'').includes(sep)
      ? `"${String(x).replace(/"/g,'""')}"`
      : String(x||''))).join(sep));
  });

  const blob = Utilities.newBlob(linhas.join('\n'), 'text/csv', 'export-semfas.csv');
  const file = DriveApp.createFile(blob);
  setPublicSharing_(file);
  return { fileId: file.getId(), link: buildDriveLinks_(file.getId()).download };
}

/************************
 * SAÚDE / DIAGNÓSTICO
 ************************/
function healthCheck(){
  const info = baixa_ping();
  return {
    ok: true,
    sheet: info.sheet,
    rows: info.rows,
    cols: info.cols,
    idx: info.idx,
    tz: TZ,
    now: Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss")
  };
}

function versaoScript(){
  return { version: 'v2026.01.06-FIXES-ROLEVIEW-DUPLICIDADE-EMAILHDR', planilha: PLANILHA_ID };
}

/************************
 * MENU (opcional no Editor)
 ************************/
function onOpen(){
  try{
    SpreadsheetApp.getUi()
      .createMenu('SEMFAS')
      .addItem('Backfill Protocolos','garantirProtocolosPreenchidos')
      .addItem('Health Check','healthCheck')
      .addItem('Processar Fila de E-mails','processEmailQueue')
      .addToUi();
  }catch(_){}
}

/************************
 * CORREÇÃO DE STATUS DA FILA (se necessário)
 ************************/
function corrigirStatusOutbox(){
  const sh = ensureOutboxSheet_();
  const lr = sh.getLastRow();
  if (lr < 2) return { corrigidos: 0 };
  const rng = sh.getRange(2,1,lr-1,10);
  const vals = rng.getValues();
  let c=0;
  for (let i=0;i<vals.length;i++){
    if (vals[i][6] === 'SENTE'){
      vals[i][6] = 'SENT';
      c++;
    }
  }
  rng.setValues(vals);
  return { corrigidos: c };
}

/************************
 * LOGIN — CRUD DE USUÁRIOS (Tela ANALISTA)
 ************************/
function ensureLoginSheet_(){
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  let sh = ss.getSheetByName('Login');
  if (!sh){
    sh = ss.insertSheet('Login');
    sh.appendRow(['Setor','Usuário','Senha','E-mail','Perfil']);
  }
  return sh;
}

function listarUsuariosLogin(){
  const sh = ensureLoginSheet_();
  const vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return { usuarios: [], setores: [] };

  const usuarios = [];
  const setSet = new Set();

  for (let i=1;i<vals.length;i++){
    const row = vals[i];
    const setorRaw = row[0] || '';
    const usuarioRaw = row[1] || '';
    const emailRaw = row[3] || '';
    const roleRaw  = row[4] || 'usuario';

    const setor = cleanSectorLabel_(setorRaw);
    const usuario = String(usuarioRaw||'').trim();
    const email   = String(emailRaw||'').trim();
    let role      = String(roleRaw||'usuario').trim().toLowerCase() || 'usuario';

    if (setor) setSet.add(setor);

    usuarios.push({
      linha: i+1,
      setor,
      usuario,
      email,
      role
    });
  }

  const setores = Array.from(setSet).sort((a,b)=>a.localeCompare(b,'pt-BR',{sensitivity:'base'}));
  return { usuarios, setores };
}

function criarUsuarioLogin(payload){
  payload = payload || {};
  let setor   = cleanSectorLabel_(payload.setor || '');
  let usuario = String(payload.usuario || '').trim();
  let senha   = String(payload.senha   || '').toString();
  let email   = String(payload.email   || '').trim();
  let role    = String(payload.role    || 'usuario').trim().toLowerCase() || 'usuario';

  if (!setor || !usuario || !senha){
    return { ok:false, msg:'Setor, usuário e senha são obrigatórios.' };
  }

  const sh = ensureLoginSheet_();
  const vals = sh.getDataRange().getValues();
  const setorKey = canonicalSectorKey_(setor);
  const userKey  = canonicalSectorKey_(usuario);

  for (let i=1;i<vals.length;i++){
    const sRow = cleanSectorLabel_(vals[i][0] || '');
    const uRow = String(vals[i][1] || '');
    if (canonicalSectorKey_(sRow) === setorKey &&
        canonicalSectorKey_(uRow) === userKey){
      return { ok:false, msg:'Já existe um usuário com esse login para este setor.' };
    }
  }

  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try{
    sh.appendRow([setor, usuario, senha, email, role]);
  } finally {
    lock.releaseLock();
  }
  return { ok:true };
}

function atualizarUsuarioLogin(payload){
  payload = payload || {};
  const linha = parseInt(payload.linha, 10);
  if (!linha || linha < 2){
    return { ok:false, msg:'Linha inválida.' };
  }

  let setor   = cleanSectorLabel_(payload.setor || '');
  let usuario = String(payload.usuario || '').trim();
  let email   = String(payload.email   || '').trim();
  let role    = String(payload.role    || 'usuario').trim().toLowerCase() || 'usuario';

  if (!setor || !usuario){
    return { ok:false, msg:'Setor e usuário são obrigatórios.' };
  }

  const sh = ensureLoginSheet_();
  const lastRow = sh.getLastRow();
  if (linha > lastRow){
    return { ok:false, msg:'Linha fora do intervalo da planilha.' };
  }

  const vals = sh.getDataRange().getValues();
  const setorKey = canonicalSectorKey_(setor);
  const userKey  = canonicalSectorKey_(usuario);

  for (let i=1;i<vals.length;i++){
    const rowIndex = i+1;
    if (rowIndex === linha) continue;
    const sRow = cleanSectorLabel_(vals[i][0] || '');
    const uRow = String(vals[i][1] || '');
    if (canonicalSectorKey_(sRow) === setorKey &&
        canonicalSectorKey_(uRow) === userKey){
      return { ok:false, msg:'Já existe um usuário com esse login para este setor.' };
    }
  }

  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    sh.getRange(linha,1).setValue(setor);
    sh.getRange(linha,2).setValue(usuario);
    sh.getRange(linha,4).setValue(email);
    sh.getRange(linha,5).setValue(role);
  } finally {
    lock.releaseLock();
  }
  return { ok:true };
}

function alterarSenhaLogin(linha, novaSenha){
  const row = parseInt(linha, 10);
  if (!row || row < 2){
    return { ok:false, msg:'Linha inválida.' };
  }
  novaSenha = String(novaSenha || '').toString();
  if (!novaSenha){
    return { ok:false, msg:'Senha não pode ser vazia.' };
  }

  const sh = ensureLoginSheet_();
  const lastRow = sh.getLastRow();
  if (row > lastRow){
    return { ok:false, msg:'Linha fora do intervalo da planilha.' };
  }

  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    sh.getRange(row,3).setValue(novaSenha);
  } finally {
    lock.releaseLock();
  }
  return { ok:true };
}

function resetarSenhaLogin(linha){
  const row = parseInt(linha, 10);
  if (!row || row < 2){
    return { ok:false, msg:'Linha inválida.' };
  }
  const sh = ensureLoginSheet_();
  const lastRow = sh.getLastRow();
  if (row > lastRow){
    return { ok:false, msg:'Linha fora do intervalo da planilha.' };
  }

  const nova = 'S' + Math.floor(100000 + Math.random()*900000);

  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    sh.getRange(row,3).setValue(nova);
  } finally {
    lock.releaseLock();
  }
  return { ok:true, senha:nova };
}

function excluirUsuarioLogin(linha){
  const row = parseInt(linha, 10);
  if (!row || row < 2){
    return { ok:false, msg:'Linha inválida.' };
  }
  const sh = ensureLoginSheet_();
  const lastRow = sh.getLastRow();
  if (row > lastRow){
    return { ok:false, msg:'Linha fora do intervalo da planilha.' };
  }

  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    sh.deleteRow(row);
  } finally {
    lock.releaseLock();
  }
  return { ok:true };
}

/************************
 * CHAMADOS TI
 ************************/
const CHAMADOS_SHEET_NAME = 'Chamados';
const CHAMADOS_SUBFOLDER_NAME = 'CHAMADOS_TI';
const CHAMADOS_MAX_BYTES = 5 * 1024 * 1024; // 5MB

function ensureChamadosSheet_(){
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  let sh = ss.getSheetByName(CHAMADOS_SHEET_NAME);
  if (!sh){
    sh = ss.insertSheet(CHAMADOS_SHEET_NAME);
    sh.appendRow([
      'Carimbo','Protocolo',
      'Nome','Email','Telefone',
      'Setor/Local','Categoria','Prioridade',
      'Descrição',
      'Status','Responsável','Atualizado em','Obs',
      'Anexo (Link)','Anexo (Nome)','Anexo (Mime)',
      'Origem','UserAgent'
    ]);
  } else {
    const must = [
      'Carimbo','Protocolo','Nome','Email','Telefone','Setor/Local','Categoria','Prioridade',
      'Descrição','Status','Responsável','Atualizado em','Obs',
      'Anexo (Link)','Anexo (Nome)','Anexo (Mime)','Origem','UserAgent'
    ];
    must.forEach(h => ensureColumn_(sh, h));
  }
  return sh;
}

function getChamadosFolder_(){
  const root = DriveApp.getFolderById(FOLDER_PDFS_ID);
  return ensureSubfolder_(root, CHAMADOS_SUBFOLDER_NAME);
}

function nextChamadoProtocolo_(){
  const year = new Date().getFullYear();
  const key = 'TI_SEQ_' + year;
  const sp = PropertiesService.getScriptProperties();
  let n = parseInt(sp.getProperty(key) || '0', 10);
  n++;
  sp.setProperty(key, String(n));
  return `TI-${year}-${String(n).padStart(6,'0')}`;
}

function approxBytesFromBase64_(b64){
  if (!b64) return 0;
  const s = String(b64).trim();
  const clean = s.replace(/^data:([\w\/\-\+\.]+);base64,/, '');
  return Math.floor(clean.length * 0.75);
}

function salvarAnexoChamado_(attachment, protocolo){
  if (!attachment) return { link:'', name:'', mime:'' };

  let fileName = attachment.fileName || attachment.name || 'anexo';
  let mimeType = attachment.mimeType || attachment.type || MimeType.BINARY;
  let base64   = attachment.base64 || attachment.data || attachment.dataUrl || '';

  if (!base64) return { link:'', name:'', mime:'' };

  const m = String(base64).match(/^data:([\w\/\-\+\.]+);base64,(.+)$/);
  if (m){
    mimeType = m[1] || mimeType;
    base64   = m[2] || '';
  }

  const bytesApprox = approxBytesFromBase64_(base64);
  if (bytesApprox > CHAMADOS_MAX_BYTES){
    throw new Error('Anexo acima de 5MB.');
  }

  const bytes = Utilities.base64Decode(base64);
  const blob  = Utilities.newBlob(bytes, mimeType, fileName);

  const folder = getChamadosFolder_();
  const safeProto = (protocolo || 'TI').toString().trim();
  const finalName = `${safeProto} - ${fileName}`.slice(0, 180);

  const file = withRetry(() => folder.createFile(blob).setName(finalName), 3, 300);
  setPublicSharing_(file);

  const link = buildDriveLinks_(file.getId()).preview;
  return { link, name:file.getName(), mime:mimeType };
}

function ti_abrirChamado(payload){
  payload = payload || {};

  const nome  = String(payload.requester_name  || payload.nome  || '').trim();
  const email = String(payload.requester_email || payload.email || '').trim();
  const tel   = String(payload.requester_phone || payload.tel   || '').trim();
  const dep   = String(payload.department      || payload.setor || '').trim();
  const cat   = String(payload.category        || '').trim();
  const pri   = String(payload.priority        || '').trim();
  const desc  = String(payload.description     || payload.mensagem || '').trim();

  if (!nome || !email || !tel || !dep || !cat || !pri || !desc){
    return { ok:false, msg:'Campos obrigatórios do chamado não preenchidos.' };
  }

  const sh = ensureChamadosSheet_();
  const protocolo = nextChamadoProtocolo_();

  let anexo = { link:'', name:'', mime:'' };
  if (payload.attachment){
    anexo = salvarAnexoChamado_(payload.attachment, protocolo);
  }

  let user = '';
  try {
    const up = PropertiesService.getUserProperties().getProperties();
    user = up.semfas_usuario || '';
  } catch(_){}

  const row = [
    new Date(), protocolo,
    nome, email, tel,
    dep, cat, pri,
    desc,
    'ABERTO', user, new Date(), '',
    anexo.link, anexo.name, anexo.mime,
    String(payload.pagina || 'login'), String(payload.userAgent || '')
  ];

  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    sh.appendRow(row);
  } finally {
    lock.releaseLock();
  }

  return {
    ok:true,
    protocolo,
    msg:'Chamado aberto com sucesso!',
    anexo: anexo.link || ''
  };
}

function ti_listarChamados(params){
  params = params || {};
  const status = String(params.status || 'ABERTO').trim().toUpperCase();
  const q = String(params.q || '').trim().toLowerCase();

  const sh = ensureChamadosSheet_();
  const vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return { chamados: [] };

  const out = [];
  for (let i=1;i<vals.length;i++){
    const r = vals[i];
    const protocolo = String(r[1]||'').trim();
    const nome = String(r[2]||'').trim();
    const dep  = String(r[5]||'').trim();
    const cat  = String(r[6]||'').trim();
    const pri  = String(r[7]||'').trim();
    const desc = String(r[8]||'').trim();
    const st   = String(r[9]||'ABERTO').trim().toUpperCase();
    const anexo= String(r[13]||'').trim();

    if (status && status !== 'TODOS' && st !== status) continue;

    if (q){
      const hay = (protocolo+' '+nome+' '+dep+' '+cat+' '+pri+' '+desc).toLowerCase();
      if (!hay.includes(q)) continue;
    }

    out.push({
      linha: i+1,
      carimbo: r[0] ? Utilities.formatDate(new Date(r[0]), TZ, 'dd/MM/yyyy HH:mm') : '',
      protocolo,
      nome,
      email: String(r[3]||''),
      telefone: String(r[4]||''),
      setor: dep,
      categoria: cat,
      prioridade: pri,
      descricao: desc,
      status: st,
      responsavel: String(r[10]||''),
      atualizado_em: r[11] ? Utilities.formatDate(new Date(r[11]), TZ, 'dd/MM/yyyy HH:mm') : '',
      obs: String(r[12]||''),
      anexo
    });
  }

  out.sort((a,b)=> (b.linha - a.linha));
  return { chamados: out.slice(0, 1500) };
}

function ti_atualizarChamado(linha, patch){
  const row = parseInt(linha,10);
  if (!row || row < 2) throw new Error('Linha inválida.');

  patch = patch || {};
  const sh = ensureChamadosSheet_();
  if (row > sh.getLastRow()) throw new Error('Linha fora do intervalo.');

  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try{
    if (patch.status != null) sh.getRange(row, 10).setValue(String(patch.status||'').toUpperCase()); // col J
    if (patch.responsavel != null) sh.getRange(row, 11).setValue(String(patch.responsavel||''));     // col K
    sh.getRange(row, 12).setValue(new Date());                                                       // col L
    if (patch.obs != null) sh.getRange(row, 13).setValue(String(patch.obs||''));                      // col M
  } finally {
    lock.releaseLock();
  }
  return { ok:true };
}

/************************
 * (FIM)
 ************************/
