/**
 * Ponte entre o app Streamlit "Anúncio CSC-PM" e a planilha.
 *
 * Instalação (uma vez):
 *  1. Na planilha: Extensões > Apps Script. Cole este arquivo inteiro em Code.gs.
 *  2. Configurações do projeto (engrenagem) > Propriedades do script >
 *     adicione TOKEN = <uma senha longa e aleatória>.
 *  3. Implantar > Nova implantação > tipo "App da Web":
 *       Executar como: Eu
 *       Quem pode acessar: Qualquer pessoa
 *     Copie a URL terminada em /exec.
 *  4. No Streamlit Cloud (Settings > Secrets) informe url e token (ver README).
 *
 * Ações (POST com corpo JSON):
 *  - {"token":..., "acao":"ler", "data":"dd/mm/aaaa"}
 *  - {"token":..., "acao":"gravar_afastamentos", "ops":[...]}
 */

var ABA_FORMULARIO = 'Respostas ao formulário 1';
var ABA_EFETIVO = 'EFETIVO CSC';
var ABA_AFASTAMENTOS = 'AFASTAMENTOS';
var CABECALHO_AFASTAMENTOS = ['MATRÍCULA', 'NOME', 'TIPO', 'DESCRIÇÃO', 'INÍCIO', 'FIM'];

function doPost(e) {
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var token = PropertiesService.getScriptProperties().getProperty('TOKEN');
    if (!token || req.token !== token) {
      return responder_({ ok: false, erro: 'Token inválido.' });
    }
    if (req.acao === 'ler') return responder_(ler_(req.data));
    if (req.acao === 'gravar_afastamentos') return responder_(gravarAfastamentos_(req.ops || []));
    return responder_({ ok: false, erro: 'Ação desconhecida: ' + req.acao });
  } catch (err) {
    return responder_({ ok: false, erro: String(err) });
  }
}

function doGet() {
  return responder_({ ok: false, erro: 'Use POST.' });
}

function responder_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function aba_(nome) {
  var ss = SpreadsheetApp.getActive();
  var alvo = String(nome).toLowerCase();
  var abas = ss.getSheets();
  for (var i = 0; i < abas.length; i++) {
    if (abas[i].getName().trim().toLowerCase() === alvo) return abas[i];
  }
  return null;
}

function valores_(sh) {
  if (!sh || sh.getLastRow() === 0 || sh.getLastColumn() === 0) return [];
  return sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
}

/** Lê efetivo, afastamentos e as respostas do formulário da data informada. */
function ler_(data) {
  var shForm = aba_(ABA_FORMULARIO);
  var shEfet = aba_(ABA_EFETIVO);
  if (!shForm) return { ok: false, erro: 'Aba não encontrada: ' + ABA_FORMULARIO };
  if (!shEfet) return { ok: false, erro: 'Aba não encontrada: ' + ABA_EFETIVO };

  var form = valores_(shForm);
  var formulario = form;
  if (data && form.length > 0) {
    var idx = form[0].map(function (h) { return String(h).trim(); }).indexOf('Data do anúncio');
    if (idx >= 0) {
      formulario = [form[0]].concat(form.slice(1).filter(function (r) {
        return String(r[idx]).trim().indexOf(data) === 0;
      }));
    }
  }
  return {
    ok: true,
    formulario: formulario,
    efetivo: valores_(shEfet),
    afastamentos: valores_(aba_(ABA_AFASTAMENTOS)),
  };
}

function abaAfastamentos_() {
  var sh = aba_(ABA_AFASTAMENTOS);
  if (sh) return sh;
  sh = SpreadsheetApp.getActive().insertSheet(ABA_AFASTAMENTOS);
  sh.getRange(1, 1, 1, CABECALHO_AFASTAMENTOS.length)
    .setValues([CABECALHO_AFASTAMENTOS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('@');
  sh.getRange('E:F').setNumberFormat('dd/mm/yyyy');
  return sh;
}

function paraData_(txt) {
  var m = String(txt).match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!m) throw new Error('Data inválida: ' + txt);
  return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
}

function linhaPlanilha_(v) {
  // v = [matricula, nome, tipo, descricao, inicio, fim]
  return [String(v[0]), String(v[1]), String(v[2]), String(v[3]), paraData_(v[4]), paraData_(v[5])];
}

/**
 * ops: [{op:"incluir", valores:[...]}, {op:"atualizar", linha:N, matricula:"...", valores:[...]}]
 * "atualizar" só ocorre se a matrícula da linha ainda for a esperada.
 */
function gravarAfastamentos_(ops) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = abaAfastamentos_();
    var resultados = [];
    ops.forEach(function (o) {
      if (o.op === 'atualizar') {
        var atual = sh.getRange(o.linha, 1).getDisplayValue();
        if (String(atual).trim() !== String(o.matricula).trim()) {
          resultados.push({ ok: false, erro: 'Linha ' + o.linha + ' mudou; não atualizada.' });
          return;
        }
        sh.getRange(o.linha, 1, 1, 6).setValues([linhaPlanilha_(o.valores)]);
        resultados.push({ ok: true, linha: o.linha });
      } else if (o.op === 'incluir') {
        var nova = sh.getLastRow() + 1;
        sh.getRange(nova, 1).setNumberFormat('@');
        sh.getRange(nova, 1, 1, 6).setValues([linhaPlanilha_(o.valores)]);
        sh.getRange(nova, 5, 1, 2).setNumberFormat('dd/mm/yyyy');
        resultados.push({ ok: true, linha: nova });
      }
    });
    SpreadsheetApp.flush();
    return { ok: true, resultados: resultados, afastamentos: valores_(sh) };
  } finally {
    lock.releaseLock();
  }
}
