/**
 * KIT DO LUCRÉCIO — recebe os cadastros do formulário e grava numa planilha do seu Google Drive.
 * Cole este código em script.google.com (Novo projeto), salve e implante como "App da Web".
 * Na primeira vez, ele cria a planilha "Kit do Lucrécio — contatos" no seu Drive.
 */
const NOME_PLANILHA = 'Kit do Lucrécio — contatos';
const CABECALHO = ['Quando', 'Nome', 'Contato', 'Negócio', 'Faturamento', 'Veio de', 'Aceite LGPD'];

function planilha_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('PLANILHA_ID');
  if (id) { try { return SpreadsheetApp.openById(id).getSheets()[0]; } catch (e) { /* apagada: cria outra */ } }
  const ss = SpreadsheetApp.create(NOME_PLANILHA);
  const aba = ss.getSheets()[0];
  aba.setName('Contatos');
  aba.appendRow(CABECALHO);
  aba.getRange(1, 1, 1, CABECALHO.length).setFontWeight('bold').setBackground('#F2C94C');
  aba.setFrozenRows(1);
  props.setProperty('PLANILHA_ID', ss.getId());
  return aba;
}

const limpa_ = (v, max) => String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').trim().slice(0, max)
  .replace(/^[=+\-@]/, "'$&");   // evita que um texto vire fórmula na planilha

function doPost(e) {
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d.site) return resposta_('ok');                          // armadilha para robôs
    if (d.aceite !== true) return resposta_('sem aceite');        // LGPD: só grava com aceite
    const nome = limpa_(d.nome, 80), contato = limpa_(d.contato, 120);
    if (!nome || !contato) return resposta_('incompleto');
    const lock = LockService.getScriptLock(); lock.waitLock(10000);
    try {
      planilha_().appendRow([new Date(), nome, contato, limpa_(d.segmento, 60), limpa_(d.porte, 40), limpa_(d.origem, 40), 'sim']);
    } finally { lock.releaseLock(); }
    return resposta_('ok');
  } catch (err) {
    return resposta_('erro');
  }
}

function doGet() { return resposta_('Kit do Lucrécio: captação funcionando.'); }

function resposta_(t) { return ContentService.createTextOutput(t).setMimeType(ContentService.MimeType.TEXT); }

/** Rode esta função uma vez pelo editor (botão Executar) para autorizar e criar a planilha. */
function prepararPlanilha() {
  const aba = planilha_();
  Logger.log('Planilha pronta: ' + aba.getParent().getUrl());
}
