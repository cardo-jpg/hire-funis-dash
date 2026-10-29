/**
 * Backend da hire-funis-dash — Anotações da equipe + Chat IA (proxy seguro).
 * A dash (GitHub Pages) NÃO pode guardar a chave da Anthropic (página pública),
 * então ela fala com ESTE Apps Script, que guarda a chave do lado servidor.
 *
 * ───────── ONDE ESTÁ PUBLICADO ─────────
 * Precisa ser numa conta @gmail.com PESSOAL. O Workspace @cardomarketing.com.br
 * NÃO libera acesso anônimo ("Qualquer pessoa") — força login no domínio, e aí
 * a dash pública não consegue ler/gravar. (Testado: o /exec da agência redireciona
 * pro accounts.google.com com hd=cardomarketing.com.br.)
 *
 * Publicado como projeto CONTAINER-BOUND de uma planilha nova na conta pessoal:
 *   Planilha nova → Extensões → Apps Script → cola este código → Ctrl+S
 *   Implantar → Nova implantação → App da Web → Executar como "Eu",
 *   Acesso "Qualquer pessoa" → Implantar.
 * URL boa começa com script.google.com/macros/s/...  (SEM /a/macros/<dominio>/).
 * A cada mudança no código: Implantar → Gerenciar implantações → editar → Nova versão.
 *
 * (Só pro chat) Configurações do projeto → Propriedades do script →
 *   ANTHROPIC_KEY = sk-ant-...   Sem ela, as notas funcionam e o chat diz "não configurado".
 *
 * Como é container-bound, usa getActiveSpreadsheet() — a aba "Notas"
 * (Timestamp | Dashboard | Autor | Nota) fica na própria planilha pessoal.
 * A planilha de dados da dash (1l6_bsucWh3C...) NÃO é tocada nem compartilhada.
 */

var ABA_NOTAS = 'Notas';
var MODELO_IA = 'claude-haiku-5-5';   // barato e rápido pro Q&A da dash

function _notasSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ABA_NOTAS);
  if (!sh) { sh = ss.insertSheet(ABA_NOTAS); sh.appendRow(['Timestamp', 'Dashboard', 'Autor', 'Nota']); }
  return sh;
}

function _json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** GET: ?action=ping  |  ?action=notes&dash=<id> */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === 'ping') return _json_({ ok: true });
  if (p.action === 'notes') {
    var sh = _notasSheet_();
    var dash = String(p.dash || '').trim();
    var rows = sh.getDataRange().getValues();
    var out = [];
    for (var i = 1; i < rows.length; i++) {
      if (!dash || String(rows[i][1]) === dash) {
        out.push({ ts: rows[i][0], autor: rows[i][2], nota: rows[i][3] });
      }
    }
    out.reverse(); // mais recentes primeiro
    return _json_({ notes: out });
  }
  return _json_({ ok: true, hint: 'use action=notes ou POST' });
}

/** POST (corpo JSON): {type:'note', dash, autor, nota}  |  {type:'chat', pergunta, contexto} */
function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return _json_({ erro: 'json invalido' }); }

  if (body.type === 'note') {
    var nota = String(body.nota || '').trim();
    if (!nota) return _json_({ erro: 'nota vazia' });
    _notasSheet_().appendRow([
      new Date(),
      String(body.dash || '').slice(0, 80),
      String(body.autor || 'equipe').slice(0, 60),
      nota.slice(0, 2000)
    ]);
    return _json_({ ok: true });
  }

  if (body.type === 'chat') {
    var key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_KEY');
    if (!key) return _json_({ resposta: 'Chat ainda nao configurado: falta a chave da Anthropic (ANTHROPIC_KEY) nas propriedades do script.' });
    var pergunta = String(body.pergunta || '').trim();
    if (!pergunta) return _json_({ erro: 'pergunta vazia' });
    var sys = 'Voce e um analista de trafego pago da Cardo Marketing. Responde em portugues do Brasil, '
            + 'direto e curto (no maximo uns 4 paragrafos), com base SO nos dados fornecidos da dashboard. '
            + 'Se o dado nao estiver nos numeros, diga que nao tem esse dado. Nada de travessao.';
    var payload = {
      model: MODELO_IA,
      max_tokens: 900,
      system: sys,
      messages: [{ role: 'user', content: 'Pergunta: ' + pergunta + '\n\nDados atuais da dashboard:\n' + String(body.contexto || '(sem dados)') }]
    };
    try {
      var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
        method: 'post', contentType: 'application/json',
        headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
        payload: JSON.stringify(payload), muteHttpExceptions: true
      });
      var data = JSON.parse(resp.getContentText());
      if (data.content && data.content[0] && data.content[0].text) return _json_({ resposta: data.content[0].text });
      return _json_({ resposta: 'A IA nao respondeu. ' + (data.error ? data.error.message : '') });
    } catch (err) {
      return _json_({ resposta: 'Erro ao falar com a IA: ' + err });
    }
  }

  return _json_({ erro: 'type desconhecido (use note ou chat)' });
}
