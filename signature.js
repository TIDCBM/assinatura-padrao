/*
 * Assinatura padrao DCBM - suplemento do Outlook
 *
 * Roda no momento em que a pessoa clica em "Novo email" ou "Responder",
 * le o e-mail dela, busca a assinatura correspondente no assinaturas.json
 * e insere usando setSignatureAsync.
 *
 * Funciona em: Outlook na Web, novo Outlook, Outlook classico (Windows) e Mac.
 * Nao pede login: o e-mail vem do proprio contexto da caixa.
 */

// ---------------------------------------------------------------------------
// AJUSTE AQUI se o repositorio ou a pasta mudarem
var FONTE = "https://tidcbm.github.io/assinatura-padrao/assinaturas.json";
// ---------------------------------------------------------------------------

var cache = null;

Office.onReady(function () {
  // nada a fazer na carga; a acao acontece nos eventos
});

function onNewMessageComposeHandler(event) {
  aplicarAssinatura(event);
}

function onMessageReplyComposeHandler(event) {
  aplicarAssinatura(event);
}

function aplicarAssinatura(event) {
  var email = "";
  try {
    email = (Office.context.mailbox.userProfile.emailAddress || "").toLowerCase();
  } catch (e) {
    event.completed();
    return;
  }

  if (!email) {
    event.completed();
    return;
  }

  buscarAssinaturas(function (dados) {
    if (!dados) {
      event.completed();
      return;
    }

    var html = dados[email];

    // se a pessoa nao estiver no arquivo, nao faz nada:
    // melhor sem assinatura do que com a assinatura de outro
    if (!html) {
      event.completed();
      return;
    }

    try {
      Office.context.mailbox.item.body.setSignatureAsync(
        html,
        { coercionType: Office.CoercionType.Html },
        function () {
          event.completed();
        }
      );
    } catch (e) {
      event.completed();
    }
  });
}

function buscarAssinaturas(callback) {
  if (cache) {
    callback(cache);
    return;
  }

  // ?v=... evita cache agressivo: republicou o JSON, vale no proximo email
  var url = FONTE + "?v=" + Math.floor(Date.now() / 60000);

  fetch(url, { cache: "no-store" })
    .then(function (r) {
      if (!r.ok) throw new Error("HTTP " + r.status);
      return r.json();
    })
    .then(function (dados) {
      // normaliza as chaves para minusculo
      var normalizado = {};
      for (var k in dados) {
        if (Object.prototype.hasOwnProperty.call(dados, k)) {
          normalizado[k.toLowerCase()] = dados[k];
        }
      }
      cache = normalizado;
      callback(normalizado);
    })
    .catch(function () {
      callback(null);
    });
}

// Registra os manipuladores. Os nomes precisam bater com os do manifest.xml.
if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  Office.actions.associate("onMessageReplyComposeHandler", onMessageReplyComposeHandler);
}
