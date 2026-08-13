# Assinatura Padrão DCBM — Suplemento do Outlook

Aplica a assinatura padrão automaticamente em e-mails novos e respostas, no
**Outlook na Web, novo Outlook, Outlook clássico (Windows) e Outlook para Mac**.

Não usa regra de transporte (nada de rodapé no fim da thread) e não precisa de
GPO nem de mexer em estação nenhuma — está ligado à caixa de correio.

---

## Como funciona

1. A pessoa clica em **Novo email** ou **Responder**.
2. O suplemento dispara, lê o e-mail dela pelo contexto da caixa (sem login).
3. Busca esse e-mail no `assinaturas.json`.
4. Insere a assinatura com `setSignatureAsync` — no lugar certo, acima do histórico.

Quem não estiver no JSON simplesmente não recebe assinatura. Nunca recebe a de outro.

---

## Passo 1 — criar o repositório

Na conta **TIDCBM**, crie um repositório **público** chamado `assinatura-padrao`.
Precisa ser público: o Outlook carrega os arquivos sem autenticação.

Estrutura final:

```
assinatura-padrao/
├── manifest.xml
├── commands.html
├── signature.js
├── assinaturas.json        <- gerado pelo script PowerShell
└── img/
    ├── ic-phone.png
    ├── ic-mobile.png
    ├── ic-mail.png
    ├── ic-pin.png
    └── logo_simbolo_v1.png
```

As cinco imagens são as mesmas do repositório `Icones`. Copie para `img/` aqui —
o `Icones` continua existindo, intocado.

## Passo 2 — ligar o GitHub Pages

No repositório: **Settings → Pages → Source: Deploy from a branch → main → / (root) → Save**.

Isso é obrigatório. O `raw.githubusercontent.com` entrega HTML e JS como texto
puro e o navegador se recusa a executar; o Pages entrega com o tipo correto.

Confirme que estes três endereços abrem:

- https://tidcbm.github.io/assinatura-padrao/commands.html
- https://tidcbm.github.io/assinatura-padrao/signature.js
- https://tidcbm.github.io/assinatura-padrao/img/logo_simbolo_v1.png

## Passo 3 — gerar o assinaturas.json

Rode `Assinatura-Padrao-DCBM.ps1`, escolha o escopo e depois o botão
**"Gerar assinaturas.json para o suplemento"**. Suba o arquivo gerado para a raiz
do repositório.

## Passo 4 — implantar o suplemento

No **centro de administração do Microsoft 365** (admin.microsoft.com):

**Configurações → Aplicativos integrados → Carregar aplicativos personalizados**

- Tipo: **Suplemento do Office**
- Origem: **Carregar arquivo de manifesto** → envie o `manifest.xml`
- Atribuição: comece com **usuários específicos** (você e mais um) para testar;
  depois mude para **toda a organização**

A implantação leva de alguns minutos até 24 horas para chegar em todas as caixas.

---

## Manutenção

Mudou cargo, telefone ou celular de alguém no Entra ID? Rode o script de novo,
gere o `assinaturas.json` e suba. Todo mundo pega a versão nova no próximo e-mail
— sem tocar em estação, sem reinstalar nada.

Mudou o visual da assinatura? Edite `Assinatura-ComCelular.html` e
`Assinatura-SemCelular.html`, regenere o JSON e suba.

O `manifest.xml` só precisa ser reenviado se mudar URL, nome ou permissões.
Nesse caso, incremente a `<Version>`.

---

## Atenção: o JSON fica público

O `assinaturas.json` terá nome, cargo, telefone, ramal e e-mail de todos os
usuários, em um endereço público e indexável por buscadores.

São os mesmos dados que já saem no rodapé de cada e-mail enviado para fora, mas
aqui ficam reunidos em um único arquivo e fáceis de coletar em massa. Se isso for
um problema para a empresa, as alternativas são hospedar em um domínio próprio com
restrição de acesso, ou trocar o JSON por leitura via Microsoft Graph com SSO.

---

## Se não funcionar

| Sintoma | Onde olhar |
|---|---|
| Nada acontece ao criar e-mail | O suplemento chegou na caixa? Confira em Aplicativos integrados se a implantação concluiu |
| Funciona no Web mas não no clássico | Ativação por evento exige build recente do Microsoft 365 Apps |
| Assinatura sem imagens | Abra as URLs de `img/` no navegador; se der 404, o Pages não está ligado ou os arquivos não subiram |
| Só algumas pessoas recebem | Confira se o e-mail delas está no `assinaturas.json`, em minúsculas |
| Manifesto recusado no upload | Valide o XML e confira se todas as URLs respondem em HTTPS |
