# Secretário Jurídico — guia de instalação

O sistema tem duas partes:

1. **Backend no seu Google** (`Code.gs`), um script do Google Apps Script que:
   - cria a planilha "Secretário Jurídico — Agenda de Prazos e Audiências" no seu Drive, com as abas `Agenda`, `Resumo` (atrasados, próximos 30 dias e totais por mês), `Feriados` e `Config`;
   - cria no Google Agenda uma agenda própria chamada **Prazos e Audiências** e lança nela cada compromisso, com três lembretes: **2 dias antes, 1 dia antes e no próprio dia**;
   - envia todo dia, às 7h, um **e-mail-resumo** com o que está atrasado e o que vence hoje, amanhã e em 2 dias (às segundas-feiras, também o restante da semana);
   - lê a sua agenda principal para mostrar seus outros compromissos no app e apontar conflitos de horário.
2. **App no celular** (`index.html` e arquivos de apoio): você dita ou digita o compromisso, confere e salva. Ele traz agenda semanal, mensal e anual, a lista com busca por cliente ou processo e o painel de alertas.

Tempo estimado: 15 a 20 minutos, uma única vez.

---

## Parte 1 — Backend (no computador)

1. Acesse **script.google.com** e clique em **Novo projeto**.
2. Apague o conteúdo de `Código.gs` e cole todo o conteúdo de `Code.gs`.
3. Clique em **+ → HTML**, dê o nome `index` e cole o conteúdo de `index.html` (só é necessário para o modo B, mas não atrapalha).
4. Salve e renomeie o projeto para "Secretário Jurídico".
5. Escolha a função **setup** e clique em **Executar**. Autorize o acesso: conta → "Avançado" → "Acessar Secretário Jurídico (não seguro)" → **Permitir**. O aviso aparece porque o script é seu e não foi verificado pelo Google. As permissões pedidas são Planilhas, Agenda, envio de e-mail e execução agendada, todas necessárias para as funções acima.
6. No **Registro de execução** aparecem o link da planilha e o **TOKEN** (que também fica na aba `Config`, célula B8).
7. **Implantar → Nova implantação → App da Web**: Executar como **Eu**; Quem pode acessar **Qualquer pessoa**. Copie a URL que termina em `/exec`.

Sem o token, ninguém lê nem grava nada. Se ele vazar, execute `trocarToken` e atualize o app.

Ao alterar o código no futuro: **Implantar → Gerenciar implantações → editar → Nova versão** (a URL não muda).

## Parte 2 — Notificações no celular

Os alertas chegam pelo aplicativo **Google Agenda** do Android:

1. Instale ou abra o Google Agenda no celular, com a mesma conta.
2. Menu (☰) → role até a sua conta e confirme que **Prazos e Audiências** está marcada. Se não aparecer, vá em Configurações → sua conta → "Prazos e Audiências" → ative **Sincronizar**.
3. Em Configurações do Android → Apps → Google Agenda → Notificações: deixe **permitidas**.

O e-mail das 7h é um segundo canal: se a notificação da agenda for perdida, o resumo chega pelo Gmail.

## Parte 3 — Instalar o app

### Modo A (recomendado): app instalável com microfone

1. Crie uma conta gratuita em **github.com** → **New repository** → nome `prazos`, **Public** → **Create repository**.
2. **uploading an existing file** → envie `index.html`, `manifest.webmanifest`, `sw.js`, `icon-192.png`, `icon-512.png` → **Commit changes**.
3. **Settings → Pages** → Branch `main`, pasta `/ (root)` → **Save**. Em cerca de um minuto aparece o endereço `https://SEU-USUARIO.github.io/prazos/`.
4. No celular, abra o endereço no **Chrome** → menu (⋮) → **Instalar app**.
5. No app, **Ajustes** → cole a URL `/exec` e o token → **Salvar e testar**. Depois toque em **Enviar e-mail de teste**.

O GitHub guarda só o código do app. Nomes de clientes, processos e o token ficam no seu celular e na sua conta Google.

Se você já publicou o Assessor Financeiro, pode usar a mesma conta do GitHub: basta criar outro repositório.

### Modo B (sem publicar): pela URL do Apps Script

Abra a URL `/exec` no Chrome do celular → (⋮) → **Adicionar à tela inicial** → em Ajustes, informe o token. Nesse modo o Google bloqueia o botão de microfone; dite pelo **microfone do teclado** no campo de texto.

---

## Como falar com o app

Diga, em qualquer ordem: **ato processual**, **cliente** e **data** (com hora, se houver). Exemplos:

- "Audiência de instrução da cliente Maria José dia 15 de outubro às 14h na 2ª Vara do Trabalho"
- "Contestação do cliente João Silva prazo de 15 dias úteis"
- "Resposta ao evento 32 cliente Paulo José até sexta-feira"
- "Aditamento da inicial cliente Raissa até dia 12/11"
- "Sessão do júri cliente Kelvyn dia 10/12 às 8h30"

Dizer "cliente" antes do nome melhora o reconhecimento. Também são entendidos: "amanhã", "depois de amanhã", "hoje", dias da semana, "dia 20" (o próximo dia 20), "às 2 da tarde", "quatorze e trinta" e o número do processo no padrão CNJ.

Antes de salvar, o app mostra uma ficha para conferência. Ali você corrige qualquer campo e usa a calculadora de prazo (número de dias + data de início + úteis/corridos).

## Contagem de prazos: o que o app faz e o que você precisa conferir

Quando a frase traz "em N dias" ou "prazo de N dias", o app calcula a data assim:

- exclui o dia do começo e inclui o do vencimento (CPC, art. 224);
- em **dias úteis** (CPC, art. 219; CLT, art. 775), pula fins de semana, feriados e, se ativado, o período de 20/12 a 20/01 (CPC, art. 220; CLT, art. 775-A);
- em **dias corridos** (CPP, art. 798), conta todos os dias e, se o vencimento cair em dia não útil, prorroga para o primeiro dia útil seguinte (CPP, art. 798, § 3º).

Feriados já considerados: os nacionais (inclusive 20/11), Carnaval (segunda e terça), Sexta-feira Santa, Corpus Christi, 19/03 (São José) e 25/03 (Data Magna do Ceará). **Não estão incluídos**: feriados municipais, pontos facultativos e suspensões de expediente do TJCE, TRT-7 ou TRF-5. Lance esses dias na aba `Feriados` da planilha; o app passa a considerá-los.

O termo inicial padrão é **a data em que você dita**. Se a intimação ou publicação ocorreu em outro dia, ajuste o campo "Contagem a partir de" na ficha. A contagem automática ajuda a organizar a agenda, mas não substitui a conferência do prazo no sistema do tribunal (PJe, eproc, SAJ), sobretudo diante de prazos em dobro, intimação eletrônica com prazo de ciência ou regras do JEC e do processo penal.
