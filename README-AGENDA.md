# Agenda automática — Fabiano Trompetista 13007

## O que esta integração faz

A equipe edita uma planilha do Google Sheets no Drive. O site continua usando os mesmos cards e a mesma identidade visual, mas passa a carregar os compromissos automaticamente.

- Compromissos com data passada deixam de aparecer automaticamente.
- Agendas fixas continuam visíveis.
- A coluna **Exibir** permite publicar ou ocultar um item sem apagá-lo.
- Os compromissos com data são ordenados cronologicamente.
- O site tenta carregar a versão online a cada acesso.
- Se a conexão falhar, usa a última agenda armazenada no navegador; se não houver cache, mantém os cards estáticos do HTML.
- O Apps Script usa cache de até 5 minutos para reduzir chamadas.

## 1. Criar a planilha no Google Drive

1. Acesse https://script.google.com/ com a conta Google que será responsável pela agenda.
2. Crie um **Novo projeto**.
3. Substitua o conteúdo de `Code.gs` pelo arquivo `google-apps-script/Code.gs` deste pacote.
4. Salve o projeto.
5. No seletor de funções, escolha `criarEstruturaAgendaFabiano`.
6. Clique em **Executar**.
7. Autorize o acesso ao Google Drive e ao Google Sheets.

O script criará:

- Pasta: **Agenda – Fabiano 13007**
- Planilha: **Agenda Pública – Fabiano 13007**
- Aba: **Agenda**
- Aba: **Instruções**

Os cinco compromissos fornecidos inicialmente já estarão cadastrados.

## 2. Implantar como Web App

No Apps Script:

1. Clique em **Implantar > Nova implantação**.
2. Selecione **App da Web**.
3. Em **Executar como**, escolha **Você**.
4. Em **Quem pode acessar**, escolha a opção que permita acesso público ao endpoint (inclusive visitantes não autenticados, quando essa opção estiver disponível na conta).
5. Clique em **Implantar**.
6. Copie a URL de produção que termina em `/exec`.

Não use a URL `/dev` no site; ela é destinada a testes.

## 3. Conectar o site

Abra `index.html` e localize:

```html
data-agenda-api-url="COLE_AQUI_A_URL_DO_WEB_APP_APPS_SCRIPT"
```

Troque somente o conteúdo entre aspas pela URL `/exec` obtida na implantação.

Exemplo:

```html
data-agenda-api-url="https://script.google.com/macros/s/SEU_ID/exec"
```

Depois publique `index.html`, `style.css` e `script.js` normalmente na hospedagem.

## 4. Como a equipe atualiza a agenda

Na aba **Agenda**, cada linha é um compromisso.

Colunas:

- **Exibir**: `Sim` ou `Não`.
- **Data**: para compromissos com data definida.
- **Dia da semana**: texto mostrado no site.
- **Horário**: aceita horários (`19:00`) ou períodos (`Manhã`).
- **Evento**: título do compromisso.
- **Local**: local exibido no card.
- **Tipo**: `Evento` ou `Agenda fixa`.
- **Recorrência**: por exemplo, `Todas as terças`.
- **Observações**: texto opcional.

### Compromisso normal
Preencha Data, Dia da semana, Horário, Evento e Local. Use `Evento` na coluna Tipo.

### Agenda fixa
Deixe a Data vazia, escolha `Agenda fixa` e informe a Recorrência.

### Ocultar um compromisso
Mude **Exibir** para `Não`. Não é necessário apagar a linha.

## 5. Atualização e cache

O site consulta a agenda ao ser aberto. O endpoint do Apps Script mantém um cache curto de até **5 minutos**. Se for necessário publicar uma alteração imediatamente, execute a função `limparCacheAgenda()` no Apps Script.

## 6. Transferência para a conta da campanha

Se o projeto e a planilha forem criados primeiro em uma conta pessoal e depois transferidos para outra conta ou domínio Google Workspace, faça uma nova implantação do Web App pela conta responsável final e atualize a URL `/exec` no `index.html`.

## Arquivos do site alterados

- `index.html`: adiciona a URL configurável do Web App e o indicador de sincronização.
- `script.js`: carrega e renderiza a agenda automaticamente.
- `style.css`: mantém a identidade visual e acrescenta estados de sincronização/indisponibilidade.
