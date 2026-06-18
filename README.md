# Banco de Horas

> Trabalho a part-time no gabinete de informática da minha universidade com horário flexível e senti a necessidade de ter uma forma simples de registar as minhas horas diárias — para saber facilmente quantas horas tenho a mais ou a menos no final do mês.
>
> Sou estudante de Engenharia Informática no 1º ano e, com ajuda de IA (Claude da Anthropic), consegui criar esta aplicação do zero: uma web app completa construída em Google Apps Script + Google Sheets, sem servidores externos nem custos de infraestrutura.
>
> O projeto cresceu de um simples registo de horas para uma ferramenta com painel de administração, pedidos de folga, relatórios em PDF, envio automático de resumos por email, e muito mais.

---

Web app de registo e gestão de horas de trabalho, construída em **Google Apps Script + Google Sheets**.

Desenvolvida para uso interno — sem base de dados externa, sem servidor próprio. Tudo corre dentro do Google Workspace.

---

## Funcionalidades

- ✅ Registo diário de horas (até 3 períodos por dia)
- 📊 Saldo acumulado, estatísticas e gráfico donut
- 🏖️ Pedidos de folga com aprovação pelo admin
- 🔒 Admin pode trancar horários ou marcar folga obrigatória
- 💬 Comentários por dia (partilhados entre funcionário e admin)
- 📧 Resumo por email (semanal, mensal ou anual) com PDF em anexo
- 📄 Exportar relatório em TXT ou PDF (por período e funcionário)
- 👤 Painel admin com calendário, estatísticas e gestão de utilizadores
- 🔁 Desfazer/refazer alterações
- 📱 Interface responsiva (mobile)
- 🔐 Sessão persistente com `localStorage`

---

## Requisitos

- Conta Google (pessoal ou Workspace)
- Google Sheets
- Google Apps Script

---

## Instalação

### 1. Criar a Google Sheet

1. Cria uma nova Google Sheet em [sheets.google.com](https://sheets.google.com)
2. Abre o menu **Extensões → Apps Script**
3. Cola o conteúdo de `Code.gs` no editor
4. Cola o conteúdo de `index.html` num novo ficheiro HTML (`Ficheiro → Novo → Ficheiro HTML`)

### 2. Configurar

No início de `Code.gs`, ajusta as constantes:

```javascript
const START_DATE  = '2026-06-01';  // data de início do sistema (YYYY-MM-DD)
const DAILY_MINS  = 240;           // carga diária em minutos (ex: 240 = 4h, 480 = 8h)
```

Na função `initializeSheets`, substitui o email e password do admin inicial:

```javascript
[['admin@example.com', 'Administrador', 'changeme123', 'admin', true, '', '']]
```

> ⚠️ **Muda a password após o primeiro login** nas definições do perfil.

Se a tua cidade tem um feriado municipal específico, edita em `getPortugueseHolidays`:

```javascript
map[fmtKey(year, 6, 24)] = 'Feriado Municipal'; // ajusta a data e o nome
```

### 3. Inicializar

1. No Apps Script, corre a função `setupSpreadsheet` uma vez (menu **Executar**)
2. Aceita as permissões pedidas pelo Google
3. Implementa a app: **Implementar → Nova implementação → Aplicação Web**
   - Executar como: **Eu**
   - Quem tem acesso: **Qualquer pessoa** (ou só quem tem o link)
4. Copia o URL gerado — é o link da app

---

## Estrutura

```
Code.gs       — lógica de servidor (Google Apps Script)
index.html    — interface web (HTML + CSS + JS, num único ficheiro)
```

Os dados ficam guardados em folhas dentro da Google Sheet:

| Folha     | Conteúdo                          |
|-----------|-----------------------------------|
| Users     | Utilizadores e passwords          |
| Entries   | Registos diários de horas         |
| Comments  | Comentários por dia               |
| Closed    | Períodos fechados (pausas)        |
| Invited   | Emails convidados (pré-registo)   |

---

## Segurança

- Passwords guardadas em texto simples na Google Sheet (sem hash) — adequado para uso interno, não recomendado para dados sensíveis
- Acesso à Sheet controlado pelas permissões do Google Drive
- A app corre sob a conta do proprietário do script — sem exposição de credenciais externas

---

## O que aprendi ao longo do projeto

Este projeto foi o meu primeiro contacto real com desenvolvimento web e programação do lado do servidor. Aqui estão algumas das principais dificuldades que encontrei e o que aprendi com elas:

### Google Apps Script e Google Sheets como plataforma
Aprendi que o Google Apps Script corre no servidor do Google e comunica com a página web através de `google.script.run` — o que implica que cada chamada ao servidor demora 1 a 3 segundos. Perceber isto foi fundamental para entender porque a app era lenta e como otimizá-la.

Descobri também que o Google Sheets converte automaticamente valores de hora (ex: `"09:00"`) em objetos `Date`, o que causou um bug difícil de detetar: a resposta do servidor chegava `null` ao cliente. A solução foi formatar explicitamente todos os campos de hora antes de os devolver.

### Performance e otimização
Uma das maiores aprendizagens foi perceber que cada leitura de uma folha do Sheets é uma operação lenta. Ao início, o código lia a mesma folha 3 vezes para obter dados diferentes (entradas, hora mínima, hora máxima). Aprendi a consolidar essas leituras numa só função, o que reduziu significativamente o tempo de carregamento.

Aprendi também a usar o `CacheService` do Apps Script para guardar dados que raramente mudam (como a lista de utilizadores e pausas) durante 30 segundos, evitando releituras desnecessárias.

### Gestão de estado no frontend
Aprendi a diferença entre guardar dados no servidor (Google Sheets) e manter estado no cliente (variáveis JavaScript). Implementei `localStorage` para persistir a sessão do utilizador entre refreshes — sem isso, o utilizador teria de fazer login sempre que atualizasse a página.

### UX e decisões de design
Ao longo do projeto fui percebendo que pequenas decisões de interface têm grande impacto na experiência. Por exemplo:
- O auto-save (guardar automaticamente ao mudar cada hora) parecia conveniente mas tornava a app lenta — mudei para guardar só quando o utilizador clica "Guardar"
- Os botões de desfazer/refazer decorativos no painel do admin confundiam porque pareciam clicáveis mas não faziam nada — removi-os
- Os emojis nos botões do modal ficavam desalinhados — substituí por ícones SVG consistentes

### Funcionalidades mais complexas
- **Pedidos de folga com aprovação**: aprendi a gerir estados intermédios (`Pendente`, `Rejeitado`, `Aprovado`) tanto no servidor como no cliente, e a manter a interface sincronizada com esses estados
- **Exportação de PDF**: descobri que o `DocumentApp` do Apps Script não suporta todos os métodos de formatação que a documentação sugere (ex: `body.setMargins` não existe). A solução foi gerar HTML estilizado com CSS e exportá-lo como PDF via `DriveApp`
- **Envio de emails com HTML**: aprendi a usar `MailApp` para enviar emails formatados com tabelas e cores, e a anexar ficheiros PDF

### Segurança e boas práticas para publicação
Antes de publicar no GitHub, aprendi a importância de remover dados sensíveis do código (emails reais, passwords, nome da instituição) e substituí-los por placeholders genéricos, para que qualquer pessoa possa usar o projeto sem expor informação privada.
---

## Licença

MIT — uso livre, sem garantias.
