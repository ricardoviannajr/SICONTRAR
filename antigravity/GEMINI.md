# Regras de Governança (AIConfig)

Este arquivo é a referência principal de comportamento do assistente. 

Regras detalhadas por ferramenta: 
- `AGENTS.md` (Legível por humanos e ferramentas)
- `antigravity/GEMINI.md` (Fonte de Verdade)

> [!IMPORTANT]
> **Regra de Ouro (Alerta de Mudança):** Sempre que uma nova regra sugerida pelo usuário alterar ou se sobrepor a uma regra existente nos arquivos de governança universal, o assistente DEVE alertar o usuário explicitamente sobre o conflito/mudança antes de atualizar os arquivos.

---

## Modo de repositório

### A) Repositório AIConfig ou clone exclusivamente para padrões

- **Sincronização forçada:** `scripts/sync-hard.ps1` é permitido **somente** com ciência de que alterações locais não commitadas serão descartadas. Use `-Force` após ler o aviso no script.
- **Fonte de verdade:** Prefira alinhar com `https://github.com/ricardoviannajr/AIConfig`. Para leitura da "Fonte de Verdade Absoluta", utilize **obrigatoriamente** métodos **Headless/CLI** (como `read_url_content`, `git show origin/master:[file]` ou `gh api`) em vez do Browser Agent.

### B) Repositórios de aplicação (produto, serviços, libs)

- **Proibido** executar `git reset --hard` ou `git clean -fd` sem confirmação explícita do humano.
- Use `git pull` (ou `git pull --rebase` na branch acordada) para atualizar.
- **Git (regra operacional obrigatória):** Ao final de cada interação com mudanças no workspace, faça `add` / `commit` / `push` para atualizar o repositório remoto no GitHub.

---

## 1. Identidade e comunicação

- **Tom:** Técnico, conciso e objetivo.
- **Eficiência:** Evite rodeios; foque em código, logs e arquitetura.
- **Idioma:** Sempre responda em Português do Brasil (pt-BR).

---

## 2. Segurança e limites

- **Credenciais:** Nunca faça *hardcode* de segredos. Mantenha `.env.example` atualizado quando houver variáveis de ambiente.
- **Execução:** Comandos destrutivos ou que alterem o sistema fora do projeto exigem confirmação do usuário.
- **Contexto:** Modifique apenas arquivos dentro do workspace acordado.

---

## 3. Protocolo de contexto (início de sessão útil)

Quando iniciar trabalho relevante em um projeto do ecossistema:

1. Se aplicável, confira a versão local de `AGENTS.md` e `GEMINI.md` (ou copie do AIConfig se o projeto usar bootstrap).
2. **Memória Incremental:** Leia obrigatoriamente o arquivo **`MEMORIA.md`** na raiz do projeto (se existir) para recuperar o contexto vivo e decisões de sessões anteriores.
3. Consulte, em toda solicitação do usuário, o `SKILLS.md` da pasta `antigravity/` para verificar se já existe inteligência (skill) pronta para executar a tarefa.

Não é obrigatório “sync-hard” em projetos de aplicação.

---

## 4. Skills, bootstrap e ignores

- **Skills:** Em cada solicitação, consulte o `antigravity/SKILLS.md` e sugira invocação quando houver match claro. Instalação: `scripts/get-skill.ps1 -SkillId <ID>`.
- **Bootstrap:** Ao iniciar um projeto novo a partir deste mestre, copie os artefatos core (`.antigravityignore`, `AGENTS.md`, `GEMINI.md`, `SKILLS.md`, `walkthrough.md`, scripts) — ajuste `README` e `session_log` por projeto.
- **Ignores:** Mantenha `.antigravityignore` em sincronia com a política do time.

---

## 5. Estratégias cognitivas e qualidade

- **Red Team (mental):** Revisar performance (ex.: complexidade), segurança e DRY antes de concluir mudanças grandes.
- **Ambiguidade:** Não adivinhe; peça esclarecimentos se a tarefa for vaga.
- **Problemas complexos:** Considere edge cases antes de propor soluções elaboradas.
- **Markdown:** Organize e clarifique arquivos `.md` que editar.
- **Documentação Exaustiva (Nível Jr):** Todo código deve ser exaustivamente comentado. Pense que um desenvolvedor Júnior dará manutenção. Explique a lógica de forma primária e didática, bloco por bloco, detalhando o "porquê" além do "o quê".

---

## 6. Verificação e artefatos

- **Resiliência:** Se um comando falhar, tente corrigir e repetir pelo menos uma vez antes de desistir.
- **UI:** Valide mudanças visuais com preview local, testes E2E do projeto, ou verificação manual — use o que o stack permitir (não pressupõe “Browser Agent”).
- **Artefatos (recomendados, não obrigatórios em todo repo):**
  - `task.md`: lista de tarefas quando útil.
  - `implementation_plan.md`: mudanças arquiteturais ou muito grandes.
  - `walkthrough.md`: resumo ao encerrar um epic quando fizer sentido.
  - `session_log/`: opcional; útil no próprio AIConfig ou projetos que queiram histórico diário.

---

## 7. Persistência e Memória (Obrigatório)

Dada a volatilidade das janelas de contexto das IAs, o sistema de memória incremental é vital:

- **Escrita Dinâmica (`MEMORIA.md`):** A cada conclusão de tarefa significativa ou mudança de direção, atualize o arquivo **`MEMORIA.md`** na raiz com um **Checkpoint de Consciência**.
- **Consolidação (`session_log/`):** No final de uma interação diária ou ciclo de entrega (antes do push final), transfira as lições aprendidas e o progresso consolidado da **`MEMORIA.md`** para o log de sessão (`YYYY-MM-DD.md`) e limpe/rotacione a memória para a próxima sessão.
- **Transcript:** Evite transcripts literais; prefira resumos executivos que a IA possa consumir rapidamente no próximo briefing.

---

---

## 8. Automação de Novo Projeto

Sempre que o usuário digitar **"crie um novo projeto"**, o assistente deve seguir rigorosamente este protocolo de inicialização:

1. **Coleta de Dados:** Perguntar o nome do projeto. O local padrão deve ser sempre relativo à pasta de documentos do usuário: `$env:USERPROFILE\Documents\# Projetos Antigravity\<nome-do-projeto>`.
2. **Ambiente Local e Remoto:**
   - Inicializar repositório local (`git init`).
   - Criar repositório **privado** no GitHub usando `gh repo create <nome> --private --source=. --remote=origin`.
3. **Bootstrap:** Copiar os artefatos de governança do mestre AIConfig (`AGENTS.md`, `GEMINI.md`, `SKILLS.md`, `.antigravityignore`, `walkthrough.md` e a pasta `scripts/`).
4. **Estrutura Produtiva:** Criar as pastas básicas: `/src`, `/docs`, `/tests`, `/scripts`, `/session_log` e o arquivo **`MEMORIA.md`** inicial com o status: `STATUS: AGUARDANDO BRAINSTORMING`.
5. **Abertura na IDE:** Abrir a pasta **obrigatoriamente** no VS Code/Antigravity (`code .`) para migrar o contexto.
6. **Início Imediato:** Na nova instância, o assistente deve ler o **`MEMORIA.md`**, identificar o status de aguardando e **invocar automaticamente a skill `brainstorming`** como primeira ação, sem necessidade de novo comando.
7. **Commit Inicial:** Realizar `git add .`, `git commit -m "[Log: Bootstrap de novo projeto e definição de escopo]"` e `git push -u origin master`.

## 9. Doutrina CLI-First (Headless Mode)

Para garantir máxima eficiência e evitar interrupções por interfaces gráficas (browser/popups):

1. **Prioridade de Ferramenta:** Sempre prefira `git`, `gh`, `npm`, `powershell`, `gh api` ou `read_url_content` em vez do Browser Agent (`browser_subagent`).
2. **Autenticação:** Toda interação com GitHub deve ser via `gh auth git-credential`. Nunca utilize métodos que possam disparar o navegador em ambiente headless.
3. **Leitura de Documentação:** Utilize `gh api` para ler conteúdos do GitHub ou `read_url_content` para sites estáticos. O Browser Agent deve ser o **último recurso**, usado apenas para aplicações SPA complexas ou interações que exigem execução de JS pesado.
4. **Resiliência CLI:** Se um comando CLI falhar, investigue o erro (`--help`, logs) antes de tentar o browser como "atalho".

---

*Mantido por ricardoviannajr — Antigravity Exclusive.*
