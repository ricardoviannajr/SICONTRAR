# Regras de Governança (AIConfig)

Este arquivo é a referência principal de comportamento do assistente. 

Regras detalhadas: 
- `AGENTS.md` (Padrões de Código)
- `GEMINI.md` (Governança e Protocolos)
- `SKILLS.md` (Catálogo de Habilidades)

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
- **Execução:**
    - Comandos destrutivos ou que alterem o sistema fora do projeto exigem confirmação do usuário.
    - Comandos Git (`add`, `commit`, `push`) estão **pré-autorizados** e não exigem confirmação.
- **Contexto:** Modifique apenas arquivos dentro do workspace acordado.

---

## 3. Protocolo de contexto (início de sessão útil)

Quando iniciar trabalho relevante em um projeto do ecossistema:

1. Se aplicável, confira a versão local de `AGENTS.md` e `GEMINI.md` (ou copie do AIConfig se o projeto usar bootstrap).
2. **Memória Incremental:** Leia obrigatoriamente o arquivo **`MEMORIA.md`** na raiz do projeto (se existir) para recuperar o contexto vivo e decisões de sessões anteriores.
3. Consulte, em toda solicitação do usuário, o `SKILLS.md` na raiz do projeto para verificar se já existe inteligência (skill) pronta para executar a tarefa.

Não é obrigatório “sync-hard” em projetos de aplicação.

---

## 4. Skills, bootstrap e ignores

- **Skills:** Em cada solicitação, consulte o `SKILLS.md` (raiz) e sugira invocação quando houver match claro. Instalação: `scripts/get-skill.ps1 -SkillId <ID>`.
- **Bootstrap:** Ao iniciar um projeto novo a partir deste mestre, copie os artefatos core (`.antigravityignore`, `AGENTS.md`, `GEMINI.md`, `SKILLS.md`, `walkthrough.md`, scripts) — ajuste `README` e `session_log` por projeto.
- **Ignores:** Mantenha `.antigravityignore` em sincronia com a política do time.

---

## 5. Estratégias cognitivas e qualidade

- **Red Team (mental):** Revisar performance (ex.: complexidade O(n)), segurança e DRY antes de concluir mudanças grandes.
- **Ambiguidade:** Não adivinhe; peça esclarecimentos se a tarefa for vaga.
- **Chain of Thought:** Detalhe desafios e edge cases antes de propor soluções complexas. Raciocínio explícito reduz erros.
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

## 10. Integridade Estrutural de Projeto (Project Lockdown)

Todo projeto no ecossistema Antigravity deve manter conformidade com a estrutura base.
O assistente é responsável por verificar e alertar — nunca por ignorar silenciosamente.

### Artefatos Obrigatórios (todo projeto, sem exceção)

| Arquivo | Propósito | Criado em |
|---------|-----------|-----------|
| `GEMINI.md` | Governança (cópia do mestre ou stub) | Bootstrap |
| `AGENTS.md` | Padrões de código | Bootstrap |
| `SKILLS.md` | Catálogo de skills disponíveis | Bootstrap |
| `MEMORIA.md` | Contexto vivo entre sessões | Bootstrap |
| `README.md` | Documentação premium do projeto | Bootstrap |
| `.antigravityignore` | Política de arquivos ignorados | Bootstrap |

### Regras de Enforcement

1. **Check de Integridade:** No início de qualquer interação com um projeto, o assistente deve verificar silenciosamente a presença dos 6 artefatos obrigatórios. Se algum estiver ausente, alertar o usuário ANTES de iniciar a tarefa solicitada.
2. **MEMORIA.md é sagrada:** Toda interação que altere código, arquitetura ou decisões técnicas DEVE terminar com atualização do `MEMORIA.md` usando o template padrão (`MEMORIA_TEMPLATE.md`). Sem checkpoint = interação incompleta.
3. **Proibido desviar do padrão base:** Nenhum projeto pode ter uma versão customizada de `GEMINI.md` ou `AGENTS.md` que contradiga a fonte mestre (AIConfig), exceto adições locais explicitamente marcadas como `## [LOCAL]`.
4. **Stubs são transitórios:** Arquivos com `# DEPRECADO (stub temporário)` devem ser resolvidos (substituídos pelo conteúdo real) na próxima interação com o projeto.
5. **Propagação Obrigatória:** Toda alteração em `GEMINI.md` ou `AGENTS.md` no AIConfig (fonte da verdade) **deve ser propagada imediatamente** para todos os destinos. Sem propagação = alteração incompleta.

### Protocolo de Propagação (Obrigatório)

Sempre que `GEMINI.md` ou `AGENTS.md` forem alterados no repositório AIConfig:

1. **Commit e Push no AIConfig:** Confirmar a alteração na fonte da verdade primeiro.
2. **Replicar para o Global:** Copiar o arquivo atualizado para `~/.gemini/` (regras globais do sistema).
   ```powershell
   Copy-Item "AIConfig\GEMINI.md" "$env:USERPROFILE\.gemini\GEMINI.md" -Force
   ```
3. **Propagar para todos os projetos:** Iterar sobre todos os repositórios em `$env:USERPROFILE\Documents\# Projetos Antigravity\` e sobrescrever o `GEMINI.md` (ou `AGENTS.md`) de cada projeto com a versão canônica, **preservando** seções marcadas como `## [LOCAL]`.
   ```powershell
   # Exemplo de propagação via PowerShell:
   Get-ChildItem "$env:USERPROFILE\Documents\# Projetos Antigravity" -Directory |
     Where-Object { $_.Name -ne "AIConfig" } |
     ForEach-Object {
       $dest = Join-Path $_.FullName "GEMINI.md"
       if (Test-Path $dest) {
         Copy-Item "AIConfig\GEMINI.md" $dest -Force
         # Commit automático no projeto destino
         git -C $_.FullName add GEMINI.md
         git -C $_.FullName commit -m "[Sync: GEMINI.md atualizado via AIConfig]"
         git -C $_.FullName push origin master 2>$null
       }
     }
   ```
4. **Validação:** Confirmar que o número de projetos atualizados corresponde ao total esperado. Reportar quaisquer falhas de push.

> [!IMPORTANT]
> **Seções `## [LOCAL]`** em projetos individuais devem ser preservadas durante a propagação. O assistente deve fazer merge inteligente: substituir o conteúdo canônico e **anexar** as seções locais ao final do arquivo.

---

*Mantido por ricardoviannajr — Antigravity Exclusive.*
