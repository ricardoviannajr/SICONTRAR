# Agentes Especializados e Padrões de Código (AIConfig)

Este arquivo é a visão consolidada para humanos e ferramentas. 

**Governança Geral** está em [`Core.md`](Core.md).

> [!IMPORTANT]
> **Regra de Ouro:** Alertar o usuário antes de alterar/sobrepor regras existentes nos arquivos de governança.

---

## 1. Perfis de Agentes Especializados

- **Principal:** Orquestração, padrões e governança.
- **Senior Fullstack:** Arquitetura, performance e DRY.
- **Red Team Reviewer:** Auditoria e edge cases.

---

## 2. Padrões de Qualidade de Código

- **Tamanho de arquivo:** Máximo de 300 linhas.
- **Tamanho de função:** Máximo de 30 linhas.
- **Complexidade:** Limite ciclomático de 10.
- **Exportações:** Prefira `named exports`.
- **Indentação:** 2 espaços.
- **Documentação:** O código deve ser autoexplicativo (Clean Code). Comentários devem existir apenas para justificar o "porquê" de decisões complexas ou não intuitivas (nível Pleno/Sênior). Evite documentação bloco-a-bloco.
- **README.md:** Todo repositório DEVE ter um arquivo `README.md` de alta qualidade (premium), atualizado conforme o progresso do projeto.

---

## 3. Requisitos de Teste

- Escreva testes unitários para funções utilitárias.
- Cobertura mínima de 80% em código novo.

---

## 4. Tratamento de erros e validação

- **Validação:** Valide todos os inputs do usuário antes do processamento.
- **Tratamento:** Envolva chamadas de API externas em try/catch com tratamento de erro adequado.
- **Logs técnicos:** Use o logger estruturado para depuração, evitando `console.log`.

## 5. Regras de Máquina (rules/)

As regras de automação e comportamento estão versionadas em `rules/` (arquivos `.mdc`) e devem ser seguidas rigorosamente.

## 6. Regras Operacionais Obrigatórias

- **Auto-Pilot:** Após a aprovação inicial de um plano ou tarefa, o Colaborador deve agir de forma autônoma, realizando todas as edições, comandos e commits necessários sem interromper o usuário para confirmações intermediárias. Priorize as ferramentas nativas de arquivos do Agente para evitar popups de autorização da IDE.
- **Git:** Ao final de cada interação, `add` / `commit` / `push`. Mandatório, pré-autorizado, e executado via terminal com `SafeToAutoRun: true`.
- **Tooling Selection:** Ferramentas Nativas > CLI > Browser. Prefira ferramentas nativas do Agente (`write_to_file`, `replace_file_content`, etc.) para edição de arquivos; ferramentas CLI (`git`, `gh`, etc.) para controle de versão; e o Browser apenas como último recurso.
- **Skills:** Consultar `SKILLS.md` in toda nova solicitação.
- **Memória:** Atualizar **`MEMORIA.md`** em checkpoints de consciência. **Obrigatório:** O arquivo deve ser um **Resumo Executivo Conciso**. Registre o status atual, decisões chave (sem snippets de código) e próximos passos em bullet points curtos para economizar contexto e garantir continuidade eficiente.
- **Higiene de Contexto (Tokens):** Após atualizar a **`MEMORIA.md`**, o assistente DEVE sugerir o encerramento da sessão ou limpeza de contexto se o consumo de tokens estiver elevado, garantindo que o próximo ciclo inicie com o "briefing" destilado.

- **Invisibilidade (Regra 11):** É proibido deixar qualquer rastro de origem não-humana no código ou documentação. Substitua termos como "Assistente", "Bot" ou "Módulo de IA" por "Desenvolvedor" ou "Padrão do Sistema".

---
