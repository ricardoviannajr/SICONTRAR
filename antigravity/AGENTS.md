# Agentes Especializados e Padrões de Código (AIConfig)

Este arquivo é a visão consolidada para humanos e ferramentas. 

**Governança Geral** está em [`GEMINI.md`](GEMINI.md).

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
- **Documentação:** TODO CÓDIGO GERADO deve ser exaustivamente comentado (nível Jr). Explique a lógica de forma primária e didática, bloco por bloco, para que qualquer iniciante possa manter o sistema.
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

## 5. Regras de Máquina (antigravity/rules)

As regras de automação e comportamento estão versionadas em `antigravity/rules/` e devem ser seguidas rigorosamente.

## 4. Regras Operacionais Obrigatórias

- **Git:** Ao final de cada interação, `add` / `commit` / `push`.
- **Tooling Selection:** CLI over Browser. Prefira ferramentas de linha de comando (`gh`, `git`, `ps1`) para TODAS as tarefas, exceto validação de UI visual ou interações com SPAs dinâmicos.
- **Skills:** Consultar `SKILLS.md` em toda nova solicitação.
- **Memória:** Atualizar **`MEMORIA.md`** em checkpoints de consciência. **Obrigatório:** Detalhar decisões técnicas, obstáculos ("pulo do gato") e lições aprendidas, evitando resumos rasos de tarefas.


---
