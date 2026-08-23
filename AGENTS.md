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

---
