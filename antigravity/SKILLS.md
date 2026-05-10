# 🌌 Antigravity Skills Catalog

---

## Uso no Antigravity

O instalador envia os arquivos para o destino certo conforme o parâmetro `-Target`.

| Destino | Caminho típico (Windows) | Comando |
| :--- | :--- | :--- |
| **Antigravity** | `<projeto>/.antigravity/skills/<ID>/` | `.\scripts\get-skill.ps1 -SkillId <ID>` |

- **Skills do projeto:** você pode copiar uma pasta `<ID>` com `SKILL.md` para `.antigravity/skills/<ID>/` no repositório da aplicação para versionar a skill junto do código.

---

## 📖 Sobre este Catálogo

> [!IMPORTANT]
> **Este é um Catálogo Federado.** Para manter sua IDE rápida e o contexto limpo, as skills listadas abaixo **não estão instaladas localmente**. Elas residem no repositório remoto `antigravity-awesome-skills`.

## 🚀 Como Utilizar

1. Procure a skill desejada no índice abaixo.
2. Solicite ao assistente: "**Invoque a skill [ID]**" (ou instale com `get-skill.ps1`).
3. O script baixa os arquivos para `.antigravity/skills/`.

---

## 🗂️ Índice de Categorias

- [AI & Agents](#ai-agents)
- [Automation & SaaS](#automation-saas)
- [Backend & Cloud](#backend-cloud)
- [Data & Analytics](#data-analytics)
- [Development Tools](#development-tools)
- [General & Productivity](#general-productivity)
- [Legal & Finance](#legal-finance)
- [Management & Strategy](#management-strategy)
- [Marketing & Product](#marketing-product)
- [Security & Privacy](#security-privacy)
- [Web & Frontend](#web-frontend)

---

## 📁 AI & Agents

| ID | Descrição |
| :--- | :--- |
| `acceptance-orchestrator` | Use when a coding task should be driven end-to-end from issue intake through implementation, review, deployment, and acceptance verification with m... |
| `activecampaign-automation` | Automate ActiveCampaign tasks via Rube MCP (Composio): manage contacts, tags, list subscriptions, automation enrollment, and tasks. Always search t... |
| `advanced-evaluation` | This skill should be used when the user asks to "implement LLM-as-judge", "compare model outputs", "create evaluation rubrics", "mitigate evaluatio... |
| `agent-evaluation` | Testing and benchmarking LLM agents including behavioral testing, capability assessment, reliability metrics, and production monitoring—where even ... |
| `agent-framework-azure-ai-py` | Build Azure AI Foundry agents using the Microsoft Agent Framework Python SDK (agent-framework-azure-ai). Use when creating persistent agents with A... |
| `agent-manager-skill` | Manage multiple local CLI agents via tmux sessions (start/stop/monitor/assign) with cron-friendly scheduling. |
| `agent-memory-mcp` | A hybrid memory system that provides persistent, searchable knowledge management for AI agents (Architecture, Patterns, Decisions). |
| `agent-memory-systems` | Memory is the cornerstone of intelligent agents. Without it, every interaction starts from zero. This skill covers the architecture of agent memory... |
| `agent-orchestration-improve-agent` | Systematic improvement of existing agents through performance analysis, prompt engineering, and continuous iteration. |
| `agent-orchestration-multi-agent-optimize` | Optimize multi-agent systems with coordinated profiling, workload distribution, and cost-aware orchestration. Use when improving agent performance,... |
| `agent-orchestrator` | Meta-skill que orquestra todos os agentes do ecossistema. Scan automatico de skills, match por capacidades, coordenacao de workflows multi-skill e ... |
| `agent-tool-builder` | Tools are how AI agents interact with the world. A well-designed tool is the difference between an agent that works and one that hallucinates, fail... |
| `agentfolio` | Skill for discovering and researching autonomous AI agents, tools, and ecosystems using the AgentFolio directory. |
| `agentmail` | Email infrastructure for AI agents. Create accounts, send/receive emails, manage webhooks, and check karma balance via the AgentMail API. |
| `agents-md` | This skill should be used when the user asks to "create AGENTS.md", "update AGENTS.md", "maintain agent docs", "set up CLAUDE.md", or needs to keep... |
| `agents-v2-py` | Build container-based Foundry Agents with Azure AI Projects SDK (ImageBasedHostedAgentDefinition). Use when creating hosted agents with custom cont... |
| `ai-agent-development` | AI agent development workflow for building autonomous agents, multi-agent systems, and agent orchestration with CrewAI, LangGraph, and custom agents. |
| `ai-agents-architect` | Expert in designing and building autonomous AI agents. Masters tool use, memory systems, planning strategies, and multi-agent orchestration. Use wh... |
| `ai-analyzer` | AI驱动的综合健康分析系统，整合多维度健康数据、识别异常模式、预测健康风险、提供个性化建议。支持智能问答和AI健康报告生成。 |
| `ai-engineer` | Build production-ready LLM applications, advanced RAG systems, and intelligent agents. Implements vector search, multimodal AI, agent orchestration... |
| `ai-ml` | AI and machine learning workflow covering LLM application development, RAG implementation, agent architecture, ML pipelines, and AI-powered features. |
| `ai-native-cli` | Design spec with 98 rules for building CLI tools that AI agents can safely use. Covers structured JSON output, error handling, input contracts, saf... |
| `ai-product` | Every product will be AI-powered. The question is whether you'll build it right or ship a demo that falls apart in production. This skill covers LL... |
| `ai-studio-image` | Geracao de imagens humanizadas via Google AI Studio (Gemini). Fotos realistas estilo influencer ou educacional com iluminacao natural e imperfeicoe... |
| `ai-wrapper-product` | Expert in building products that wrap AI APIs (OpenAI, Anthropic, etc.) into focused tools people will pay for. Not just 'ChatGPT but different' - ... |
| `airflow-dag-patterns` | Build production Apache Airflow DAGs with best practices for operators, sensors, testing, and deployment. Use when creating data pipelines, orchest... |
| `airtable-automation` | Automate Airtable tasks via Rube MCP (Composio): records, bases, tables, fields, views. Always search tools first for current schemas. |
| `amazon-alexa` | Integracao completa com Amazon Alexa para criar skills de voz inteligentes, transformar Alexa em assistente com Claude como cerebro (projeto Auri) ... |
| `andrej-karpathy` | Agente que simula Andrej Karpathy — ex-Director of AI da Tesla, co-fundador da OpenAI, fundador da Eureka Labs, e o maior educador de deep learning... |
| `antigravity-skill-orchestrator` | A meta-skill that understands task requirements, dynamically selects appropriate skills, tracks successful skill combinations using agent-memory-mc... |
| `api-documenter` | Master API documentation with OpenAPI 3.1, AI-powered tools, and modern developer experience practices. Create interactive docs, generate SDKs, and... |
| `apify-ultimate-scraper` | Universal AI-powered web scraper for any platform. Scrape data from Instagram, Facebook, TikTok, YouTube, Google Maps, Google Search, Google Trends... |
| `app-builder` | Main application building orchestrator. Creates full-stack applications from natural language requests. Determines project type, selects tech stack... |
| `appdeploy` | Deploy web apps with backend APIs, database, and file storage. Use when the user asks to deploy or publish a website or web app and wants a public ... |
| `architecture` | Architectural decision-making framework. Requirements analysis, trade-off evaluation, ADR documentation. Use when making architecture decisions or ... |
| `audio-transcriber` | Transform audio recordings into professional Markdown documentation with intelligent summaries using LLM integration |
| `auri-core` | Auri: assistente de voz inteligente (Alexa + Claude claude-opus-4-20250805). Visao do produto, persona Vitoria Neural, stack AWS, modelo Free/Pro/B... |
| `autonomous-agent-patterns` | Design patterns for building autonomous coding agents. Covers tool integration, permission systems, browser automation, and human-in-the-loop workf... |
| `autonomous-agents` | Autonomous agents are AI systems that can independently decompose goals, plan actions, execute tools, and self-correct without constant human guida... |
| `avalonia-viewmodels-zafiro` | Optimal ViewModel and Wizard creation patterns for Avalonia using Zafiro and ReactiveUI. |
| `awt-e2e-testing` | AI-powered E2E web testing — eyes and hands for AI coding tools. Declarative YAML scenarios, Playwright execution, visual matching (OpenCV + OCR), ... |
| `azure-ai-agents-persistent-dotnet` | Azure AI Agents Persistent SDK for .NET. Low-level SDK for creating and managing AI agents with threads, messages, runs, and tools. |
| `azure-ai-agents-persistent-java` | Azure AI Agents Persistent SDK for Java. Low-level SDK for creating and managing AI agents with threads, messages, runs, and tools. |
| `azure-ai-anomalydetector-java` | Build anomaly detection applications with Azure AI Anomaly Detector SDK for Java. Use when implementing univariate/multivariate anomaly detection, ... |
| `azure-ai-contentsafety-java` | Build content moderation applications with Azure AI Content Safety SDK for Java. Use when implementing text/image analysis, blocklist management, o... |
| `azure-ai-contentsafety-py` | Azure AI Content Safety SDK for Python. Use for detecting harmful content in text and images with multi-severity classification. |
| `azure-ai-contentsafety-ts` | Analyze text and images for harmful content using Azure AI Content Safety (@azure-rest/ai-content-safety). Use when moderating user-generated conte... |
| `azure-ai-contentunderstanding-py` | Azure AI Content Understanding SDK for Python. Use for multimodal content extraction from documents, images, audio, and video. |
| `azure-ai-document-intelligence-dotnet` | Azure AI Document Intelligence SDK for .NET. Extract text, tables, and structured data from documents using prebuilt and custom models. |
| `azure-ai-ml-py` | Azure Machine Learning SDK v2 for Python. Use for ML workspaces, jobs, models, datasets, compute, and pipelines. |
| `azure-ai-openai-dotnet` | Azure OpenAI SDK for .NET. Client library for Azure OpenAI and OpenAI services. Use for chat completions, embeddings, image generation, audio trans... |
| `azure-ai-projects-dotnet` | Azure AI Projects SDK for .NET. High-level client for Azure AI Foundry projects including agents, connections, datasets, deployments, evaluations, ... |
| `azure-ai-projects-java` | Azure AI Projects SDK for Java. High-level SDK for Azure AI Foundry project management including connections, datasets, indexes, and evaluations. |
| `azure-ai-projects-py` | Build AI applications using the Azure AI Projects Python SDK (azure-ai-projects). Use when working with Foundry project clients, creating versioned... |
| `azure-ai-projects-ts` | Build AI applications using Azure AI Projects SDK for JavaScript (@azure/ai-projects). Use when working with Foundry project clients, agents, conne... |
| `azure-ai-textanalytics-py` | Azure AI Text Analytics SDK for sentiment analysis, entity recognition, key phrases, language detection, PII, and healthcare NLP. Use for natural l... |
| `azure-ai-transcription-py` | Azure AI Transcription SDK for Python. Use for real-time and batch speech-to-text transcription with timestamps and diarization. |
| `azure-ai-translation-document-py` | Azure AI Document Translation SDK for batch translation of documents with format preservation. Use for translating Word, PDF, Excel, PowerPoint, an... |
| `azure-ai-translation-text-py` | Azure AI Text Translation SDK for real-time text translation, transliteration, language detection, and dictionary lookup. Use for translating text ... |
| `azure-ai-vision-imageanalysis-java` | Build image analysis applications with Azure AI Vision SDK for Java. Use when implementing image captioning, OCR text extraction, object detection,... |
| `azure-ai-vision-imageanalysis-py` | Azure AI Vision Image Analysis SDK for captions, tags, objects, OCR, people detection, and smart cropping. Use for computer vision and image unders... |
| `azure-ai-voicelive-dotnet` | Azure AI Voice Live SDK for .NET. Build real-time voice AI applications with bidirectional WebSocket communication. |
| `azure-ai-voicelive-java` | Azure AI VoiceLive SDK for Java. Real-time bidirectional voice conversations with AI assistants using WebSocket. |
| `azure-ai-voicelive-py` | Build real-time voice AI applications using Azure AI Voice Live SDK (azure-ai-voicelive). Use this skill when creating Python applications that nee... |
| `azure-ai-voicelive-ts` | Azure AI Voice Live SDK for JavaScript/TypeScript. Build real-time voice AI applications with bidirectional WebSocket communication. |
| `azure-communication-callautomation-java` | Build call automation workflows with Azure Communication Services Call Automation Java SDK. Use when implementing IVR systems, call routing, call r... |
| `azure-cosmos-java` | Azure Cosmos DB SDK for Java. NoSQL database operations with global distribution, multi-model support, and reactive patterns. |
| `azure-data-tables-java` | Build table storage applications with Azure Tables SDK for Java. Use when working with Azure Table Storage or Cosmos DB Table API for NoSQL key-val... |
| `azure-data-tables-py` | Azure Tables SDK for Python (Storage and Cosmos DB). Use for NoSQL key-value storage, entity CRUD, and batch operations. |
| `azure-functions` | Expert patterns for Azure Functions development including isolated worker model, Durable Functions orchestration, cold start optimization, and prod... |
| `azure-mgmt-arizeaiobservabilityeval-dotnet` | Azure Resource Manager SDK for Arize AI Observability and Evaluation (.NET). |
| `azure-mgmt-weightsandbiases-dotnet` | Azure Weights & Biases SDK for .NET. ML experiment tracking and model management via Azure Marketplace. Use for creating W&B instances, managing SS... |
| `azure-search-documents-dotnet` | Azure AI Search SDK for .NET (Azure.Search.Documents). Use for building search applications with full-text, vector, semantic, and hybrid search. |
| `azure-search-documents-py` | Azure AI Search SDK for Python. Use for vector search, hybrid search, semantic ranking, indexing, and skillsets. |
| `azure-search-documents-ts` | Build search applications using Azure AI Search SDK for JavaScript (@azure/search-documents). Use when creating/managing indexes, implementing vect... |
| `azure-storage-blob-java` | Build blob storage applications with Azure Storage Blob SDK for Java. Use when uploading, downloading, or managing files in Azure Blob Storage, wor... |
| `azure-storage-blob-py` | Azure Blob Storage SDK for Python. Use for uploading, downloading, listing blobs, managing containers, and blob lifecycle. |
| `azure-storage-blob-rust` | Azure Blob Storage SDK for Rust. Use for uploading, downloading, and managing blobs and containers. |
| `azure-storage-blob-ts` | Azure Blob Storage JavaScript/TypeScript SDK (@azure/storage-blob) for blob operations. Use for uploading, downloading, listing, and managing blobs... |
| `azure-storage-file-datalake-py` | Azure Data Lake Storage Gen2 SDK for Python. Use for hierarchical file systems, big data analytics, and file/directory operations. |
| `azure-storage-file-share-py` | Azure Storage File Share SDK for Python. Use for SMB file shares, directories, and file operations in the cloud. |
| `azure-storage-file-share-ts` | Azure File Share JavaScript/TypeScript SDK (@azure/storage-file-share) for SMB file share operations. |
| `azure-storage-queue-py` | Azure Queue Storage SDK for Python. Use for reliable message queuing, task distribution, and asynchronous processing. |
| `azure-storage-queue-ts` | Azure Queue Storage JavaScript/TypeScript SDK (@azure/storage-queue) for message queue operations. Use for sending, receiving, peeking, and deletin... |
| `backend-development-feature-development` | Orchestrate end-to-end backend feature development from requirements to deployment. Use when coordinating multi-phase feature delivery across teams... |
| `bdi-mental-states` | This skill should be used when the user asks to "model agent mental states", "implement BDI architecture", "create belief-desire-intention models",... |
| `beautiful-prose` | A hard-edged writing style contract for timeless, forceful English prose without modern AI tics. Use when users ask for prose or rewrites that must... |
| `behavioral-modes` | AI operational modes (brainstorm, implement, debug, review, teach, ship, orchestrate). Use to adapt behavior based on task type. |
| `bill-gates` | Agente que simula Bill Gates — cofundador da Microsoft, arquiteto da industria de software comercial, estrategista tecnologico global, investidor s... |
| `blockrun` | Use when user needs capabilities Claude lacks (image generation, real-time X/Twitter data) or explicitly requests external models ("blockrun", "use... |
| `blueprint` | Turn a one-line objective into a step-by-step construction plan any coding agent can execute cold. Each step has a self-contained context brief — a... |
| `box-automation` | Automate Box cloud storage operations including file upload/download, search, folder management, sharing, collaborations, and metadata queries via ... |
| `brand-guidelines-anthropic` | Applies Anthropic's official brand colors and typography to any sort of artifact that may benefit from having Anthropic's look-and-feel. Use it whe... |
| `brand-guidelines-community` | Applies Anthropic's official brand colors and typography to any sort of artifact that may benefit from having Anthropic's look-and-feel. Use it whe... |
| `browser-automation` | Browser automation powers web testing, scraping, and AI agent interactions. The difference between a flaky script and a reliable system comes down ... |
| `bullmq-specialist` | BullMQ expert for Redis-backed job queues, background processing, and reliable async execution in Node.js/TypeScript applications. Use when: bullmq... |
| `business-analyst` | Master modern business analysis with AI-powered analytics, real-time dashboards, and data-driven insights. Build comprehensive KPI frameworks, pred... |
| `c-pro` | Write efficient C code with proper memory management, pointer |
| `cc-skill-continuous-learning` | Development skill from everything-claude-code |
| `cc-skill-strategic-compact` | Development skill from everything-claude-code |
| `cirq` | Google quantum computing framework. Use when targeting Google Quantum AI hardware, designing noise-aware circuits, or running quantum characterizat... |
| `clarity-gate` | Pre-ingestion verification for epistemic quality in RAG systems. Ensures documents are properly qualified before entering knowledge bases. Produces... |
| `claude-ally-health` | A health assistant skill for medical information analysis, symptom tracking, and wellness guidance. |
| `claude-code-expert` | Especialista profundo em Claude Code - CLI da Anthropic. Maximiza produtividade com atalhos, hooks, MCPs, configuracoes avancadas, workflows, CLAUD... |
| `claude-code-guide` | Master guide for using Claude Code effectively. Includes configuration templates, prompting strategies "Thinking" keywords, debugging techniques, a... |
| `claude-d3js-skill` | Creating interactive data visualisations using d3.js. This skill should be used when creating custom charts, graphs, network diagrams, geographic v... |
| `claude-in-chrome-troubleshooting` | Diagnose and fix Claude in Chrome MCP extension connectivity issues. Use when mcp__claude-in-chrome__* tools fail, return "Browser extension is not... |
| `claude-monitor` | Monitor de performance do Claude Code e sistema local. Diagnostica lentidao, mede CPU/RAM/disco, verifica API latency e gera relatorios de saude do... |
| `claude-scientific-skills` | Scientific research and analysis skills |
| `claude-speed-reader` | -Speed read Claude's responses at 600+ WPM using RSVP with Spritz-style ORP highlighting |
| `claude-win11-speckit-update-skill` | Windows 11 system management |
| `cloudflare-workers-expert` | Expert in Cloudflare Workers and the Edge Computing ecosystem. Covers Wrangler, KV, D1, Durable Objects, and R2 storage. |
| `code-documentation-doc-generate` | You are a documentation expert specializing in creating comprehensive, maintainable documentation from code. Generate API docs, architecture diagra... |
| `code-review-ai-ai-review` | You are an expert AI-powered code review specialist combining automated static analysis, intelligent pattern recognition, and modern DevOps practic... |
| `code-reviewer` | Elite code review expert specializing in modern AI-powered code |
| `codex-review` | Professional code review with auto CHANGELOG generation, integrated with Codex AI |
| `competitive-landscape` | This skill should be used when the user asks to \\"analyze competitors", "assess competitive landscape", "identify differentiation", "evaluate mark... |
| `comprehensive-review-pr-enhance` | Generate structured PR descriptions from diffs, add review checklists, risk assessments, and test coverage summaries. Use when the user says "write... |
| `computer-use-agents` | Build AI agents that interact with computers like humans do - viewing screens, moving cursors, clicking buttons, and typing text. Covers Anthropic'... |
| `computer-vision-expert` | SOTA Computer Vision Expert (2026). Specialized in YOLO26, Segment Anything 3 (SAM 3), Vision Language Models, and real-time spatial analysis. |
| `conductor-setup` | Configure a Rails project to work with Conductor (parallel coding agents) |
| `content-marketer` | Elite content marketing strategist specializing in AI-powered content creation, omnichannel distribution, SEO optimization, and data-driven perform... |
| `context-agent` | Agente de contexto para continuidade entre sessoes. Salva resumos, decisoes, tarefas pendentes e carrega briefing automatico na sessao seguinte. |
| `context-compression` | This skill should be used when the user asks to "compress context", "summarize conversation history", "implement compaction", "reduce token usage",... |
| `context-degradation` | This skill should be used when the user asks to "diagnose context problems", "fix lost-in-middle issues", "debug agent failures", "understand conte... |
| `context-fundamentals` | This skill should be used when the user asks to "understand context", "explain context windows", "design agent architecture", "debug context issues... |
| `context-manager` | Elite AI context engineering specialist mastering dynamic context management, vector databases, knowledge graphs, and intelligent memory systems. |
| `context-optimization` | This skill should be used when the user asks to "optimize context", "reduce token costs", "improve context efficiency", "implement KV-cache optimiz... |
| `context-window-management` | Strategies for managing LLM context windows including summarization, trimming, routing, and avoiding context rot Use when: context window, token li... |
| `context7-auto-research` | Automatically fetch latest library/framework documentation for Claude Code via Context7 API |
| `conversation-memory` | Persistent memory systems for LLM conversations including short-term, long-term, and entity-based memory Use when: conversation memory, remember, m... |
| `core-components` | Core component library and design system patterns. Use when building UI, using design tokens, or working with the component library. |
| `cqrs-implementation` | Implement Command Query Responsibility Segregation for scalable architectures. Use when separating read and write models, optimizing query performa... |
| `crewai` | Expert in CrewAI - the leading role-based multi-agent framework used by 60% of Fortune 500 companies. Covers agent design with roles and goals, tas... |
| `customer-support` | Elite AI-powered customer support specialist mastering conversational AI, automated ticketing, sentiment analysis, and omnichannel support experien... |
| `data-engineer` | Build scalable data pipelines, modern data warehouses, and real-time streaming architectures. Implements Apache Spark, dbt, Airflow, and cloud-nati... |
| `data-engineering-data-driven-feature` | Build features guided by data insights, A/B testing, and continuous measurement using specialized agents for analysis, implementation, and experime... |
| `data-scientist` | Expert data scientist for advanced analytics, machine learning, and statistical modeling. Handles complex data analysis, predictive modeling, and b... |
| `data-structure-protocol` | Give agents persistent structural memory of a codebase — navigate dependencies, track public APIs, and understand why connections exist without re-... |
| `database-architect` | Expert database architect specializing in data layer design from scratch, technology selection, schema modeling, and scalable database architectures. |
| `dbt-transformation-patterns` | Master dbt (data build tool) for analytics engineering with model organization, testing, documentation, and incremental strategies. Use when buildi... |
| `deep-research` | Execute autonomous multi-step research using Google Gemini Deep Research Agent. Use for: market analysis, competitive landscaping, literature revie... |
| `design-orchestration` | Orchestrates design workflows by routing work through brainstorming, multi-agent review, and execution readiness in the correct order. |
| `devcontainer-setup` | Creates devcontainers with Claude Code, language-specific tooling (Python/Node/Rust/Go), and persistent volumes. Use when adding devcontainer suppo... |
| `dispatching-parallel-agents` | Use when facing 2+ independent tasks that can be worked on without shared state or sequential dependencies |
| `documentation-generation-doc-generate` | You are a documentation expert specializing in creating comprehensive, maintainable documentation from code. Generate API docs, architecture diagra... |
| `documentation-templates` | Documentation templates and structure guidelines. README, API docs, code comments, and AI-friendly documentation. |
| `docx-official` | Comprehensive document creation, editing, and analysis with support for tracked changes, comments, formatting preservation, and text extraction. Wh... |
| `domain-driven-design` | Plan and route Domain-Driven Design work from strategic modeling to tactical implementation and evented architecture patterns. |
| `draw` | Vector graphics and diagram creation, format conversion (ODG/SVG/PDF) with LibreOffice Draw. |
| `earllm-build` | Build, maintain, and extend the EarLLM One Android project — a Kotlin/Compose app that connects Bluetooth earbuds to an LLM via voice pipeline. Use... |
| `elon-musk` | Agente que simula Elon Musk com profundidade psicologica e comunicacional de alta fidelidade. Ativado para: "fale como Elon", "simule Elon Musk", "... |
| `embedding-strategies` | Select and optimize embedding models for semantic search and RAG applications. Use when choosing embedding models, implementing chunking strategies... |
| `emblemai-crypto-wallet` | Crypto wallet management across 7 blockchains via EmblemAI Agent Hustle API. Balance checks, token swaps, portfolio analysis, and transaction execu... |
| `energy-procurement` | Codified expertise for electricity and gas procurement, tariff optimisation, demand charge management, renewable PPA evaluation, and multi-facility... |
| `enhance-prompt` | Transforms vague UI ideas into polished, Stitch-optimized prompts. Enhances specificity, adds UI/UX keywords, injects design system context, and st... |
| `error-debugging-multi-agent-review` | Use when working with error debugging multi agent review |
| `evaluation` | Build evaluation frameworks for agent systems |
| `event-sourcing-architect` | Expert in event sourcing, CQRS, and event-driven architecture patterns. Masters event store design, projection building, saga orchestration, and ev... |
| `fal-audio` | Text-to-speech and speech-to-text using fal.ai audio models |
| `fal-generate` | Generate images and videos using fal.ai AI models |
| `fal-image-edit` | AI-powered image editing with style transfer and object removal |
| `fal-platform` | Platform APIs for model management, pricing, and usage tracking |
| `fal-upscale` | Upscale and enhance image and video resolution using AI |
| `fal-workflow` | Generate workflow JSON files for chaining AI models |
| `ffuf-claude-skill` | Web fuzzing with ffuf |
| `figma-automation` | Automate Figma tasks via Rube MCP (Composio): files, components, design tokens, comments, exports. Always search tools first for current schemas. |
| `file-uploads` | Expert at handling file uploads and cloud storage. Covers S3, Cloudflare R2, presigned URLs, multipart uploads, and image optimization. Knows how t... |
| `fp-pragmatic` | A practical, jargon-free guide to functional programming - the 80/20 approach that gets results without the academic overhead |
| `fp-ts-pragmatic` | A practical, jargon-free guide to fp-ts functional programming - the 80/20 approach that gets results without the academic overhead. Use when writi... |
| `framework-migration-legacy-modernize` | Orchestrate a comprehensive legacy system modernization using the strangler fig pattern, enabling gradual replacement of outdated components while ... |
| `free-tool-strategy` | When the user wants to plan, evaluate, or build a free tool for marketing purposes — lead generation, SEO value, or brand awareness. Also use when ... |
| `full-stack-orchestration-full-stack-feature` | Use when working with full stack orchestration full stack feature |
| `game-development` | Game development orchestrator. Routes to platform-specific skills based on project needs. |
| `gemini-api-dev` | Use this skill when building applications with Gemini models, Gemini API, working with multimodal content (text, images, audio, video), implementin... |
| `gemini-api-integration` | Use when integrating Google Gemini API into projects. Covers model selection, multimodal inputs, streaming, function calling, and production best p... |
| `geo-fundamentals` | Generative Engine Optimization for AI search engines (ChatGPT, Claude, Perplexity). |
| `geoffrey-hinton` | Agente que simula Geoffrey Hinton — Godfather of Deep Learning, Prêmio Turing 2018, criador do backpropagation e das Deep Belief Networks. Use quan... |
| `git-pr-workflows-git-workflow` | Orchestrate a comprehensive git workflow from code review through PR creation, leveraging specialized agents for quality assurance, testing, and de... |
| `github-workflow-automation` | Automate GitHub workflows with AI assistance. Includes PR reviews, issue triage, CI/CD integration, and Git operations. Use when automating GitHub ... |
| `hierarchical-agent-memory` | Scoped CLAUDE.md memory system that reduces context token spend. Creates directory-level context files, tracks savings via dashboard, and routes ag... |
| `hig-components-controls` | Apple HIG guidance for selection and input controls including pickers, toggles, sliders, steppers, segmented controls, combo boxes, text fields, te... |
| `hig-technologies` | Apple HIG guidance for Apple technology integrations: Siri, Apple Pay, HealthKit, HomeKit, ARKit, machine learning, generative AI, iCloud, Sign in ... |
| `hosted-agents` | Build background agents in sandboxed environments. Use for hosted coding agents, sandboxed VMs, Modal sandboxes, and remote coding environments. |
| `hosted-agents-v2-py` | Build hosted agents using Azure AI Projects SDK with ImageBasedHostedAgentDefinition. Use when creating container-based agents in Azure AI Foundry. |
| `hugging-face-cli` | Execute Hugging Face Hub operations using the `hf` CLI. Use when the user needs to download models/datasets/spaces, upload files to Hub repositorie... |
| `hugging-face-datasets` | Create and manage datasets on Hugging Face Hub. Supports initializing repos, defining configs/system prompts, streaming row updates, and SQL-based ... |
| `hugging-face-evaluation` | Add and manage evaluation results in Hugging Face model cards. Supports extracting eval tables from README content, importing scores from Artificia... |
| `hugging-face-model-trainer` | This skill should be used when users want to train or fine-tune language models using TRL (Transformer Reinforcement Learning) on Hugging Face Jobs... |
| `hybrid-search-implementation` | Combine vector and keyword search for improved retrieval. Use when implementing RAG systems, building search engines, or when neither approach alon... |
| `iconsax-library` | Extensive icon library and AI-driven icon generation skill for premium UI/UX design. |
| `ilya-sutskever` | Agente que simula Ilya Sutskever — co-fundador da OpenAI, ex-Chief Scientist, fundador da SSI. Use quando quiser perspectivas sobre: AGI safety-fir... |
| `image-studio` | Studio de geracao de imagens inteligente — roteamento automatico entre ai-studio-image (fotos humanizadas/influencer) e stability-ai (arte/ ilustra... |
| `imagen` | AI image generation skill powered by Google Gemini, enabling seamless visual content creation for UI placeholders, documentation, and design assets. |
| `incident-response-smart-fix` | [Extended thinking: This workflow implements a sophisticated debugging and resolution pipeline that leverages AI-assisted debugging tools and obser... |
| `infinite-gratitude` | Multi-agent research skill for parallel research execution (10 agents, battle-tested with real case studies). |
| `internal-comms-anthropic` | A set of resources to help me write all kinds of internal communications, using the formats that my company likes to use. Claude should use this sk... |
| `internal-comms-community` | A set of resources to help me write all kinds of internal communications, using the formats that my company likes to use. Claude should use this sk... |
| `kubernetes-architect` | Expert Kubernetes architect specializing in cloud-native infrastructure, advanced GitOps workflows (ArgoCD/Flux), and enterprise container orchestr... |
| `kubernetes-deployment` | Kubernetes deployment workflow for container orchestration, Helm charts, service mesh, and production-ready K8s configurations. |
| `langchain-architecture` | Design LLM applications using the LangChain framework with agents, memory, and tool integration patterns. Use when building LangChain applications,... |
| `langfuse` | Expert in Langfuse - the open-source LLM observability platform. Covers tracing, prompt management, evaluation, datasets, and integration with Lang... |
| `langgraph` | Expert in LangGraph - the production-grade framework for building stateful, multi-actor AI applications. Covers graph construction, state managemen... |
| `last30days` | Research a topic from the last 30 days on Reddit + X + Web, become an expert, and write copy-paste-ready prompts for the user's target tool. |
| `lightning-architecture-review` | Review Bitcoin Lightning Network protocol designs, compare channel factory approaches, and analyze Layer 2 scaling tradeoffs. Covers trust models, ... |
| `linear-claude-skill` | Manage Linear issues, projects, and teams |
| `llm-app-patterns` | Production-ready patterns for building LLM applications. Covers RAG pipelines, agent architectures, prompt IDEs, and LLMOps monitoring. Use when de... |
| `llm-application-dev-ai-assistant` | You are an AI assistant development expert specializing in creating intelligent conversational interfaces, chatbots, and AI-powered applications. D... |
| `llm-application-dev-langchain-agent` | You are an expert LangChain agent developer specializing in production-grade AI systems using LangChain 0.1+ and LangGraph. |
| `llm-application-dev-prompt-optimize` | You are an expert prompt engineer specializing in crafting effective prompts for LLMs through advanced techniques including constitutional AI, chai... |
| `llm-evaluation` | Implement comprehensive evaluation strategies for LLM applications using automated metrics, human feedback, and benchmarking. Use when testing LLM ... |
| `llm-ops` | LLM Operations -- RAG, embeddings, vector databases, fine-tuning, prompt engineering avancado, custos de LLM, evals de qualidade e arquiteturas de ... |
| `llm-prompt-optimizer` | Use when improving prompts for any LLM. Applies proven prompt engineering techniques to boost output quality, reduce hallucinations, and cut token ... |
| `llm-structured-output` | Get reliable JSON, enums, and typed objects from LLMs using response_format, tool_use, and schema-constrained decoding across OpenAI, Anthropic, an... |
| `local-llm-expert` | Master local LLM inference, model selection, VRAM optimization, and local deployment using Ollama, llama.cpp, vLLM, and LM Studio. Expert in quanti... |
| `m365-agents-ts` | Microsoft 365 Agents SDK for TypeScript/Node.js. |
| `magic-animator` | AI-powered animation tool for creating motion in logos, UI, icons, and social media assets. |
| `makepad-splash` | CRITICAL: Use for Makepad Splash scripting language. Triggers on: splash language, makepad script, makepad scripting, script!, cx.eval, makepad dyn... |
| `manifest` | Install and configure the Manifest observability plugin for your agents. Use when setting up telemetry, configuring API keys, or troubleshooting th... |
| `marketing-psychology` | Apply behavioral science and mental models to marketing decisions, prioritized using a psychological leverage and feasibility scoring system. |
| `mcp-builder` | Guide for creating high-quality MCP (Model Context Protocol) servers that enable LLMs to interact with external services through well-designed tool... |
| `mcp-builder-ms` | Guide for creating high-quality MCP (Model Context Protocol) servers that enable LLMs to interact with external services through well-designed tool... |
| `memory-safety-patterns` | Implement memory-safe programming with RAII, ownership, smart pointers, and resource management across Rust, C++, and C. Use when writing safe syst... |
| `memory-systems` | Design short-term, long-term, and graph-based memory architectures |
| `ml-engineer` | Build production ML systems with PyTorch 2.x, TensorFlow, and modern ML frameworks. Implements model serving, feature engineering, A/B testing, and... |
| `ml-pipeline-workflow` | Build end-to-end MLOps pipelines from data preparation through model training, validation, and production deployment. Use when creating ML pipeline... |
| `mlops-engineer` | Build comprehensive ML pipelines, experiment tracking, and model registries with MLflow, Kubeflow, and modern MLOps tools. |
| `molykit` | CRITICAL: Use for MolyKit AI chat toolkit. Triggers on: BotClient, OpenAI, SSE streaming, AI chat, molykit, PlatformSend, spawn(), ThreadToken, cro... |
| `monetization` | Estrategia e implementacao de monetizacao para produtos digitais - Stripe, subscriptions, pricing experiments, freemium, upgrade flows, churn preve... |
| `multi-advisor` | Conselho de especialistas — consulta multiplos agentes do ecossistema em paralelo para analise multi-perspectiva de qualquer topico. Ativa personas... |
| `multi-agent-brainstorming` | Simulate a structured peer-review process using multiple specialized agents to validate designs, surface hidden assumptions, and identify failure m... |
| `multi-agent-patterns` | This skill should be used when the user asks to "design multi-agent system", "implement supervisor pattern", "create swarm architecture", "coordina... |
| `multi-cloud-architecture` | Design multi-cloud architectures using a decision framework to select and integrate services across AWS, Azure, and GCP. Use when building multi-cl... |
| `nanobanana-ppt-skills` | AI-powered PPT generation with document analysis and styled images |
| `nerdzao-elite-gemini-high` | Modo Elite Coder + UX Pixel-Perfect otimizado especificamente para Gemini 3.1 Pro High. Workflow completo com foco em qualidade máxima e eficiência... |
| `nosql-expert` | Expert guidance for distributed NoSQL databases (Cassandra, DynamoDB). Focuses on mental models, query-first modeling, single-table design, and avo... |
| `odoo-backup-strategy` | Complete Odoo backup and restore strategy: database dumps, filestore backup, automated scheduling, cloud storage upload, and tested restore procedu... |
| `odoo-ecommerce-configurator` | Expert guide for Odoo eCommerce and Website: product catalog, payment providers, shipping methods, SEO, and order-to-fulfillment workflow. |
| `odoo-module-developer` | Expert guide for creating custom Odoo modules. Covers __manifest__.py, model inheritance, ORM patterns, and module structure best practices. |
| `odoo-performance-tuner` | Expert guide for diagnosing and fixing Odoo performance issues: slow queries, worker configuration, memory limits, PostgreSQL tuning, and profiling... |
| `parallel-agents` | Multi-agent orchestration patterns. Use when multiple independent tasks can run with different domain expertise or when comprehensive analysis requ... |
| `pdf-official` | Comprehensive PDF manipulation toolkit for extracting text and tables, creating new PDFs, merging/splitting documents, and handling forms. When Cla... |
| `performance-testing-review-ai-review` | You are an expert AI-powered code review specialist combining automated static analysis, intelligent pattern recognition, and modern DevOps practic... |
| `performance-testing-review-multi-agent-review` | Use when working with performance testing review multi agent review |
| `pipecat-friday-agent` | Build a low-latency, Iron Man-inspired tactical voice assistant (F.R.I.D.A.Y.) using Pipecat, Gemini, and OpenAI. |
| `playwright-java` | Scaffold, write, debug, and enhance enterprise-grade Playwright E2E tests in Java using Page Object Model, JUnit 5, Allure reporting, and parallel ... |
| `podcast-generation` | Generate AI-powered podcast-style audio narratives using Azure OpenAI's GPT Realtime Mini model via WebSocket. Use when building text-to-speech fea... |
| `polars` | Fast in-memory DataFrame library for datasets that fit in RAM. Use when pandas is too slow but data still fits in memory. Lazy evaluation, parallel... |
| `pptx-official` | Presentation creation, editing, and analysis. When Claude needs to work with presentations (.pptx files) for: (1) Creating new presentations, (2) M... |
| `prisma-expert` | Prisma ORM expert for schema design, migrations, query optimization, relations modeling, and database operations. Use PROACTIVELY for Prisma schema... |
| `product-design` | Design de produto nivel Apple — sistemas visuais, UX flows, acessibilidade, linguagem visual proprietaria, design tokens, prototipagem e handoff. C... |
| `product-manager` | Senior PM agent with 6 knowledge domains, 30+ frameworks, 12 templates, and 32 SaaS metrics with formulas. Pure Markdown, zero scripts. |
| `programmatic-seo` | Design and evaluate programmatic SEO strategies for creating SEO-driven pages at scale using templates and structured data. |
| `progressive-estimation` | Estimate AI-assisted and hybrid human+agent development work with research-backed PERT statistics and calibration feedback loops |
| `project-development` | This skill should be used when the user asks to "start an LLM project", "design batch pipeline", "evaluate task-model fit", "structure agent projec... |
| `projection-patterns` | Build read models and projections from event streams. Use when implementing CQRS read sides, building materialized views, or optimizing query perfo... |
| `prometheus-configuration` | Set up Prometheus for comprehensive metric collection, storage, and monitoring of infrastructure and applications. Use when implementing metrics co... |
| `prompt-caching` | Caching strategies for LLM prompts including Anthropic prompt caching, response caching, and CAG (Cache Augmented Generation) Use when: prompt cach... |
| `prompt-engineer` | Transforms user prompts into optimized prompts using frameworks (RTF, RISEN, Chain of Thought, RODES, Chain of Density, RACE, RISE, STAR, SOAP, CLE... |
| `prompt-engineering` | Expert guide on prompt engineering patterns, best practices, and optimization techniques. Use when user wants to improve prompts, learn prompting s... |
| `prompt-engineering-patterns` | Master advanced prompt engineering techniques to maximize LLM performance, reliability, and controllability in production. Use when optimizing prom... |
| `prompt-library` | Curated collection of high-quality prompts for various use cases. Includes role-based prompts, task-specific templates, and prompt refinement techn... |
| `pydantic-models-py` | Create Pydantic models following the multi-model pattern with Base, Create, Update, Response, and InDB variants. Use when defining API request/resp... |
| `python-performance-optimization` | Profile and optimize Python code using cProfile, memory profilers, and performance best practices. Use when debugging slow Python code, optimizing ... |
| `quant-analyst` | Build financial models, backtest trading strategies, and analyze market data. Implements risk metrics, portfolio optimization, and statistical arbi... |
| `rag-engineer` | Expert in building Retrieval-Augmented Generation systems. Masters embedding models, vector databases, chunking strategies, and retrieval optimizat... |
| `rag-implementation` | RAG (Retrieval-Augmented Generation) implementation workflow covering embedding selection, vector database setup, chunking strategies, and retrieva... |
| `react-flow-architect` | Expert ReactFlow architect for building interactive graph applications with hierarchical node-edge systems, performance optimization, and auto-layo... |
| `recallmax` | FREE — God-tier long-context memory for AI agents. Injects 500K-1M clean tokens, auto-summarizes with tone/intent preservation, compresses 14-turn ... |
| `robius-widget-patterns` | CRITICAL: Use for Robius widget patterns. Triggers on: apply_over, TextOrImage, modal, 可复用, 模态, collapsible, drag drop, reusable widget, widget des... |
| `saga-orchestration` | Implement saga patterns for distributed transactions and cross-aggregate workflows. Use when coordinating multi-step business processes, handling c... |
| `sam-altman` | Agente que simula Sam Altman — CEO da OpenAI, ex-presidente da Y Combinator, arquiteto da era AGI. Use quando quiser: perspectivas sobre startups e... |
| `scientific-writing` | Core skill for the deep research and writing tool. Write scientific manuscripts in full paragraphs (never bullet points). Use two-stage process wit... |
| `scikit-learn` | Machine learning in Python with scikit-learn. Use for classification, regression, clustering, model evaluation, and ML pipelines. |
| `seek-and-analyze-video` | Seek and analyze video content using Memories.ai Large Visual Memory Model for persistent video intelligence |
| `seo-fundamentals` | Core principles of SEO including E-E-A-T, Core Web Vitals, technical foundations, content quality, and how modern search engines evaluate pages. |
| `shader-programming-glsl` | Expert guide for writing efficient GLSL shaders (Vertex/Fragment) for web and game engines, covering syntax, uniforms, and common effects. |
| `similarity-search-patterns` | Implement efficient similarity search with vector databases. Use when building semantic search, implementing nearest neighbor queries, or optimizin... |
| `skill-check` | Validate Claude Code skills against the agentskills specification. Catches structural, semantic, and naming issues before users do. |
| `skill-creator-ms` | Guide for creating effective skills for AI coding agents working with Azure SDKs and Microsoft Foundry services. Use when creating new skills or up... |
| `skill-developer` | Create and manage Claude Code skills following Anthropic best practices. Use when creating new skills, modifying skill-rules.json, understanding tr... |
| `skill-improver` | Iteratively reviews and fixes Claude Code skill quality issues until they meet standards. Runs automated fix-review cycles using the skill-reviewer... |
| `skill-installer` | Instala, valida, registra e verifica novas skills no ecossistema. 10 checks de seguranca, copia, registro no orchestrator e verificacao pos-instala... |
| `skill-seekers` | -Automatically convert documentation websites, GitHub repositories, and PDFs into Claude AI skills in minutes. |
| `skill-writer` | Create and improve agent skills following the Agent Skills specification. Use when asked to create, write, or update skills. |
| `social-orchestrator` | Orquestrador unificado de canais sociais — coordena Instagram, Telegram e WhatsApp em um unico fluxo de trabalho. Publicacao cross-channel, metrica... |
| `spark-optimization` | Optimize Apache Spark jobs with partitioning, caching, shuffle optimization, and memory tuning. Use when improving Spark performance, debugging slo... |
| `spline-3d-integration` | Use when adding interactive 3D scenes from Spline.design to web projects, including React embedding and runtime control API. |
| `sql-pro` | Master modern SQL with cloud-native databases, OLTP/OLAP optimization, and advanced query techniques. Expert in performance tuning, data modeling, ... |
| `stability-ai` | Geracao de imagens via Stability AI (SD3.5, Ultra, Core). Text-to-image, img2img, inpainting, upscale, remove-bg, search-replace. 15 estilos artist... |
| `startup-analyst` | Expert startup business analyst specializing in market sizing, financial modeling, competitive analysis, and strategic planning for early-stage com... |
| `startup-business-analyst-financial-projections` | Create detailed 3-5 year financial model with revenue, costs, cash flow, and scenarios |
| `startup-financial-modeling` | This skill should be used when the user asks to \\"create financial projections", "build a financial model", "forecast revenue", "calculate burn ra... |
| `statsmodels` | Statistical models library for Python. Use when you need specific model classes (OLS, GLM, mixed models, ARIMA) with detailed diagnostics, residual... |
| `steve-jobs` | Agente que simula Steve Jobs — cofundador da Apple, CEO da Pixar, fundador da NeXT, o maior designer de produtos tecnologicos da historia e o mais ... |
| `stitch-loop` | Teaches agents to iteratively build websites using Stitch with an autonomous baton-passing loop pattern |
| `stitch-ui-design` | Expert guide for creating effective prompts for Google Stitch AI UI design tool. Use when user wants to design UI/UX in Stitch, create app interfac... |
| `subagent-driven-development` | Use when executing implementation plans with independent tasks in the current session |
| `supabase-automation` | Automate Supabase database queries, table management, project administration, storage, edge functions, and SQL execution via Rube MCP (Composio). A... |
| `superpowers-lab` | Lab environment for Claude superpowers |
| `tailwind-design-system` | Build scalable design systems with Tailwind CSS, design tokens, component libraries, and responsive patterns. Use when creating component libraries... |
| `tailwind-patterns` | Tailwind CSS v4 principles. CSS-first configuration, container queries, modern patterns, design token architecture. |
| `task-intelligence` | Protocolo de Inteligência Pré-Tarefa — ativa TODOS os agentes relevantes do ecossistema ANTES de executar qualquer tarefa solicitada pelo usuário. ... |
| `tdd-orchestrator` | Master TDD orchestrator specializing in red-green-refactor discipline, multi-agent workflow coordination, and comprehensive test-driven development... |
| `team-collaboration-standup-notes` | You are an expert team communication specialist focused on async-first standup practices, AI-assisted note generation from commit history, and effe... |
| `telegram-bot-builder` | Expert in building Telegram bots that solve real problems - from simple automation to complex AI-powered bots. Covers bot architecture, the Telegra... |
| `temporal-python-pro` | Master Temporal workflow orchestration with Python SDK. Implements durable workflows, saga patterns, and distributed transactions. Covers async/awa... |
| `test-automator` | Master AI-powered test automation with modern frameworks, self-healing tests, and comprehensive quality engineering. Build scalable testing strateg... |
| `threejs-loaders` | Three.js asset loading - GLTF, textures, images, models, async patterns. Use when loading 3D models, textures, HDR environments, or managing loadin... |
| `threejs-shaders` | Three.js shaders - GLSL, ShaderMaterial, uniforms, custom effects. Use when creating custom visual effects, modifying vertices, writing fragment sh... |
| `tool-design` | Build tools that agents can use effectively, including architectural reduction patterns |
| `trigger-dev` | Trigger.dev expert for background jobs, AI workflows, and reliable async execution with excellent developer experience and TypeScript-first design.... |
| `ui-skills` | Opinionated, evolving constraints to guide agents when building interfaces |
| `uniprot-database` | Direct REST API access to UniProt. Protein searches, FASTA retrieval, ID mapping, Swiss-Prot/TrEMBL. For Python workflows with multiple databases, ... |
| `unit-testing-test-generate` | Generate comprehensive, maintainable unit tests across languages with strong coverage and edge case focus. |
| `vector-database-engineer` | Expert in vector databases, embedding strategies, and semantic search implementation. Masters Pinecone, Weaviate, Qdrant, Milvus, and pgvector for ... |
| `vector-index-tuning` | Optimize vector index performance for latency, recall, and memory. Use when tuning HNSW parameters, selecting quantization strategies, or scaling v... |
| `vercel-ai-sdk-expert` | Expert in the Vercel AI SDK. Covers Core API (generateText, streamText), UI hooks (useChat, useCompletion), tool calling, and streaming UI componen... |
| `vexor` | Vector-powered CLI for semantic file search with a Claude/Codex skill |
| `videodb-skills` | Upload, stream, search, edit, transcribe, and generate AI video and audio using the VideoDB SDK. |
| `vizcom` | AI-powered product design tool for transforming sketches into full-fidelity 3D renders. |
| `voice-agents` | Voice agents represent the frontier of AI interaction - humans speaking naturally with AI systems. The challenge isn't just speech recognition and ... |
| `voice-ai-development` | Expert in building voice AI applications - from real-time voice agents to voice-enabled apps. Covers OpenAI Realtime API, Vapi for voice agents, De... |
| `voice-ai-engine-development` | Build real-time conversational AI voice engines using async worker pipelines, streaming transcription, LLM agents, and TTS synthesis with interrupt... |
| `warren-buffett` | Agente que simula Warren Buffett — o maior investidor do seculo XX e XXI, CEO da Berkshire Hathaway, discipulo de Benjamin Graham e socio intelectu... |
| `web-artifacts-builder` | Suite of tools for creating elaborate, multi-component claude.ai HTML artifacts using modern frontend web technologies (React, Tailwind CSS, shadcn... |
| `web-scraper` | Web scraping inteligente multi-estrategia. Extrai dados estruturados de paginas web (tabelas, listas, precos). Paginacao, monitoramento e export CS... |
| `workflow-automation` | Workflow automation is the infrastructure that makes AI agents reliable. Without durable execution, a network hiccup during a 10-step payment flow ... |
| `workflow-orchestration-patterns` | Design durable workflows with Temporal for distributed systems. Covers workflow vs activity separation, saga patterns, state management, and determ... |
| `writing-skills` | Use when creating, updating, or improving agent skills. |
| `xlsx-official` | Comprehensive spreadsheet creation, editing, and analysis with support for formulas, formatting, data analysis, and visualization. When Claude need... |
| `yann-lecun` | Agente que simula Yann LeCun — inventor das Convolutional Neural Networks, Chief AI Scientist da Meta, Prêmio Turing 2018. Use quando quiser: persp... |
| `yann-lecun-debate` | Sub-skill de debates e posições de Yann LeCun. Cobre críticas técnicas detalhadas aos LLMs, rivalidades intelectuais (LeCun vs Hinton, Sutskever, R... |
| `yann-lecun-filosofia` | Sub-skill filosófica e pedagógica de Yann LeCun. Cobre filosofia do open source (LLaMA, soberania tecnológica, analogia Linux), análise de incentiv... |
| `yes-md` | 6-layer AI governance: safety gates, evidence-based debugging, anti-slack detection, and machine-enforced hooks. Makes AI safe, thorough, and honest. |

[↑ Voltar ao Índice](#indice-de-categorias)

---

## 📁 Automation & SaaS

| ID | Descrição |
| :--- | :--- |
| `amplitude-automation` | Automate Amplitude tasks via Rube MCP (Composio): events, user activity, cohorts, user identification. Always search tools first for current schemas. |
| `angular-ui-patterns` | Modern Angular UI patterns for loading states, error handling, and data display. Use when building UI components, handling async data, or managing ... |
| `api-documentation` | API documentation workflow for generating OpenAPI specs, creating developer guides, and maintaining comprehensive API documentation. |
| `apify-actor-development` | Develop, debug, and deploy Apify Actors - serverless cloud programs for web scraping, automation, and data processing. Use when creating new Actors... |
| `apify-actorization` | Convert existing projects into Apify Actors - serverless cloud programs. Actorize JavaScript/TypeScript (SDK with Actor.init/exit), Python (async c... |
| `asana-automation` | Automate Asana tasks via Rube MCP (Composio): tasks, projects, sections, teams, workspaces. Always search tools first for current schemas. |
| `async-python-patterns` | Master Python asyncio, concurrent programming, and async/await patterns for high-performance applications. Use when building async APIs, concurrent... |
| `aws-skills` | AWS development with infrastructure automation and cloud architecture patterns |
| `azure-communication-callingserver-java` | Azure Communication Services CallingServer (legacy) Java SDK. Note - This SDK is deprecated. Use azure-communication-callautomation instead for new... |
| `azure-mgmt-botservice-dotnet` | Azure Resource Manager SDK for Bot Service in .NET. Management plane operations for creating and managing Azure Bot resources, channels (Teams, Dir... |
| `bamboohr-automation` | Automate BambooHR tasks via Rube MCP (Composio): employees, time-off, benefits, dependents, employee updates. Always search tools first for current... |
| `basecamp-automation` | Automate Basecamp project management, to-dos, messages, people, and to-do list organization via Rube MCP (Composio). Always search tools first for ... |
| `bash-pro` | Master of defensive Bash scripting for production automation, CI/CD pipelines, and system utilities. Expert in safe, portable, and testable shell s... |
| `bash-scripting` | Bash scripting workflow for creating production-ready shell scripts with defensive patterns, error handling, and testing. |
| `billing-automation` | Build automated billing systems for recurring payments, invoicing, subscription lifecycle, and dunning management. Use when implementing subscripti... |
| `biopython` | Comprehensive molecular biology toolkit. Use for sequence manipulation, file parsing (FASTA/GenBank/PDB), phylogenetics, and programmatic NCBI/PubM... |
| `bitbucket-automation` | Automate Bitbucket repositories, pull requests, branches, issues, and workspace management via Rube MCP (Composio). Always search tools first for c... |
| `brevo-automation` | Automate Brevo (Sendinblue) tasks via Rube MCP (Composio): manage email campaigns, create/edit templates, track senders, and monitor campaign perfo... |
| `cal-com-automation` | Automate Cal.com tasks via Rube MCP (Composio): manage bookings, check availability, configure webhooks, and handle teams. Always search tools firs... |
| `calc` | Spreadsheet creation, format conversion (ODS/XLSX/CSV), formulas, data automation with LibreOffice Calc. |
| `calendly-automation` | Automate Calendly scheduling, event management, invitee tracking, availability checks, and organization administration via Rube MCP (Composio). Alw... |
| `canva-automation` | Automate Canva tasks via Rube MCP (Composio): designs, exports, folders, brand templates, autofill. Always search tools first for current schemas. |
| `changelog-automation` | Automate changelog generation from commits, PRs, and releases following Keep a Changelog format. Use when setting up release workflows, generating ... |
| `cicd-automation-workflow-automate` | You are a workflow automation expert specializing in creating efficient CI/CD pipelines, GitHub Actions workflows, and automated development proces... |
| `circleci-automation` | Automate CircleCI tasks via Rube MCP (Composio): trigger pipelines, monitor workflows/jobs, retrieve artifacts and test metadata. Always search too... |
| `clickup-automation` | Automate ClickUp project management including tasks, spaces, folders, lists, comments, and team operations via Rube MCP (Composio). Always search t... |
| `close-automation` | Automate Close CRM tasks via Rube MCP (Composio): create leads, manage calls/SMS, handle tasks, and track notes. Always search tools first for curr... |
| `cloud-devops` | Cloud infrastructure and DevOps workflow covering AWS, Azure, GCP, Kubernetes, Terraform, CI/CD, monitoring, and cloud-native development. |
| `coda-automation` | Automate Coda tasks via Rube MCP (Composio): manage docs, pages, tables, rows, formulas, permissions, and publishing. Always search tools first for... |
| `conductor-implement` | Execute tasks from a track's implementation plan following TDD workflow |
| `confluence-automation` | Automate Confluence page creation, content search, space management, labels, and hierarchy navigation via Rube MCP (Composio). Always search tools ... |
| `convertkit-automation` | Automate ConvertKit (Kit) tasks via Rube MCP (Composio): manage subscribers, tags, broadcasts, and broadcast stats. Always search tools first for c... |
| `create-pr` | Alias for sentry-skills:pr-writer. Use when users explicitly ask for "create-pr" or reference the legacy skill name. Redirects to the canonical PR ... |
| `csharp-pro` | Write modern C# code with advanced features like records, pattern matching, and async/await. Optimizes .NET applications, implements enterprise pat... |
| `database` | Database development and operations workflow covering SQL, NoSQL, database design, migrations, optimization, and data engineering. |
| `database-admin` | Expert database administrator specializing in modern cloud databases, automation, and reliability engineering. |
| `datadog-automation` | Automate Datadog tasks via Rube MCP (Composio): query metrics, search logs, manage monitors/dashboards, create events and downtimes. Always search ... |
| `dbos-golang` | DBOS Go SDK for building reliable, fault-tolerant applications with durable workflows. Use this skill when writing Go code with DBOS, creating work... |
| `dbos-python` | DBOS Python SDK for building reliable, fault-tolerant applications with durable workflows. Use this skill when writing Python code with DBOS, creat... |
| `dbos-typescript` | DBOS TypeScript SDK for building reliable, fault-tolerant applications with durable workflows. Use this skill when writing TypeScript code with DBO... |
| `deployment-engineer` | Expert deployment engineer specializing in modern CI/CD pipelines, GitOps workflows, and advanced deployment automation. |
| `deployment-procedures` | Production deployment principles and decision-making. Safe deployment workflows, rollback strategies, and verification. Teaches thinking, not scripts. |
| `development` | Comprehensive web, mobile, and backend development workflow bundling frontend, backend, full-stack, and mobile development skills for end-to-end ap... |
| `discord-automation` | Automate Discord tasks via Rube MCP (Composio): messages, channels, roles, webhooks, reactions. Always search tools first for current schemas. |
| `discord-bot-architect` | Specialized skill for building production-ready Discord bots. Covers Discord.js (JavaScript) and Pycord (Python), gateway intents, slash commands, ... |
| `distributed-debugging-debug-trace` | You are a debugging expert specializing in setting up comprehensive debugging environments, distributed tracing, and diagnostic tools. Configure de... |
| `django-pro` | Master Django 5.x with async views, DRF, Celery, and Django Channels. Build scalable web applications with proper architecture, testing, and deploy... |
| `documentation` | Documentation generation workflow covering API docs, architecture docs, README files, code comments, and technical writing. |
| `docusign-automation` | Automate DocuSign tasks via Rube MCP (Composio): templates, envelopes, signatures, document management. Always search tools first for current schemas. |
| `dotnet-backend-patterns` | Master C#/.NET backend development patterns for building robust APIs, MCP servers, and enterprise applications. Covers async/await, dependency inje... |
| `dropbox-automation` | Automate Dropbox file management, sharing, search, uploads, downloads, and folder operations via Rube MCP (Composio). Always search tools first for... |
| `dx-optimizer` | Developer Experience specialist. Improves tooling, setup, and workflows. Use PROACTIVELY when setting up new projects, after team feedback, or when... |
| `e2e-testing` | End-to-end testing workflow with Playwright for browser automation, visual regression, cross-browser testing, and CI/CD integration. |
| `expo-cicd-workflows` | Helps understand and write EAS workflow YAML files for Expo projects. Use this skill when the user asks about CI/CD or workflows in an Expo or EAS ... |
| `fastapi-pro` | Build high-performance async APIs with FastAPI, SQLAlchemy 2.0, and Pydantic V2. Master microservices, WebSockets, and modern Python async patterns. |
| `fastapi-templates` | Create production-ready FastAPI projects with async patterns, dependency injection, and comprehensive error handling. Use when building new FastAPI... |
| `finishing-a-development-branch` | Use when implementation is complete, all tests pass, and you need to decide how to integrate the work - guides completion of development work by pr... |
| `fp-async` | Practical async patterns using TaskEither - clean pipelines instead of try/catch hell, with real API examples |
| `fp-taskeither-ref` | Quick reference for TaskEither. Use when user needs async error handling, API calls, or Promise-based operations that can fail. |
| `freshdesk-automation` | Automate Freshdesk helpdesk operations including tickets, contacts, companies, notes, and replies via Rube MCP (Composio). Always search tools firs... |
| `freshservice-automation` | Automate Freshservice ITSM tasks via Rube MCP (Composio): create/update tickets, bulk operations, service requests, and outbound emails. Always sea... |
| `game-art` | Game art principles. Visual style selection, asset pipeline, animation workflow. |
| `git-advanced-workflows` | Master advanced Git workflows including rebasing, cherry-picking, bisect, worktrees, and reflog to maintain clean history and recover from any situ... |
| `git-hooks-automation` | Master Git hooks setup with Husky, lint-staged, pre-commit framework, and commitlint. Automate code quality gates, formatting, linting, and commit ... |
| `git-pr-workflows-onboard` | You are an **expert onboarding specialist and knowledge transfer architect** with deep experience in remote-first organizations, technical team int... |
| `git-pr-workflows-pr-enhance` | You are a PR optimization expert specializing in creating high-quality pull requests that facilitate efficient code reviews. Generate comprehensive... |
| `github-actions-templates` | Create production-ready GitHub Actions workflows for automated testing, building, and deploying applications. Use when setting up CI/CD with GitHub... |
| `github-automation` | Automate GitHub repositories, issues, pull requests, branches, CI/CD, and permissions via Rube MCP (Composio). Manage code workflows, review PRs, s... |
| `gitlab-automation` | Automate GitLab project management, issues, merge requests, pipelines, branches, and user operations via Rube MCP (Composio). Always search tools f... |
| `gitlab-ci-patterns` | Build GitLab CI/CD pipelines with multi-stage workflows, caching, and distributed runners for scalable automation. Use when implementing GitLab CI/... |
| `gitops-workflow` | Implement GitOps workflows with ArgoCD and Flux for automated, declarative Kubernetes deployments with continuous reconciliation. Use when implemen... |
| `gmail-automation` | Interact with Gmail - search emails, read messages, send emails, create drafts, and manage labels. Use when user asks to: search email, read email,... |
| `go-concurrency-patterns` | Master Go concurrency with goroutines, channels, sync primitives, and context. Use when building concurrent Go applications, implementing worker po... |
| `go-playwright` | Expert capability for robust, stealthy, and efficient browser automation using Playwright Go. |
| `go-rod-master` | Comprehensive guide for browser automation and web scraping with go-rod (Chrome DevTools Protocol) including stealth anti-bot-detection patterns. |
| `google-analytics-automation` | Automate Google Analytics tasks via Rube MCP (Composio): run reports, list accounts/properties, funnels, pivots, key events. Always search tools fi... |
| `google-calendar-automation` | Interact with Google Calendar - list calendars, view events, create/update/delete events, and find free time. Use when user asks to: check calendar... |
| `google-docs-automation` | Interact with Google Docs - create documents, search by title, read content, and edit text. Use when user asks to: create a Google Doc, find a docu... |
| `google-drive-automation` | Interact with Google Drive - search files, find folders, list contents, download files, upload files, create folders, move, copy, rename, and trash... |
| `google-sheets-automation` | Read and write Google Sheets spreadsheets - get content, update cells, append rows, fetch specific ranges, search for spreadsheets, and view metada... |
| `google-slides-automation` | Read and write Google Slides presentations - get text, find presentations, create presentations, add slides, replace text, and manage slide content... |
| `googlesheets-automation` | Automate Google Sheets operations (read, write, format, filter, manage spreadsheets) via Rube MCP (Composio). Read/write data, manage tabs, apply f... |
| `helpdesk-automation` | Automate HelpDesk tasks via Rube MCP (Composio): list tickets, manage views, use canned responses, and configure custom fields. Always search tools... |
| `hubspot-automation` | Automate HubSpot CRM operations (contacts, companies, deals, tickets, properties) via Rube MCP using Composio integration. |
| `hugging-face-dataset-viewer` | Use this skill for Hugging Face Dataset Viewer API workflows that fetch subset/split metadata, paginate rows, search text, apply filters, download ... |
| `impress` | Presentation creation, format conversion (ODP/PPTX/PDF), slide automation with LibreOffice Impress. |
| `inngest` | Inngest expert for serverless-first background jobs, event-driven workflows, and durable execution without managing queues or workers. Use when: in... |
| `instagram-automation` | Automate Instagram tasks via Rube MCP (Composio): create posts, carousels, manage media, get insights, and publishing limits. Always search tools f... |
| `intercom-automation` | Automate Intercom tasks via Rube MCP (Composio): conversations, contacts, companies, segments, admins. Always search tools first for current schemas. |
| `javascript-mastery` | Comprehensive JavaScript reference covering 33+ essential concepts every developer should know. From fundamentals like primitives and closures to a... |
| `javascript-pro` | Master modern JavaScript with ES6+, async patterns, and Node.js APIs. Handles promises, event loops, and browser/Node compatibility. |
| `jira-automation` | Automate Jira tasks via Rube MCP (Composio): issues, projects, sprints, boards, comments, users. Always search tools first for current schemas. |
| `k6-load-testing` | Comprehensive k6 load testing skill for API, browser, and scalability testing. Write realistic load scenarios, analyze results, and integrate with ... |
| `klaviyo-automation` | Automate Klaviyo tasks via Rube MCP (Composio): manage email/SMS campaigns, inspect campaign messages, track tags, and monitor send jobs. Always se... |
| `linear-automation` | Automate Linear tasks via Rube MCP (Composio): issues, projects, cycles, teams, labels. Always search tools first for current schemas. |
| `linkedin-automation` | Automate LinkedIn tasks via Rube MCP (Composio): create posts, manage profile, company info, comments, and image uploads. Always search tools first... |
| `linux-troubleshooting` | Linux system troubleshooting workflow for diagnosing and resolving system issues, performance problems, and service failures. |
| `magic-ui-generator` | Utilizes Magic by 21st.dev to generate, compare, and integrate multiple production-ready UI component variations. |
| `mailchimp-automation` | Automate Mailchimp email marketing including campaigns, audiences, subscribers, segments, and analytics via Rube MCP (Composio). Always search tool... |
| `make-automation` | Automate Make (Integromat) tasks via Rube MCP (Composio): operations, enums, language and timezone lookups. Always search tools first for current s... |
| `matplotlib` | Low-level plotting library for full customization. Use when you need fine-grained control over every plot element, creating novel plot types, or in... |
| `microsoft-teams-automation` | Automate Microsoft Teams tasks via Rube MCP (Composio): send messages, manage channels, create meetings, handle chats, and search messages. Always ... |
| `miro-automation` | Automate Miro tasks via Rube MCP (Composio): boards, items, sticky notes, frames, sharing, connectors. Always search tools first for current schemas. |
| `mixpanel-automation` | Automate Mixpanel tasks via Rube MCP (Composio): events, segmentation, funnels, cohorts, user profiles, JQL queries. Always search tools first for ... |
| `mobile-developer` | Develop React Native, Flutter, or native mobile apps with modern architecture patterns. Masters cross-platform development, native integrations, of... |
| `modern-javascript-patterns` | Master ES6+ features including async/await, destructuring, spread operators, arrow functions, promises, modules, iterators, generators, and functio... |
| `monday-automation` | Automate Monday.com work management including boards, items, columns, groups, subitems, and updates via Rube MCP (Composio). Always search tools fi... |
| `multiplayer` | Multiplayer game development principles. Architecture, networking, synchronization. |
| `n8n-mcp-tools-expert` | Expert guide for using n8n-mcp MCP tools effectively. Use when searching for nodes, validating configurations, accessing templates, managing workfl... |
| `n8n-workflow-patterns` | Proven workflow architectural patterns from real n8n workflows. Use when building new workflows, designing workflow structure, choosing workflow pa... |
| `nerdzao-elite` | Senior Elite Software Engineer (15+) and Senior Product Designer. Full workflow with planning, architecture, TDD, clean code, and pixel-perfect UX ... |
| `notion-automation` | Automate Notion tasks via Rube MCP (Composio): pages, databases, blocks, comments, users. Always search tools first for current schemas. |
| `notion-template-business` | Expert in building and selling Notion templates as a business - not just making templates, but building a sustainable digital product business. Cov... |
| `observability-engineer` | Build production-ready monitoring, logging, and tracing systems. Implements comprehensive observability strategies, SLI/SLO management, and inciden... |
| `odoo-manufacturing-advisor` | Expert guide for Odoo Manufacturing: Bills of Materials (BoM), Work Centers, routings, MRP planning, and production order workflows. |
| `odoo-purchase-workflow` | Expert guide for Odoo Purchase: RFQ → PO → Receipt → Vendor Bill workflow, purchase agreements, vendor price lists, and 3-way matching. |
| `odoo-shopify-integration` | Connect Odoo with Shopify: sync products, inventory, orders, and customers using the Shopify API and Odoo's external API or connector modules. |
| `odoo-woocommerce-bridge` | Sync Odoo with WooCommerce: products, inventory, orders, and customers via WooCommerce REST API and Odoo external API. |
| `office-productivity` | Office productivity workflow covering document creation, spreadsheet automation, presentation generation, and integration with LibreOffice and Micr... |
| `one-drive-automation` | Automate OneDrive file management, search, uploads, downloads, sharing, permissions, and folder operations via Rube MCP (Composio). Always search t... |
| `os-scripting` | Operating system and shell scripting troubleshooting workflow for Linux, macOS, and Windows. Covers bash scripting, system administration, debuggin... |
| `outlook-automation` | Automate Outlook tasks via Rube MCP (Composio): emails, calendar, contacts, folders, attachments. Always search tools first for current schemas. |
| `outlook-calendar-automation` | Automate Outlook Calendar tasks via Rube MCP (Composio): create events, manage attendees, find meeting times, and handle invitations. Always search... |
| `pagerduty-automation` | Automate PagerDuty tasks via Rube MCP (Composio): manage incidents, services, schedules, escalation policies, and on-call rotations. Always search ... |
| `pakistan-payments-stack` | Design and implement production-grade Pakistani payment integrations (JazzCash, Easypaisa, bank/PSP rails, optional Raast) for SaaS with PKR billin... |
| `paypal-integration` | Integrate PayPal payment processing with support for express checkout, subscriptions, and refund management. Use when implementing PayPal payments,... |
| `pipedrive-automation` | Automate Pipedrive CRM operations including deals, contacts, organizations, activities, notes, and pipeline management via Rube MCP (Composio). Alw... |
| `playwright-skill` | Complete browser automation with Playwright. Auto-detects dev servers, writes clean test scripts to /tmp. Test pages, fill forms, take screenshots,... |
| `postgresql-optimization` | PostgreSQL database optimization workflow for query tuning, indexing strategies, performance analysis, and production database management. |
| `posthog-automation` | Automate PostHog tasks via Rube MCP (Composio): events, feature flags, projects, user profiles, annotations. Always search tools first for current ... |
| `postmark-automation` | Automate Postmark email delivery tasks via Rube MCP (Composio): send templated emails, manage templates, monitor delivery stats and bounces. Always... |
| `pubmed-database` | Direct REST API access to PubMed. Advanced Boolean/MeSH queries, E-utilities API, batch processing, citation management. For Python workflows, pref... |
| `python-patterns` | Python development principles and decision-making. Framework selection, async patterns, type hints, project structure. Teaches thinking, not copying. |
| `python-pro` | Master Python 3.12+ with modern features, async programming, performance optimization, and production-ready practices. Expert in the latest Python ... |
| `react-flow-node-ts` | Create React Flow node components with TypeScript types, handles, and Zustand integration. Use when building custom nodes for React Flow canvas, cr... |
| `react-native-architecture` | Build production React Native apps with Expo, navigation, native modules, offline sync, and cross-platform patterns. Use when developing mobile app... |
| `react-ui-patterns` | Modern React UI patterns for loading states, error handling, and data fetching. Use when building UI components, handling async data, or managing U... |
| `reddit-automation` | Automate Reddit tasks via Rube MCP (Composio): search subreddits, create posts, manage comments, and browse top content. Always search tools first ... |
| `render-automation` | Automate Render tasks via Rube MCP (Composio): services, deployments, projects. Always search tools first for current schemas. |
| `robius-app-architecture` | CRITICAL: Use for Robius app architecture patterns. Triggers on: Tokio, async, submit_async_request, 异步, 架构, SignalToUI, Cx::post_action, worker ta... |
| `robius-matrix-integration` | CRITICAL: Use for Matrix SDK integration with Makepad. Triggers on: Matrix SDK, sliding sync, MatrixRequest, timeline, matrix-sdk, matrix client, r... |
| `rust-async-patterns` | Master Rust async programming with Tokio, async traits, error handling, and concurrent patterns. Use when building async Rust applications, impleme... |
| `rust-pro` | Master Rust 1.75+ with modern async patterns, advanced type system features, and production-ready systems programming. |
| `salesforce-automation` | Automate Salesforce tasks via Rube MCP (Composio): leads, contacts, accounts, opportunities, SOQL queries. Always search tools first for current sc... |
| `salesforce-development` | Expert patterns for Salesforce platform development including Lightning Web Components (LWC), Apex triggers and classes, REST/Bulk APIs, Connected ... |
| `scanpy` | Standard single-cell RNA-seq analysis pipeline. Use for QC, normalization, dimensionality reduction (PCA/UMAP/t-SNE), clustering, differential expr... |
| `segment-automation` | Automate Segment tasks via Rube MCP (Composio): track events, identify users, manage groups, page views, aliases, batch operations. Always search t... |
| `sendgrid-automation` | Automate SendGrid email operations including sending emails, managing contacts/lists, sender identities, templates, and analytics via Rube MCP (Com... |
| `sentry-automation` | Automate Sentry tasks via Rube MCP (Composio): manage issues/events, configure alerts, track releases, monitor projects and teams. Always search to... |
| `shopify-apps` | Expert patterns for Shopify app development including Remix/React Router apps, embedded apps with App Bridge, webhook handling, GraphQL Admin API, ... |
| `shopify-automation` | Automate Shopify tasks via Rube MCP (Composio): products, orders, customers, inventory, collections. Always search tools first for current schemas. |
| `slack-automation` | Automate Slack messaging, channel management, search, reactions, and threads via Rube MCP (Composio). Send messages, search conversations, manage c... |
| `slack-gif-creator` | Knowledge and utilities for creating animated GIFs optimized for Slack. Provides constraints, validation tools, and animation concepts. Use when us... |
| `square-automation` | Automate Square tasks via Rube MCP (Composio): payments, orders, invoices, locations. Always search tools first for current schemas. |
| `sred-work-summary` | Go back through the previous year of work and create a Notion doc that groups relevant links into projects that can then be documented as SRED proj... |
| `stripe-automation` | Automate Stripe tasks via Rube MCP (Composio): customers, charges, subscriptions, invoices, products, refunds. Always search tools first for curren... |
| `stripe-integration` | Implement Stripe payment processing for robust, PCI-compliant payment flows including checkout, subscriptions, and webhooks. Use when integrating S... |
| `tanstack-query-expert` | Expert in TanStack Query (React Query) — asynchronous state management. Covers data fetching, stale time configuration, mutations, optimistic updat... |
| `tdd-workflow` | Test-Driven Development workflow principles. RED-GREEN-REFACTOR cycle. |
| `tdd-workflows-tdd-cycle` | Use when working with tdd workflows tdd cycle |
| `tdd-workflows-tdd-green` | Implement the minimal code needed to make failing tests pass in the TDD green phase. |
| `tdd-workflows-tdd-red` | Generate failing tests for the TDD red phase to define expected behavior and edge cases. |
| `tdd-workflows-tdd-refactor` | Use when working with tdd workflows tdd refactor |
| `team-collaboration-issue` | You are a GitHub issue resolution expert specializing in systematic bug investigation, feature implementation, and collaborative development workfl... |
| `telegram` | Integracao completa com Telegram Bot API. Setup com BotFather, mensagens, webhooks, inline keyboards, grupos, canais. Boilerplates Node.js e Python. |
| `telegram-automation` | Automate Telegram tasks via Rube MCP (Composio): send messages, manage chats, share photos/documents, and handle bot commands. Always search tools ... |
| `temporal-golang-pro` | Use when building durable distributed systems with Temporal Go SDK. Covers deterministic workflow rules, mTLS worker configs, and advanced patterns. |
| `temporal-python-testing` | Test Temporal workflows with pytest, time-skipping, and mocking strategies. Covers unit testing, integration testing, replay testing, and local dev... |
| `terraform-infrastructure` | Terraform infrastructure as code workflow for provisioning cloud resources, creating reusable modules, and managing infrastructure at scale. |
| `terraform-specialist` | Expert Terraform/OpenTofu specialist mastering advanced IaC automation, state management, and enterprise infrastructure patterns. |
| `testing-patterns` | Jest testing patterns, factory functions, mocking strategies, and TDD workflow. Use when writing unit tests, creating test factories, or following ... |
| `testing-qa` | Comprehensive testing and QA workflow covering unit testing, integration testing, E2E testing, browser automation, and quality assurance. |
| `tiktok-automation` | Automate TikTok tasks via Rube MCP (Composio): upload/publish videos, post photos, manage content, and view user profiles/stats. Always search tool... |
| `todoist-automation` | Automate Todoist task management, projects, sections, filtering, and bulk operations via Rube MCP (Composio). Always search tools first for current... |
| `trello-automation` | Automate Trello boards, cards, and workflows via Rube MCP (Composio). Create cards, manage lists, assign members, and search across boards programm... |
| `twitter-automation` | Automate Twitter/X tasks via Rube MCP (Composio): posts, search, users, bookmarks, lists, media. Always search tools first for current schemas. |
| `uv-package-manager` | Master the uv package manager for fast Python dependency management, virtual environments, and modern Python project workflows. Use when setting up... |
| `vercel-automation` | Automate Vercel tasks via Rube MCP (Composio): manage deployments, domains, DNS, env vars, projects, and teams. Always search tools first for curre... |
| `webflow-automation` | Automate Webflow CMS collections, site publishing, page management, asset uploads, and ecommerce orders via Rube MCP (Composio). Always search tool... |
| `wellally-tech` | Integrate digital health data sources (Apple Health, Fitbit, Oura Ring) and connect to WellAlly.tech knowledge base. Import external health device ... |
| `whatsapp-automation` | Automate WhatsApp Business tasks via Rube MCP (Composio): send messages, manage templates, upload media, and handle contacts. Always search tools f... |
| `whatsapp-cloud-api` | Integracao com WhatsApp Business Cloud API (Meta). Mensagens, templates, webhooks HMAC-SHA256, automacao de atendimento. Boilerplates Node.js e Pyt... |
| `wordpress-theme-development` | WordPress theme development workflow covering theme architecture, template hierarchy, custom post types, block editor support, and responsive design. |
| `wordpress-woocommerce-development` | WooCommerce store development workflow covering store setup, payment integration, shipping configuration, and customization. |
| `workflow-patterns` | Use this skill when implementing tasks according to Conductor's TDD workflow, handling phase checkpoints, managing git commits for tasks, or unders... |
| `wrike-automation` | Automate Wrike project management via Rube MCP (Composio): create tasks/folders, manage projects, assign work, and track progress. Always search to... |
| `writer` | Document creation, format conversion (ODT/DOCX/PDF), mail merge, and automation with LibreOffice Writer. |
| `x-twitter-scraper` | X (Twitter) data platform skill — tweet search, user lookup, follower extraction, engagement metrics, giveaway draws, monitoring, webhooks, 19 extr... |
| `youtube-automation` | Automate YouTube tasks via Rube MCP (Composio): upload videos, manage playlists, search content, get analytics, and handle comments. Always search ... |
| `zapier-make-patterns` | No-code automation democratizes workflow building. Zapier and Make (formerly Integromat) let non-developers automate business processes without wri... |
| `zendesk-automation` | Automate Zendesk tasks via Rube MCP (Composio): tickets, users, organizations, replies. Always search tools first for current schemas. |
| `zoho-crm-automation` | Automate Zoho CRM tasks via Rube MCP (Composio): create/update records, search contacts, manage leads, and convert leads. Always search tools first... |
| `zoom-automation` | Automate Zoom meeting creation, management, recordings, webinars, and participant tracking via Rube MCP (Composio). Always search tools first for c... |

[↑ Voltar ao Índice](#indice-de-categorias)

---

## 📁 Backend & Cloud

| ID | Descrição |
| :--- | :--- |
| `algolia-search` | Expert patterns for Algolia search implementation, indexing strategies, React InstantSearch, and relevance tuning Use when: adding search to, algol... |
| `android_ui_verification` | Automated end-to-end UI testing and verification on an Android Emulator using ADB. |
| `api-design-principles` | Master REST and GraphQL API design principles to build intuitive, scalable, and maintainable APIs that delight developers. Use when designing new A... |
| `api-documentation-generator` | Generate comprehensive, developer-friendly API documentation from code, including endpoints, parameters, examples, and best practices |
| `api-patterns` | API design principles and decision-making. REST vs GraphQL vs tRPC selection, response formats, versioning, pagination. |
| `apify-audience-analysis` | Understand audience demographics, preferences, behavior patterns, and engagement quality across Facebook, Instagram, YouTube, and TikTok. |
| `apify-brand-reputation-monitoring` | Track reviews, ratings, sentiment, and brand mentions across Google Maps, Booking.com, TripAdvisor, Facebook, Instagram, YouTube, and TikTok. Use w... |
| `apify-competitor-intelligence` | Analyze competitor strategies, content, pricing, ads, and market positioning across Google Maps, Booking.com, Facebook, Instagram, YouTube, and Tik... |
| `apify-market-research` | Analyze market conditions, geographic opportunities, pricing, consumer behavior, and product validation across Google Maps, Facebook, Instagram, Bo... |
| `apify-trend-analysis` | Discover and track emerging trends across Google Trends, Instagram, Facebook, YouTube, and TikTok to inform content strategy. |
| `architecture-patterns` | Implement proven backend architecture patterns including Clean Architecture, Hexagonal Architecture, and Domain-Driven Design. Use when architectin... |
| `aws-cost-cleanup` | Automated cleanup of unused AWS resources to reduce costs |
| `aws-cost-optimizer` | Comprehensive AWS cost analysis and optimization recommendations using AWS CLI and Cost Explorer |
| `aws-serverless` | Specialized skill for building production-ready serverless applications on AWS. Covers Lambda functions, API Gateway, DynamoDB, SQS/SNS event-drive... |
| `azure-ai-translation-ts` | Build translation applications using Azure Translation SDKs for JavaScript (@azure-rest/ai-translation-text, @azure-rest/ai-translation-document). ... |
| `azure-appconfiguration-java` | Azure App Configuration SDK for Java. Centralized application configuration management with key-value settings, feature flags, and snapshots. |
| `azure-appconfiguration-ts` | Build applications using Azure App Configuration SDK for JavaScript (@azure/app-configuration). Use when working with configuration settings, featu... |
| `azure-communication-chat-java` | Build real-time chat applications with Azure Communication Services Chat Java SDK. Use when implementing chat threads, messaging, participants, rea... |
| `azure-communication-sms-java` | Send SMS messages with Azure Communication Services SMS Java SDK. Use when implementing SMS notifications, alerts, OTP delivery, bulk messaging, or... |
| `azure-compute-batch-java` | Azure Batch SDK for Java. Run large-scale parallel and HPC batch jobs with pools, jobs, tasks, and compute nodes. |
| `azure-eventgrid-dotnet` | Azure Event Grid SDK for .NET. Client library for publishing and consuming events with Azure Event Grid. Use for event-driven architectures, pub/su... |
| `azure-eventgrid-java` | Build event-driven applications with Azure Event Grid SDK for Java. Use when publishing events, implementing pub/sub patterns, or integrating with ... |
| `azure-eventhub-dotnet` | Azure Event Hubs SDK for .NET. |
| `azure-keyvault-certificates-rust` | Azure Key Vault Certificates SDK for Rust. Use for creating, importing, and managing certificates. |
| `azure-messaging-webpubsub-java` | Build real-time web applications with Azure Web PubSub SDK for Java. Use when implementing WebSocket-based messaging, live updates, chat applicatio... |
| `azure-mgmt-apicenter-dotnet` | Azure API Center SDK for .NET. Centralized API inventory management with governance, versioning, and discovery. |
| `azure-mgmt-apimanagement-dotnet` | Azure Resource Manager SDK for API Management in .NET. |
| `azure-mgmt-fabric-dotnet` | Azure Resource Manager SDK for Fabric in .NET. |
| `azure-mgmt-mongodbatlas-dotnet` | Manage MongoDB Atlas Organizations as Azure ARM resources using Azure.ResourceManager.MongoDBAtlas SDK. Use when creating, updating, listing, or de... |
| `azure-microsoft-playwright-testing-ts` | Run Playwright tests at scale using Azure Playwright Workspaces (formerly Microsoft Playwright Testing). Use when scaling browser tests across clou... |
| `azure-monitor-opentelemetry-exporter-java` | Azure Monitor OpenTelemetry Exporter for Java. Export OpenTelemetry traces, metrics, and logs to Azure Monitor/Application Insights. |
| `azure-monitor-opentelemetry-ts` | Instrument applications with Azure Monitor and OpenTelemetry for JavaScript (@azure/monitor-opentelemetry). Use when adding distributed tracing, me... |
| `azure-resource-manager-cosmosdb-dotnet` | Azure Resource Manager SDK for Cosmos DB in .NET. |
| `azure-resource-manager-durabletask-dotnet` | Azure Resource Manager SDK for Durable Task Scheduler in .NET. |
| `azure-resource-manager-playwright-dotnet` | Azure Resource Manager SDK for Microsoft Playwright Testing in .NET. |
| `azure-resource-manager-redis-dotnet` | Azure Resource Manager SDK for Redis in .NET. |
| `azure-servicebus-dotnet` | Azure Service Bus SDK for .NET. Enterprise messaging with queues, topics, subscriptions, and sessions. |
| `azure-servicebus-ts` | Build messaging applications using Azure Service Bus SDK for JavaScript (@azure/service-bus). Use when implementing queues, topics/subscriptions, m... |
| `azure-web-pubsub-ts` | Build real-time messaging applications using Azure Web PubSub SDKs for JavaScript (@azure/web-pubsub, @azure/web-pubsub-client). Use when implement... |
| `backend-architect` | Expert backend architect specializing in scalable API design, microservices architecture, and distributed systems. |
| `backend-dev-guidelines` | Opinionated backend development standards for Node.js + Express + TypeScript microservices. Covers layered architecture, BaseController pattern, de... |
| `bash-linux` | Bash/Linux terminal patterns. Critical commands, piping, error handling, scripting. Use when working on macOS or Linux systems. |
| `bun-development` | Modern JavaScript/TypeScript development with Bun runtime. Covers package management, bundling, testing, and migration from Node.js. Use when worki... |
| `cc-skill-coding-standards` | Universal coding standards, best practices, and patterns for TypeScript, JavaScript, React, and Node.js development. |
| `closed-loop-delivery` | Use when a coding task must be completed against explicit acceptance criteria with minimal user re-intervention across implementation, review feedb... |
| `cloud-architect` | Expert cloud architect specializing in AWS/Azure/GCP multi-cloud infrastructure design, advanced IaC (Terraform/OpenTofu/CDK), FinOps cost optimiza... |
| `cloudformation-best-practices` | CloudFormation template optimization, nested stacks, drift detection, and production-ready patterns. Use when writing or reviewing CF templates. |
| `code-review-excellence` | Master effective code review practices to provide constructive feedback, catch bugs early, and foster knowledge sharing while maintaining team mora... |
| `copy-editing` | When the user wants to edit, review, or improve existing marketing copy. Also use when the user mentions 'edit this copy,' 'review my copy,' 'copy ... |
| `cost-optimization` | Optimize cloud costs through resource rightsizing, tagging strategies, reserved instances, and spending analysis. Use when reducing cloud expenses,... |
| `deployment-validation-config-validate` | You are a configuration management expert specializing in validating, testing, and ensuring the correctness of application configurations. Create c... |
| `devops-deploy` | DevOps e deploy de aplicacoes — Docker, CI/CD com GitHub Actions, AWS Lambda, SAM, Terraform, infraestrutura como codigo e monitoramento. Ativar pa... |
| `dotnet-architect` | Expert .NET backend architect specializing in C#, ASP.NET Core, Entity Framework, Dapper, and enterprise application patterns. |
| `e2e-testing-patterns` | Master end-to-end testing with Playwright and Cypress to build reliable test suites that catch bugs, improve confidence, and enable fast deployment... |
| `event-store-design` | Design and implement event stores for event-sourced systems. Use when building event sourcing infrastructure, choosing event store technologies, or... |
| `exa-search` | Semantic search, similar content discovery, and structured research using Exa API |
| `expo-api-routes` | Guidelines for creating API routes in Expo Router with EAS Hosting |
| `expo-deployment` | Deploy Expo apps to production |
| `flutter-expert` | Master Flutter development with Dart 3, advanced widgets, and multi-platform deployment. |
| `fp-backend` | Functional programming patterns for Node.js/Deno backend development using fp-ts, ReaderTaskEither, and functional dependency injection |
| `gcp-cloud-run` | Specialized skill for building production-ready serverless applications on GCP. Covers Cloud Run services (containerized), Cloud Run Functions (eve... |
| `helm-chart-scaffolding` | Design, organize, and manage Helm charts for templating and packaging Kubernetes applications with reusable configurations. Use when creating Helm ... |
| `hybrid-cloud-architect` | Expert hybrid cloud architect specializing in complex multi-cloud solutions across AWS/Azure/GCP and private clouds (OpenStack/VMware). |
| `hybrid-cloud-networking` | Configure secure, high-performance connectivity between on-premises infrastructure and cloud platforms using VPN and dedicated connections. Use whe... |
| `istio-traffic-management` | Configure Istio traffic management including routing, load balancing, circuit breakers, and canary deployments. Use when implementing service mesh ... |
| `iterate-pr` | Iterate on a PR until CI passes. Use when you need to fix CI failures, address review feedback, or continuously push fixes until all checks are gre... |
| `java-pro` | Master Java 21+ with modern features like virtual threads, pattern matching, and Spring Boot 3.x. Expert in the latest Java ecosystem including Gra... |
| `javascript-typescript-typescript-scaffold` | You are a TypeScript project architecture expert specializing in scaffolding production-ready Node.js and frontend applications. Generate complete ... |
| `linux-shell-scripting` | This skill should be used when the user asks to "create bash scripts", "automate Linux tasks", "monitor system resources", "backup files", "manage ... |
| `makepad-deployment` | CRITICAL: Use for Makepad packaging and deployment. Triggers on: deploy, package, APK, IPA, 打包, 部署, cargo-packager, cargo-makepad, WASM, Android, i... |
| `makepad-font` | CRITICAL: Use for Makepad font and text rendering. Triggers on: makepad font, makepad text, makepad glyph, makepad typography, font atlas, text lay... |
| `makepad-platform` | CRITICAL: Use for Makepad cross-platform support. Triggers on: makepad platform, makepad os, makepad macos, makepad windows, makepad linux, makepad... |
| `makepad-reference` | CRITICAL: Use for Makepad troubleshooting and reference. Triggers on: troubleshoot, error, debug, fix, problem, issue, no matching field, parse err... |
| `minecraft-bukkit-pro` | Master Minecraft server plugin development with Bukkit, Spigot, and Paper APIs. |
| `moodle-external-api-development` | Create custom external web service APIs for Moodle LMS. Use when implementing web services for course management, user tracking, quiz operations, o... |
| `n8n-node-configuration` | Operation-aware node configuration guidance. Use when configuring nodes, understanding property dependencies, determining required fields, choosing... |
| `network-101` | This skill should be used when the user asks to "set up a web server", "configure HTTP or HTTPS", "perform SNMP enumeration", "configure SMB shares... |
| `odoo-migration-helper` | Step-by-step guide for migrating Odoo custom modules between versions (v14→v15→v16→v17). Covers API changes, deprecated methods, and view migration. |
| `react-nextjs-development` | React and Next.js 14+ application development with App Router, Server Components, TypeScript, Tailwind CSS, and modern frontend patterns. |
| `react-state-management` | Master modern React state management with Redux Toolkit, Zustand, Jotai, and React Query. Use when setting up global state, managing server state, ... |
| `receiving-code-review` | Use when receiving code review feedback, before implementing suggestions, especially if feedback seems unclear or technically questionable - requir... |
| `reference-builder` | Creates exhaustive technical references and API documentation. Generates comprehensive parameter listings, configuration guides, and searchable ref... |
| `server-management` | Server management principles and decision-making. Process management, monitoring strategy, and scaling decisions. Teaches thinking, not commands. |
| `shopify-development` | Build Shopify apps, extensions, themes using GraphQL Admin API, Shopify CLI, Polaris UI, and Liquid. |
| `terraform-aws-modules` | Terraform module creation for AWS — reusable modules, state management, and HCL best practices. Use when building or reviewing Terraform AWS infras... |
| `terraform-module-library` | Build reusable Terraform modules for AWS, Azure, and GCP infrastructure following infrastructure-as-code best practices. Use when creating infrastr... |
| `terraform-skill` | Terraform infrastructure as code best practices |
| `twilio-communications` | Build communication features with Twilio: SMS messaging, voice calls, WhatsApp Business API, and user verification (2FA). Covers the full spectrum ... |
| `unity-developer` | Build Unity games with optimized C# scripts, efficient rendering, and proper asset management. Masters Unity 6 LTS, URP/HDRP pipelines, and cross-p... |
| `upstash-qstash` | Upstash QStash expert for serverless message queues, scheduled jobs, and reliable HTTP-based task delivery without managing infrastructure. Use whe... |
| `vercel-deployment` | Expert knowledge for deploying to Vercel with Next.js Use when: vercel, deploy, deployment, hosting, production. |
| `videodb` | Video and audio perception, indexing, and editing. Ingest files/URLs/live streams, build visual/spoken indexes, search with timestamps, edit timeli... |

[↑ Voltar ao Índice](#indice-de-categorias)

---

## 📁 Data & Analytics

| ID | Descrição |
| :--- | :--- |
| `10-andruia-skill-smith` | Ingeniero de Sistemas de Andru.ia. Diseña, redacta y despliega nuevas habilidades (skills) dentro del repositorio siguiendo el Estándar de Diamante. |
| `analytics-product` | Analytics de produto — PostHog, Mixpanel, eventos, funnels, cohorts, retencao, north star metric, OKRs e dashboards de produto. Ativar para: config... |
| `api-testing-observability-api-mock` | You are an API mocking expert specializing in realistic mock services for development, testing, and demos. Design mocks that simulate real API beha... |
| `apify-content-analytics` | Track engagement metrics, measure campaign ROI, and analyze content performance across Instagram, Facebook, YouTube, and TikTok. |
| `apify-ecommerce` | Scrape e-commerce data for pricing intelligence, customer reviews, and seller discovery across Amazon, Walmart, eBay, IKEA, and 50+ marketplaces. U... |
| `apify-lead-generation` | Generates B2B/B2C leads by scraping Google Maps, websites, Instagram, TikTok, Facebook, LinkedIn, YouTube, and Google Search. Use when user asks to... |
| `app-store-optimization` | Complete App Store Optimization (ASO) toolkit for researching, optimizing, and tracking mobile app performance on Apple App Store and Google Play S... |
| `application-performance-performance-optimization` | Optimize end-to-end application performance with profiling, observability, and backend/frontend tuning. Use when coordinating performance optimizat... |
| `astropy` | Comprehensive Python library for astronomy and astrophysics. This skill should be used when working with astronomical data including celestial coor... |
| `azd-deployment` | Deploy containerized applications to Azure Container Apps using Azure Developer CLI (azd). Use when setting up azd projects, writing azure.yaml con... |
| `azure-ai-document-intelligence-ts` | Extract text, tables, and structured data from documents using Azure Document Intelligence (@azure-rest/ai-document-intelligence). Use when process... |
| `azure-ai-formrecognizer-java` | Build document analysis applications with Azure Document Intelligence (Form Recognizer) SDK for Java. Use when extracting text, tables, key-value p... |
| `azure-appconfiguration-py` | Azure App Configuration SDK for Python. Use for centralized configuration management, feature flags, and dynamic settings. |
| `azure-containerregistry-py` | Azure Container Registry SDK for Python. Use for managing container images, artifacts, and repositories. |
| `azure-cosmos-py` | Azure Cosmos DB SDK for Python (NoSQL API). Use for document CRUD, queries, containers, and globally distributed data. |
| `azure-cosmos-rust` | Azure Cosmos DB SDK for Rust (NoSQL API). Use for document CRUD, queries, containers, and globally distributed data. |
| `azure-cosmos-ts` | Azure Cosmos DB JavaScript/TypeScript SDK (@azure/cosmos) for data plane operations. Use for CRUD operations on documents, queries, bulk operations... |
| `azure-eventgrid-py` | Azure Event Grid SDK for Python. Use for publishing events, handling CloudEvents, and event-driven architectures. |
| `azure-eventhub-java` | Build real-time streaming applications with Azure Event Hubs SDK for Java. Use when implementing event streaming, high-throughput data ingestion, o... |
| `azure-eventhub-py` | Azure Event Hubs SDK for Python streaming. Use for high-throughput event ingestion, producers, consumers, and checkpointing. |
| `azure-eventhub-rust` | Azure Event Hubs SDK for Rust. Use for sending and receiving events, streaming data ingestion. |
| `azure-eventhub-ts` | Build event streaming applications using Azure Event Hubs SDK for JavaScript (@azure/event-hubs). Use when implementing high-throughput event inges... |
| `azure-maps-search-dotnet` | Azure Maps SDK for .NET. Location-based services including geocoding, routing, rendering, geolocation, and weather. Use for address search, directi... |
| `azure-messaging-webpubsubservice-py` | Azure Web PubSub Service SDK for Python. Use for real-time messaging, WebSocket connections, and pub/sub patterns. |
| `azure-mgmt-apicenter-py` | Azure API Center Management SDK for Python. Use for managing API inventory, metadata, and governance across your organization. |
| `azure-mgmt-apimanagement-py` | Azure API Management SDK for Python. Use for managing APIM services, APIs, products, subscriptions, and policies. |
| `azure-mgmt-applicationinsights-dotnet` | Azure Application Insights SDK for .NET. Application performance monitoring and observability resource management. |
| `azure-mgmt-botservice-py` | Azure Bot Service Management SDK for Python. Use for creating, managing, and configuring Azure Bot Service resources. |
| `azure-mgmt-fabric-py` | Azure Fabric Management SDK for Python. Use for managing Microsoft Fabric capacities and resources. |
| `azure-monitor-ingestion-java` | Azure Monitor Ingestion SDK for Java. Send custom logs to Azure Monitor via Data Collection Rules (DCR) and Data Collection Endpoints (DCE). |
| `azure-monitor-ingestion-py` | Azure Monitor Ingestion SDK for Python. Use for sending custom logs to Log Analytics workspace via Logs Ingestion API. |
| `azure-monitor-opentelemetry-exporter-py` | Azure Monitor OpenTelemetry Exporter for Python. Use for low-level OpenTelemetry export to Application Insights. |
| `azure-monitor-opentelemetry-py` | Azure Monitor OpenTelemetry Distro for Python. Use for one-line Application Insights setup with auto-instrumentation. |
| `azure-monitor-query-java` | Azure Monitor Query SDK for Java. Execute Kusto queries against Log Analytics workspaces and query metrics from Azure resources. |
| `azure-monitor-query-py` | Azure Monitor Query SDK for Python. Use for querying Log Analytics workspaces and Azure Monitor metrics. |
| `azure-postgres-ts` | Connect to Azure Database for PostgreSQL Flexible Server from Node.js/TypeScript using the pg (node-postgres) package. |
| `azure-resource-manager-mysql-dotnet` | Azure MySQL Flexible Server SDK for .NET. Database management for MySQL Flexible Server deployments. |
| `azure-resource-manager-postgresql-dotnet` | Azure PostgreSQL Flexible Server SDK for .NET. Database management for PostgreSQL Flexible Server deployments. |
| `azure-resource-manager-sql-dotnet` | Azure Resource Manager SDK for Azure SQL in .NET. |
| `azure-servicebus-py` | Azure Service Bus SDK for Python messaging. Use for queues, topics, subscriptions, and enterprise messaging patterns. |
| `azure-speech-to-text-rest-py` | Azure Speech to Text REST API for short audio (Python). Use for simple speech recognition of audio files up to 60 seconds without the Speech SDK. |
| `backtesting-frameworks` | Build robust backtesting systems for trading strategies with proper handling of look-ahead bias, survivorship bias, and transaction costs. Use when... |
| `base` | Database management, forms, reports, and data operations with LibreOffice Base. |
| `baseline-ui` | Validates animation durations, enforces typography scale, checks component accessibility, and prevents layout anti-patterns in Tailwind CSS project... |
| `binary-analysis-patterns` | Master binary analysis patterns including disassembly, decompilation, control flow analysis, and code pattern recognition. Use when analyzing execu... |
| `cc-skill-backend-patterns` | Backend architecture patterns, API design, database optimization, and server-side best practices for Node.js, Express, and Next.js API routes. |
| `cc-skill-clickhouse-io` | ClickHouse database patterns, query optimization, analytics, and data engineering best practices for high-performance analytical workloads. |
| `cdk-patterns` | Common AWS CDK patterns and constructs for building cloud infrastructure with TypeScript, Python, or Java. Use when designing reusable CDK stacks a... |
| `citation-management` | Comprehensive citation management for academic research. Search Google Scholar and PubMed for papers, extract accurate metadata, validate citations... |
| `claimable-postgres` | Provision instant temporary Postgres databases via Claimable Postgres by Neon (pg.new). No login or credit card required. Use for quick Postgres en... |
| `clean-code` | Applies principles from Robert C. Martin's 'Clean Code'. Use this skill when writing, reviewing, or refactoring code to ensure high quality, readab... |
| `code-simplifier` | Simplifies and refines code for clarity, consistency, and maintainability while preserving all functionality. Use when asked to "simplify code", "c... |
| `copilot-sdk` | Build applications powered by GitHub Copilot using the Copilot SDK. Use when creating programmatic integrations with Copilot across Node.js/TypeScr... |
| `daily` | Documentation and capabilities reference for Daily |
| `daily-news-report` | Scrapes content based on a preset URL list, filters high-quality technical information, and generates daily Markdown reports. |
| `data-engineering-data-pipeline` | You are a data pipeline architecture expert specializing in scalable, reliable, and cost-effective data pipelines for batch and streaming data proc... |
| `data-quality-frameworks` | Implement data quality validation with Great Expectations, dbt tests, and data contracts. Use when building data quality pipelines, implementing va... |
| `data-storytelling` | Transform data into compelling narratives using visualization, context, and persuasive structure. Use when presenting analytics to stakeholders, cr... |
| `database-cloud-optimization-cost-optimize` | You are a cloud cost optimization expert specializing in reducing infrastructure expenses while maintaining performance and reliability. Analyze cl... |
| `database-design` | Database design principles and decision-making. Schema design, indexing strategy, ORM selection, serverless databases. |
| `database-migration` | Execute database migrations across ORMs and platforms with zero-downtime strategies, data transformation, and rollback procedures. Use when migrati... |
| `database-migrations-migration-observability` | Migration monitoring, CDC, and observability infrastructure |
| `database-migrations-sql-migrations` | SQL database migrations with zero-downtime strategies for PostgreSQL, MySQL, and SQL Server. Focus on data integrity and rollback plans. |
| `database-optimizer` | Expert database optimizer specializing in modern performance tuning, query optimization, and scalable architectures. |
| `ddd-strategic-design` | Design DDD strategic artifacts including subdomains, bounded contexts, and ubiquitous language for complex business domains. |
| `dependency-upgrade` | Manage major dependency version upgrades with compatibility analysis, staged rollout, and comprehensive testing. Use when upgrading framework versi... |
| `devops-troubleshooter` | Expert DevOps troubleshooter specializing in rapid incident response, advanced debugging, and modern observability. |
| `drizzle-orm-expert` | Expert in Drizzle ORM for TypeScript — schema design, relational queries, migrations, and serverless database integration. Use when building type-s... |
| `dwarf-expert` | Provides expertise for analyzing DWARF debug files and understanding the DWARF debug format/standard (v3-v5). Triggers when understanding DWARF inf... |
| `error-debugging-error-analysis` | You are an expert error analysis specialist with deep expertise in debugging distributed systems, analyzing production incidents, and implementing ... |
| `error-debugging-error-trace` | You are an error tracking and observability expert specializing in implementing comprehensive error monitoring solutions. Set up error tracking sys... |
| `error-diagnostics-error-analysis` | You are an expert error analysis specialist with deep expertise in debugging distributed systems, analyzing production incidents, and implementing ... |
| `error-diagnostics-error-trace` | You are an error tracking and observability expert specializing in implementing comprehensive error monitoring solutions. Set up error tracking sys... |
| `firecrawl-scraper` | Deep web scraping, screenshots, PDF parsing, and website crawling using Firecrawl API |
| `food-database-query` | Food Database Query |
| `fp-data-transforms` | Everyday data transformations using functional patterns - arrays, objects, grouping, aggregation, and null-safe access |
| `fp-option-ref` | Quick reference for Option type. Use when user needs to handle nullable values, optional data, or wants to avoid null checks. |
| `fp-pipe-ref` | Quick reference for pipe and flow. Use when user needs to chain functions, compose operations, or build data pipelines in fp-ts. |
| `fp-react` | Practical patterns for using fp-ts with React - hooks, state, forms, data fetching. Works with React 18/19, Next.js 14/15. |
| `fp-ts-react` | Practical patterns for using fp-ts with React - hooks, state, forms, data fetching. Use when building React apps with functional programming patter... |
| `frontend-dev-guidelines` | Opinionated frontend development standards for modern React + TypeScript applications. Covers Suspense-first data fetching, lazy loading, feature-b... |
| `frontend-mobile-development-component-scaffold` | You are a React component architecture expert specializing in scaffolding production-ready, accessible, and performant components. Generate complet... |
| `frontend-ui-dark-ts` | Build dark-themed React applications using Tailwind CSS with custom theming, glassmorphism effects, and Framer Motion animations. Use when creating... |
| `gdpr-data-handling` | Implement GDPR-compliant data handling with consent management, data subject rights, and privacy by design. Use when building systems that process ... |
| `grafana-dashboards` | Create and manage production Grafana dashboards for real-time visualization of system and application metrics. Use when building monitoring dashboa... |
| `graphql` | GraphQL gives clients exactly the data they need - no more, no less. One endpoint, typed schema, introspection. But the flexibility that makes it p... |
| `grpc-golang` | Build production-ready gRPC services in Go with mTLS, streaming, and observability. Use when designing Protobuf contracts with Buf or implementing ... |
| `hugging-face-tool-builder` | Use this skill when the user wants to build tool/scripts or achieve a task where using data from the Hugging Face API would help. This is especiall... |
| `incident-responder` | Expert SRE incident responder specializing in rapid problem resolution, modern observability, and comprehensive incident management. |
| `instagram` | Integracao completa com Instagram via Graph API. Publicacao, analytics, comentarios, DMs, hashtags, agendamento, templates e gestao de contas Busin... |
| `ios-developer` | Develop native iOS applications with Swift/SwiftUI. Masters iOS 18, SwiftUI, UIKit integration, Core Data, networking, and App Store optimization. |
| `junta-leiloeiros` | Coleta e consulta dados de leiloeiros oficiais de todas as 27 Juntas Comerciais do Brasil. Scraper multi-UF, banco SQLite, API FastAPI e exportacao... |
| `keyword-extractor` | Extracts up to 50 highly relevant SEO keywords from text. Use when user wants to generate or extract keywords for given text. |
| `kpi-dashboard-design` | Design effective KPI dashboards with metrics selection, visualization best practices, and real-time monitoring patterns. Use when building business... |
| `latex-paper-conversion` | This skill should be used when the user asks to convert an academic paper in LaTeX from one format (e.g., Springer, IPOL) to another format (e.g., ... |
| `legacy-modernizer` | Refactor legacy codebases, migrate outdated frameworks, and implement gradual modernization. Handles technical debt, dependency updates, and backwa... |
| `leiloeiro-mercado` | Analise de mercado imobiliario para leiloes. Liquidez, desagio tipico, ROI, estrategias de saida (flip/reforma/renda), Selic 2025 e benchmark CDI/FII. |
| `marketing-ideas` | Provide proven marketing strategies and growth ideas for SaaS and software products, prioritized using a marketing feasibility scoring system. |
| `mobile-design` | Mobile-first design and engineering doctrine for iOS and Android apps. Covers touch interaction, performance, platform conventions, offline behavio... |
| `mobile-games` | Mobile game development principles. Touch input, battery, performance, app stores. |
| `multi-platform-apps-multi-platform` | Build and deploy the same feature consistently across web, mobile, and desktop platforms using API-first architecture and parallel implementation s... |
| `native-data-fetching` | Use when implementing or debugging ANY network request, API call, or data fetching. Covers fetch API, React Query, SWR, error handling, caching, of... |
| `neon-postgres` | Expert patterns for Neon serverless Postgres, branching, connection pooling, and Prisma/Drizzle integration Use when: neon database, serverless pos... |
| `networkx` | Comprehensive toolkit for creating, analyzing, and visualizing complex networks and graphs in Python. Use when working with network/graph data stru... |
| `nextjs-app-router-patterns` | Master Next.js 14+ App Router with Server Components, streaming, parallel routes, and advanced data fetching. Use when building Next.js application... |
| `nextjs-best-practices` | Next.js App Router principles. Server Components, data fetching, routing patterns. |
| `nft-standards` | Implement NFT standards (ERC-721, ERC-1155) with proper metadata handling, minting strategies, and marketplace integration. Use when creating NFT c... |
| `observability-monitoring-monitor-setup` | You are a monitoring and observability expert specializing in implementing comprehensive monitoring solutions. Set up metrics collection, distribut... |
| `observability-monitoring-slo-implement` | You are an SLO (Service Level Objective) expert specializing in implementing reliability standards and error budget-based practices. Design SLO fra... |
| `odoo-automated-tests` | Write and run Odoo automated tests using TransactionCase, HttpCase, and browser tour tests. Covers test data setup, mocking, and CI integration. |
| `odoo-docker-deployment` | Production-ready Docker and docker-compose setup for Odoo with PostgreSQL, persistent volumes, environment-based configuration, and Nginx reverse p... |
| `odoo-edi-connector` | Guide for implementing EDI (Electronic Data Interchange) with Odoo: X12, EDIFACT document mapping, partner onboarding, and automated order processing. |
| `odoo-project-timesheet` | Expert guide for Odoo Project and Timesheets: task stages, billable time tracking, timesheet approval, budget alerts, and invoicing from timesheets. |
| `odoo-upgrade-advisor` | Step-by-step Odoo version upgrade advisor: pre-upgrade checklist, community vs enterprise upgrade path, OCA module compatibility, and post-upgrade ... |
| `on-call-handoff-patterns` | Master on-call shift handoffs with context transfer, escalation procedures, and documentation. Use when transitioning on-call responsibilities, doc... |
| `performance-engineer` | Expert performance engineer specializing in modern observability, |
| `performance-optimizer` | Identifies and fixes performance bottlenecks in code, databases, and APIs. Measures before and after to prove improvements. |
| `php-pro` | Write idiomatic PHP code with generators, iterators, SPL data structures, and modern OOP features. Use PROACTIVELY for high-performance PHP applica... |
| `plotly` | Interactive visualization library. Use when you need hover info, zoom, pan, or web-embeddable charts. Best for dashboards, exploratory analysis, an... |
| `posix-shell-pro` | Expert in strict POSIX sh scripting for maximum portability across Unix-like systems. Specializes in shell scripts that run on any POSIX-compliant ... |
| `postgres-best-practices` | Postgres performance optimization and best practices from Supabase. Use this skill when writing, reviewing, or optimizing Postgres queries, schema ... |
| `postgresql` | Design a PostgreSQL-specific schema. Covers best-practices, data types, indexing, constraints, performance patterns, and advanced features |
| `privacy-by-design` | Use when building apps that collect user data. Ensures privacy protections are built in from the start—data minimization, consent, encryption. |
| `product-inventor` | Product Inventor e Design Alchemist de nivel maximo — combina Product Thinking, Design Systems, UI Engineering, Psicologia Cognitiva, Storytelling ... |
| `professional-proofreader` | Use when a user asks to "proofread", "review and correct", "fix grammar", "improve readability while keeping my voice", and to proofread a document... |
| `python-development-python-scaffold` | You are a Python project architecture expert specializing in scaffolding production-ready Python applications. Generate complete project structures... |
| `python-packaging` | Create distributable Python packages with proper project structure, setup.py/pyproject.toml, and publishing to PyPI. Use when packaging Python libr... |
| `python-testing-patterns` | Implement comprehensive testing strategies with pytest, fixtures, mocking, and test-driven development. Use when writing Python tests, setting up t... |
| `rehabilitation-analyzer` | 分析康复训练数据、识别康复模式、评估康复进展，并提供个性化康复建议 |
| `reverse-engineer` | Expert reverse engineer specializing in binary analysis, disassembly, decompilation, and software analysis. Masters IDA Pro, Ghidra, radare2, x64db... |
| `robius-event-action` | CRITICAL: Use for Robius event and action patterns. Triggers on: custom action, MatchEvent, post_action, cx.widget_action, handle_actions, DefaultN... |
| `robius-state-management` | CRITICAL: Use for Robius state management patterns. Triggers on: AppState, persistence, theme switch, 状态管理, Scope::with_data, save state, load stat... |
| `sankhya-dashboard-html-jsp-custom-best-pratices` | This skill should be used when the user asks for patterns, best practices, creation, or fixing of Sankhya dashboards using HTML, JSP, Java, and SQL. |
| `scala-pro` | Master enterprise-grade Scala development with functional programming, distributed systems, and big data processing. Expert in Apache Pekko, Akka, ... |
| `schema-markup` | Design, validate, and optimize schema.org structured data for eligibility, correctness, and measurable SEO impact. |
| `screen-reader-testing` | Test web applications with screen readers including VoiceOver, NVDA, and JAWS. Use when validating screen reader compatibility, debugging accessibi... |
| `seaborn` | Statistical visualization with pandas integration. Use for quick exploration of distributions, relationships, and categorical comparisons with attr... |
| `segment-cdp` | Expert patterns for Segment Customer Data Platform including Analytics.js, server-side tracking, tracking plans with Protocols, identity resolution... |
| `senior-architect` | Comprehensive software architecture skill for designing scalable, maintainable systems using ReactJS, NextJS, NodeJS, Express, React Native, Swift,... |
| `senior-frontend` | Frontend development skill for React, Next.js, TypeScript, and Tailwind CSS applications. Use when building React components, optimizing Next.js pe... |
| `senior-fullstack` | Comprehensive fullstack development skill for building complete web applications with React, Next.js, Node.js, GraphQL, and PostgreSQL. Includes pr... |
| `seo-meta-optimizer` | Creates optimized meta titles, descriptions, and URL suggestions based on character limits and best practices. Generates compelling, keyword-rich m... |
| `service-mesh-observability` | Implement comprehensive observability for service meshes including distributed tracing, metrics, and visualization. Use when setting up mesh monito... |
| `shellcheck-configuration` | Master ShellCheck static analysis configuration and usage for shell script quality. Use when setting up linting infrastructure, fixing code issues,... |
| `skill-creator` | This skill should be used when the user asks to create a new skill, build a skill, make a custom skill, develop a CLI skill, or wants to extend the... |
| `skin-health-analyzer` | Analyze skin health data, identify skin problem patterns, assess skin health status. Supports correlation analysis with nutrition, chronic diseases... |
| `slo-implementation` | Define and implement Service Level Indicators (SLIs) and Service Level Objectives (SLOs) with error budgets and alerting. Use when establishing rel... |
| `sql-optimization-patterns` | Master SQL query optimization, indexing strategies, and EXPLAIN analysis to dramatically improve database performance and eliminate slow queries. U... |
| `sympy` | Use this skill when working with symbolic mathematics in Python. This skill should be used for symbolic computation tasks including solving equatio... |
| `tavily-web` | Web search, content extraction, crawling, and research capabilities using Tavily API |
| `threejs-postprocessing` | Three.js post-processing - EffectComposer, bloom, DOF, screen effects. Use when adding visual effects, color grading, blur, glow, or creating custo... |
| `threejs-skills` | Create 3D scenes, interactive experiences, and visual effects using Three.js. Use when user requests 3D graphics, WebGL experiences, 3D visualizati... |
| `tool-use-guardian` | FREE — Intelligent tool-call reliability wrapper. Monitors, retries, fixes, and learns from tool failures. Auto-recovers from truncated JSON, timeo... |
| `ui-ux-designer` | Create interface designs, wireframes, and design systems. Masters user research, accessibility standards, and modern design tools. |
| `unity-ecs-patterns` | Master Unity ECS (Entity Component System) with DOTS, Jobs, and Burst for high-performance game development. Use when building data-oriented games,... |
| `youtube-summarizer` | Extract transcripts from YouTube videos and generate comprehensive, detailed summaries using intelligent analysis frameworks |

[↑ Voltar ao Índice](#indice-de-categorias)

---

## 📁 Development Tools

| ID | Descrição |
| :--- | :--- |
| `address-github-comments` | Use when you need to address review or issue comments on an open GitHub Pull Request using the gh CLI. |
| `bug-hunter` | Systematically finds and fixes bugs using proven debugging techniques. Traces from symptoms to root cause, implements fixes, and prevents regression. |
| `c4-architecture-c4-architecture` | Generate comprehensive C4 architecture documentation for an existing repository/codebase using a bottom-up analysis approach. |
| `c4-code` | Expert C4 Code-level documentation specialist. Analyzes code directories to create comprehensive C4 code-level documentation including function sig... |
| `c4-component` | Expert C4 Component-level documentation specialist. Synthesizes C4 Code-level documentation into Component-level architecture, defining component b... |
| `code-documentation-code-explain` | You are a code education expert specializing in explaining complex code through clear narratives, visual diagrams, and step-by-step breakdowns. Tra... |
| `code-refactoring-context-restore` | Use when working with code refactoring context restore |
| `commit` | ALWAYS use this skill when committing code changes — never commit directly without it. Creates commits following Sentry conventions with proper con... |
| `conductor-revert` | Git-aware undo by logical work unit (track, phase, or task) |
| `cpp-pro` | Write idiomatic C++ code with modern features, RAII, smart pointers, and STL algorithms. Handles templates, move semantics, and performance optimiz... |
| `create-branch` | Create a git branch following Sentry naming conventions. Use when asked to "create a branch", "new branch", "start a branch", "make a branch", "swi... |
| `ddd-tactical-patterns` | Apply DDD tactical patterns in code using entities, value objects, aggregates, repositories, and domain events with explicit invariants. |
| `debugger` | Debugging specialist for errors, test failures, and unexpected behavior. Use proactively when encountering any issues. |
| `debugging-strategies` | Master systematic debugging techniques, profiling tools, and root cause analysis to efficiently track down bugs across any codebase or technology s... |
| `debugging-toolkit-smart-debug` | Use when working with debugging toolkit smart debug |
| `distributed-tracing` | Implement distributed tracing with Jaeger and Tempo to track requests across microservices and identify performance bottlenecks. Use when debugging... |
| `elixir-pro` | Write idiomatic Elixir code with OTP patterns, supervision trees, and Phoenix LiveView. Masters concurrency, fault tolerance, and distributed systems. |
| `error-detective` | Search logs and codebases for error patterns, stack traces, and anomalies. Correlates errors across systems and identifies root causes. |
| `error-diagnostics-smart-debug` | Use when working with error diagnostics smart debug |
| `fp-errors` | Stop throwing everywhere - handle errors as values using Either and TaskEither for cleaner, more predictable code |
| `fp-ts-errors` | Handle errors as values using fp-ts Either and TaskEither for cleaner, more predictable TypeScript code. Use when implementing error handling patte... |
| `framework-migration-code-migrate` | You are a code migration expert specializing in transitioning codebases between frameworks, languages, versions, and platforms. Generate comprehens... |
| `gh-review-requests` | Fetch unread GitHub notifications for open PRs where review is requested from a specified team or opened by a team member. Use when asked to "find ... |
| `git-pushing` | Stage, commit, and push git changes with conventional commit messages. Use when user wants to commit and push changes, mentions pushing to remote, ... |
| `github-issue-creator` | Convert raw notes, error logs, voice dictation, or screenshots into crisp GitHub-flavored markdown issue reports. Use when the user pastes bug info... |
| `i18n-localization` | Internationalization and localization patterns. Detecting hardcoded strings, managing translations, locale files, RTL support. |
| `issues` | Interact with GitHub issues - create, list, and view issues. |
| `javascript-testing-patterns` | Implement comprehensive testing strategies using Jest, Vitest, and Testing Library for unit tests, integration tests, and end-to-end testing with m... |
| `kotlin-coroutines-expert` | Expert patterns for Kotlin Coroutines and Flow, covering structured concurrency, error handling, and testing. |
| `linkedin-cli` | Use when automating LinkedIn via CLI: fetch profiles, search people/companies, send messages, manage connections, create posts, and Sales Navigator. |
| `makepad-layout` | CRITICAL: Use for Makepad layout system. Triggers on: makepad layout, makepad width, makepad height, makepad flex, makepad padding, makepad margin,... |
| `market-sizing-analysis` | This skill should be used when the user asks to \\"calculate TAM\\", "determine SAM", "estimate SOM", "size the market", "calculate market opportun... |
| `oss-hunter` | Automatically hunt for high-impact OSS contribution opportunities in trending repositories. |
| `pr-writer` | ALWAYS use this skill when creating or updating pull requests — never create or edit a PR directly without it. Follows Sentry conventions for PR ti... |
| `ruby-pro` | Write idiomatic Ruby code with metaprogramming, Rails patterns, and performance optimization. Specializes in Ruby on Rails, gem development, and te... |
| `systematic-debugging` | Use when encountering any bug, test failure, or unexpected behavior, before proposing fixes |
| `test-driven-development` | Use when implementing any feature or bugfix, before writing implementation code |
| `typescript-pro` | Master TypeScript with advanced types, generics, and strict type safety. Handles complex type systems, decorators, and enterprise-grade patterns. |
| `unsplash-integration` | Integration skill for searching and fetching high-quality, free-to-use professional photography from Unsplash. |
| `using-git-worktrees` | Use when starting feature work that needs isolation from current workspace or before executing implementation plans - creates isolated git worktree... |
| `vexor-cli` | Semantic file discovery via `vexor`. Use whenever locating where something is implemented/loaded/defined in a medium or large repo, or when the fil... |
| `weightloss-analyzer` | 分析减肥数据、计算代谢率、追踪能量缺口、管理减肥阶段 |
| `wiki-changelog` | Analyzes git commit history and generates structured changelogs categorized by change type. Use when the user asks about recent changes, wants a ch... |
| `wiki-page-writer` | Generates rich technical documentation pages with dark-mode Mermaid diagrams, source code citations, and first-principles depth. Use when writing d... |
| `wiki-qa` | Answers questions about a code repository using source file analysis. Use when the user asks a question about how something works, wants to underst... |
| `wiki-researcher` | Conducts multi-turn iterative deep research on specific topics within a codebase with zero tolerance for shallow analysis. Use when the user wants ... |

[↑ Voltar ao Índice](#indice-de-categorias)

---

## 📁 General & Productivity

| ID | Descrição |
| :--- | :--- |
| `2d-games` | 2D game development principles. Sprites, tilemaps, physics, camera. |
| `3d-games` | 3D game development principles. Rendering, shaders, physics, cameras. |
| `architect-review` | Master software architect specializing in modern architecture |
| `architecture-decision-records` | Write and maintain Architecture Decision Records (ADRs) following best practices for technical decision documentation. Use when documenting signifi... |
| `arm-cortex-expert` | Senior embedded software engineer specializing in firmware and driver development for ARM Cortex-M microcontrollers (Teensy, STM32, nRF52, SAMD). |
| `bevy-ecs-expert` | Master Bevy's Entity Component System (ECS) in Rust, covering Systems, Queries, Resources, and parallel scheduling. |
| `c4-container` | Expert C4 Container-level documentation specialist. |
| `comprehensive-review-full-review` | Use when working with comprehensive review full review |
| `conductor-manage` | Manage track lifecycle: archive, restore, delete, rename, and cleanup |
| `conductor-new-track` | Create a new track with specification and phased implementation plan |
| `context-guardian` | Guardiao de contexto que preserva dados criticos antes da compactacao automatica. Snapshots, verificacao de integridade e zero perda de informacao. |
| `create-issue-gate` | Use when starting a new implementation task and an issue must be created with strict acceptance criteria gating before execution. |
| `ddd-context-mapping` | Map relationships between bounded contexts and define integration contracts using DDD context mapping patterns. |
| `django-access-review` | django-access-review |
| `evolution` | CRITICAL: Use for makepad-skills self-evolution and contribution. Triggers on: evolve, evolution, contribute, contribution, self-improve, self-impr... |
| `executing-plans` | Use when you have a written implementation plan to execute in a separate session with review checkpoints |
| `explain-like-socrates` | Explains concepts using Socratic-style dialogue. Use when the user asks to explain, teach or help understand a concept like socrates. |
| `family-health-analyzer` | 分析家族病史、评估遗传风险、识别家庭健康模式、提供个性化预防建议 |
| `favicon` | Generate favicons from a source image |
| `file-organizer` | Intelligently organizes files and folders by understanding context, finding duplicates, and suggesting better organizational structures. Use when u... |
| `fitness-analyzer` | 分析运动数据、识别运动模式、评估健身进展，并提供个性化训练建议。支持与慢性病数据的关联分析。 |
| `form-cro` | Optimize any form that is NOT signup or account registration — including lead capture, contact, demo request, application, survey, quote, and check... |
| `goal-analyzer` | 分析健康目标数据、识别目标模式、评估目标进度,并提供个性化目标管理建议。支持与营养、运动、睡眠等健康数据的关联分析。 |
| `haskell-pro` | Expert Haskell engineer specializing in advanced type systems, pure |
| `hr-pro` | Professional, ethical HR partner for hiring, onboarding/offboarding, PTO and leave, performance, compliant policies, and employee relations. |
| `incident-response-incident-response` | Use when working with incident response incident response |
| `leiloeiro-ia` | Especialista em leiloes judiciais e extrajudiciais de imoveis. Analise juridica, pericial e de mercado integrada. Orquestra os 5 modulos especializ... |
| `makepad-event-action` | CRITICAL: Use for Makepad event and action handling. Triggers on: makepad event, makepad action, Event enum, ActionTrait, handle_event, MouseDown, ... |
| `makepad-shaders` | CRITICAL: Use for Makepad shader system. Triggers on: makepad shader, makepad draw_bg, Sdf2d, makepad pixel, makepad glsl, makepad sdf, draw_quad, ... |
| `mental-health-analyzer` | 分析心理健康数据、识别心理模式、评估心理健康状况、提供个性化心理健康建议。支持与睡眠、运动、营养等其他健康数据的关联分析。 |
| `nutrition-analyzer` | 分析营养数据、识别营养模式、评估营养状况，并提供个性化营养建议。支持与运动、睡眠、慢性病数据的关联分析。 |
| `occupational-health-analyzer` | 分析职业健康数据、识别工作相关健康风险、评估职业健康状况、提供个性化职业健康建议。支持与睡眠、运动、心理健康等其他健康数据的关联分析。 |
| `odoo-orm-expert` | Master Odoo ORM patterns: search, browse, create, write, domain filters, computed fields, and performance-safe query techniques. |
| `onboarding-cro` | When the user wants to optimize post-signup onboarding, user activation, first-run experience, or time-to-value. Also use when the user mentions "o... |
| `oral-health-analyzer` | 分析口腔健康数据、识别口腔问题模式、评估口腔健康状况、提供个性化口腔健康建议。支持与营养、慢性病、用药等其他健康数据的关联分析。 |
| `paywall-upgrade-cro` | When the user wants to create or optimize in-app paywalls, upgrade screens, upsell modals, or feature gates. Also use when the user mentions "paywa... |
| `pc-games` | PC and console game development principles. Engine selection, platform features, optimization strategies. |
| `performance-profiling` | Performance profiling principles. Measurement, analysis, and optimization techniques. |
| `postmortem-writing` | Write effective blameless postmortems with root cause analysis, timelines, and action items. Use when conducting incident reviews, writing postmort... |
| `protocol-reverse-engineering` | Master network protocol reverse engineering including packet analysis, protocol dissection, and custom protocol documentation. Use when analyzing n... |
| `pypict-skill` | Pairwise test generation |
| `referral-program` | When the user wants to create, optimize, or analyze a referral program, affiliate program, or word-of-mouth strategy. Also use when the user mentio... |
| `semgrep-rule-variant-creator` | Creates language variants of existing Semgrep rules. Use when porting a Semgrep rule to specified target languages. Takes an existing rule and targ... |
| `sexual-health-analyzer` | Sexual Health Analyzer |
| `sharp-edges` | sharp-edges |
| `shodan-reconnaissance` | This skill should be used when the user asks to "search for exposed devices on the internet," "perform Shodan reconnaissance," "find vulnerable ser... |
| `skill-rails-upgrade` | Analyze Rails apps and provide upgrade assessments |
| `sleep-analyzer` | 分析睡眠数据、识别睡眠模式、评估睡眠质量，并提供个性化睡眠改善建议。支持与其他健康数据的关联分析。 |
| `speckit-updater` | SpecKit Safe Update |
| `speed` | Launch RSVP speed reader for text |
| `tcm-constitution-analyzer` | 分析中医体质数据、识别体质类型、评估体质特征,并提供个性化养生建议。支持与营养、运动、睡眠等健康数据的关联分析。 |
| `travel-health-analyzer` | 分析旅行健康数据、评估目的地健康风险、提供疫苗接种建议、生成多语言紧急医疗信息卡片。支持WHO/CDC数据集成的专业级旅行健康风险评估。 |
| `upgrading-expo` | Upgrade Expo SDK versions |
| `wireshark-analysis` | This skill should be used when the user asks to "analyze network traffic with Wireshark", "capture packets for troubleshooting", "filter PCAP files... |
| `x-article-publisher-skill` | Publish articles to X/Twitter |
| `yann-lecun-tecnico` | Sub-skill técnica de Yann LeCun. Cobre CNNs, LeNet, backpropagation, JEPA (I-JEPA, V-JEPA, MC-JEPA), AMI (Advanced Machinery of Intelligence), Self... |

[↑ Voltar ao Índice](#indice-de-categorias)

---

## 📁 Legal & Finance

| ID | Descrição |
| :--- | :--- |
| `advogado-criminal` | Advogado criminalista especializado em Maria da Penha, violencia domestica, feminicidio, direito penal brasileiro, medidas protetivas, inquerito po... |
| `advogado-especialista` | Advogado especialista em todas as areas do Direito brasileiro: familia, criminal, trabalhista, tributario, consumidor, imobiliario, empresarial, ci... |
| `alpha-vantage` | Access real-time and historical stock market data, forex rates, cryptocurrency prices, commodities, economic indicators, and 50+ technical indicato... |
| `azure-keyvault-keys-rust` | Azure Key Vault Keys SDK for Rust. Use for creating, managing, and using cryptographic keys. Triggers: "keyvault keys rust", "KeyClient rust", "cre... |
| `azure-keyvault-keys-ts` | Manage cryptographic keys using Azure Key Vault Keys SDK for JavaScript (@azure/keyvault-keys). Use when creating, encrypting/decrypting, signing, ... |
| `defi-protocol-templates` | Implement DeFi protocols with production-ready templates for staking, AMMs, governance, and lending systems. Use when building decentralized financ... |
| `employment-contract-templates` | Create employment contracts, offer letters, and HR policy documents following legal best practices. Use when drafting employment agreements, creati... |
| `godot-4-migration` | Specialized guide for migrating Godot 3.x projects to Godot 4 (GDScript 2.0), covering syntax changes, Tweens, and exports. |
| `inventory-demand-planning` | Codified expertise for demand forecasting, safety stock optimisation, replenishment planning, and promotional lift estimation at multi-location ret... |
| `legal-advisor` | Draft privacy policies, terms of service, disclaimers, and legal notices. Creates GDPR-compliant texts, cookie policies, and data processing agreem... |
| `leiloeiro-juridico` | Analise juridica de leiloes: nulidades, bem de familia, alienacao fiduciaria, CPC arts 829-903, Lei 9514/97, onus reais, embargos e jurisprudencia. |
| `leiloeiro-risco` | Analise de risco em leiloes de imoveis. Score 36 pontos, riscos juridicos/financeiros/operacionais, stress test 4 cenarios e ROI ponderado por risco. |
| `lex` | Centralized 'Truth Engine' for cross-jurisdictional legal context (US, EU, CA) and contract scaffolding. |
| `lightning-channel-factories` | Technical reference on Lightning Network channel factories, multi-party channels, LSP architectures, and Bitcoin Layer 2 scaling without soft forks... |
| `lightning-factory-explainer` | Explain Bitcoin Lightning channel factories and the SuperScalar protocol — scalable Lightning onboarding using shared UTXOs, Decker-Wattenhofer tre... |
| `lint-and-validate` | Automatic quality control, linting, and static analysis procedures. Use after every code modification to ensure syntax correctness and project stan... |
| `makepad-dsl` | CRITICAL: Use for Makepad DSL syntax and inheritance. Triggers on: makepad dsl, live_design, makepad inheritance, makepad prototype, "<Widget>", "F... |
| `mermaid-expert` | Create Mermaid diagrams for flowcharts, sequences, ERDs, and architectures. Masters syntax for all diagram types and styling. |
| `n8n-code-javascript` | Write JavaScript code in n8n Code nodes. Use when writing JavaScript in n8n, using $input/$json/$node syntax, making HTTP requests with $helpers, w... |
| `n8n-code-python` | Write Python code in n8n Code nodes. Use when writing Python in n8n, using _input/_json/_node syntax, working with standard library, or need to und... |
| `n8n-expression-syntax` | Validate n8n expression syntax and fix common errors. Use when writing n8n expressions, using {{}} syntax, accessing $json/$node variables, trouble... |
| `odoo-accounting-setup` | Expert guide for configuring Odoo Accounting: chart of accounts, journals, fiscal positions, taxes, payment terms, and bank reconciliation. |
| `odoo-inventory-optimizer` | Expert guide for Odoo Inventory: stock valuation (FIFO/AVCO), reordering rules, putaway strategies, routes, and multi-warehouse configuration. |
| `odoo-xml-views-builder` | Expert at building Odoo XML views: Form, List, Kanban, Search, Calendar, and Graph. Generates correct XML for Odoo 14-17 with proper visibility syn... |
| `powershell-windows` | PowerShell Windows patterns. Critical pitfalls, operator syntax, error handling. |
| `startup-business-analyst-business-case` | Generate comprehensive investor-ready business case document with market, solution, financials, and strategy |

[↑ Voltar ao Índice](#indice-de-categorias)

---

## 📁 Management & Strategy

| ID | Descrição |
| :--- | :--- |
| `carrier-relationship-management` | Codified expertise for managing carrier portfolios, negotiating freight rates, tracking carrier performance, allocating freight, and maintaining st... |
| `code-refactoring-tech-debt` | You are a technical debt expert specializing in identifying, quantifying, and prioritizing technical debt in software projects. Analyze the codebas... |
| `codebase-cleanup-tech-debt` | You are a technical debt expert specializing in identifying, quantifying, and prioritizing technical debt in software projects. Analyze the codebas... |
| `concise-planning` | Use when a user asks for a plan for a coding task, to generate a clear, actionable, and atomic checklist. |
| `conductor-status` | Display project status, active tracks, and next actions |
| `conductor-validator` | Validates Conductor project artifacts for completeness, consistency, and correctness. Use after setup, when diagnosing issues, or before implementa... |
| `context-management-context-restore` | Use when working with context management context restore |
| `context-management-context-save` | Use when working with context management context save |
| `diary` | Unified Diary System: A context-preserving automated logger for multi-project development. |
| `filesystem-context` | Use for file-based context management, dynamic context discovery, and reducing context window bloat. Offload context to files for just-in-time load... |
| `framework-migration-deps-upgrade` | You are a dependency management expert specializing in safe, incremental upgrades of project dependencies. Plan and execute dependency updates with... |
| `interview-coach` | Full job search coaching system — JD decoding, resume, storybank, mock interviews, transcript analysis, comp negotiation. 23 commands, persistent s... |
| `logistics-exception-management` | Codified expertise for handling freight exceptions, shipment delays, damages, losses, and carrier disputes. Informed by logistics professionals wit... |
| `mtls-configuration` | Configure mutual TLS (mTLS) for zero-trust service-to-service communication. Use when implementing zero-trust networking, certificate management, o... |
| `new-rails-project` | Create a new Rails project |
| `plan-writing` | Structured task planning with clear breakdowns, dependencies, and verification criteria. Use when implementing features, refactoring, or any multi-... |
| `quality-nonconformance` | Codified expertise for quality control, non-conformance investigation, root cause analysis, corrective action, and supplier quality management in r... |
| `readme` | When the user wants to create or update a README.md file for a project. Also use when the user says 'write readme,' 'create readme,' 'document this... |
| `remotion` | Generate walkthrough videos from Stitch projects using Remotion with smooth transitions, zooming, and text overlays |
| `risk-manager` | Monitor portfolio risk, R-multiples, and position limits. Creates hedging strategies, calculates expectancy, and implements stop-losses. |
| `skill-router` | Use when the user is unsure which skill to use or where to start. Interviews the user with targeted questions and recommends the best skill(s) from... |
| `sred-project-organizer` | Take a list of projects and their related documentation, and organize them into the SRED format for submission. |
| `startup-business-analyst-market-opportunity` | Generate comprehensive market opportunity analysis with TAM/SAM/SOM calculations |
| `startup-metrics-framework` | This skill should be used when the user asks about \\"key startup metrics", "SaaS metrics", "CAC and LTV", "unit economics", "burn multiple", "rule... |
| `templates` | Project scaffolding templates for new applications. Use when creating new projects from scratch. Contains 12 templates for various tech stacks. |
| `track-management` | Use this skill when creating, managing, or working with Conductor tracks - the logical work units for features, bugs, and refactors. Applies to spe... |
| `typescript-expert` | TypeScript and JavaScript expert with deep knowledge of type-level programming, performance optimization, monorepo management, migration strategies... |
| `uncle-bob-craft` | Use when performing code review, writing or refactoring code, or discussing architecture; complements clean-code and does not replace project linte... |

[↑ Voltar ao Índice](#indice-de-categorias)

---

## 📁 Marketing & Product

| ID | Descrição |
| :--- | :--- |
| `c4-context` | Expert C4 Context-level documentation specialist. Creates high-level system context diagrams, documents personas, user journeys, system features, a... |
| `competitor-alternatives` | When the user wants to create competitor comparison or alternative pages for SEO and sales enablement. Also use when the user mentions 'alternative... |
| `content-creator` | Create SEO-optimized marketing content with consistent brand voice. Includes brand voice analyzer, SEO optimizer, content frameworks, and social me... |
| `context-driven-development` | Use this skill when working with Conductor's context-driven development methodology, managing project context artifacts, or understanding the relat... |
| `copywriting` | Write rigorous, conversion-focused marketing copy for landing pages and emails. Enforces brief confirmation and strict no-fabrication rules. |
| `email-sequence` | When the user wants to create or optimize an email sequence, drip campaign, automated email flow, or lifecycle email program. Also use when the use... |
| `email-systems` | Email has the highest ROI of any marketing channel. $36 for every $1 spent. Yet most startups treat it as an afterthought - bulk blasts, no persona... |
| `golang-pro` | Master Go 1.21+ with modern patterns, advanced concurrency, performance optimization, and production-ready microservices. |
| `julia-pro` | Master Julia 1.10+ with modern features, performance optimization, multiple dispatch, and production-ready practices. |
| `launch-strategy` | When the user wants to plan a product launch, feature announcement, or release strategy. Also use when the user mentions 'launch,' 'Product Hunt,' ... |
| `page-cro` | Analyze and optimize individual pages for conversion performance. |
| `paid-ads` | When the user wants help with paid advertising campaigns on Google Ads, Meta (Facebook/Instagram), LinkedIn, Twitter/X, or other ad platforms. Also... |
| `popup-cro` | Create and optimize popups, modals, overlays, slide-ins, and banners to increase conversions without harming user experience or brand trust. |
| `product-manager-toolkit` | Comprehensive toolkit for product managers including RICE prioritization, customer interview analysis, PRD templates, discovery frameworks, and go-... |
| `production-scheduling` | Codified expertise for production scheduling, job sequencing, line balancing, changeover optimisation, and bottleneck resolution in discrete and ba... |
| `qiskit` | IBM quantum computing framework. Use when targeting IBM Quantum hardware, working with Qiskit Runtime for production workloads, or needing IBM opti... |
| `sales-automator` | Draft cold emails, follow-ups, and proposal templates. Creates pricing pages, case studies, and sales scripts. Use PROACTIVELY for sales outreach o... |
| `screenshots` | Generate marketing screenshots of your app using Playwright. Use when the user wants to create screenshots for Product Hunt, social media, landing ... |
| `seo-cannibalization-detector` | Analyzes multiple provided pages to identify keyword overlap and potential cannibalization issues. Suggests differentiation strategies. Use PROACTI... |
| `seo-content-planner` | Creates comprehensive content outlines and topic clusters for SEO. Plans content calendars and identifies topic gaps. Use PROACTIVELY for content s... |
| `seo-content-refresher` | Identifies outdated elements in provided content and suggests updates to maintain freshness. Finds statistics, dates, and examples that need updati... |
| `seo-content-writer` | Writes SEO-optimized content based on provided keywords and topic briefs. Creates engaging, comprehensive content following best practices. Use PRO... |
| `seo-keyword-strategist` | Analyzes keyword usage in provided content, calculates density, suggests semantic variations and LSI keywords based on the topic. Prevents over-opt... |
| `seo-snippet-hunter` | Formats content to be eligible for featured snippets and SERP features. Creates snippet-optimized content blocks based on best practices. Use PROAC... |
| `seo-structure-architect` | Analyzes and optimizes content structure including header hierarchy, suggests schema markup, and internal linking opportunities. Creates search-fri... |
| `signup-flow-cro` | When the user wants to optimize signup, registration, account creation, or trial activation flows. Also use when the user mentions "signup conversi... |
| `social-content` | When the user wants help creating, scheduling, or optimizing social media content for LinkedIn, Twitter/X, Instagram, TikTok, Facebook, or other pl... |
| `systems-programming-rust-project` | You are a Rust project architecture expert specializing in scaffolding production-ready Rust applications. Generate complete project structures wit... |
| `tutorial-engineer` | Creates step-by-step tutorials and educational content from code. Transforms complex concepts into progressive learning experiences with hands-on e... |

[↑ Voltar ao Índice](#indice-de-categorias)

---

## 📁 Security & Privacy

| ID | Descrição |
| :--- | :--- |
| `007` | Security audit, hardening, threat modeling (STRIDE/PASTA), Red/Blue Team, OWASP checks, code review, incident response, and infrastructure security... |
| `accessibility-compliance-accessibility-audit` | You are an accessibility expert specializing in WCAG compliance, inclusive design, and assistive technology compatibility. Conduct audits, identify... |
| `active-directory-attacks` | This skill should be used when the user asks to "attack Active Directory", "exploit AD", "Kerberoasting", "DCSync", "pass-the-hash", "BloodHound en... |
| `agentic-actions-auditor` | Audits GitHub Actions workflows for security vulnerabilities in AI agent integrations including Claude Code Action, Gemini CLI, OpenAI Codex, and G... |
| `ai-engineering-toolkit` | 6 production-ready AI engineering workflows: prompt evaluation (8-dimension scoring), context budget planning, RAG pipeline design, agent security ... |
| `ai-md` | Convert human-written CLAUDE.md into AI-native structured-label format. Battle-tested across 4 models. Same rules, fewer tokens, higher compliance. |
| `analytics-tracking` | Design, audit, and improve analytics tracking systems that produce reliable, decision-ready data. |
| `analyze-project` | Forensic root cause analyzer for Antigravity sessions. Classifies scope deltas, rework patterns, root causes, hotspots, and auto-improves prompts/h... |
| `anti-reversing-techniques` | Understand anti-reversing, obfuscation, and protection techniques encountered during software analysis. Use when analyzing protected binaries, bypa... |
| `antigravity-workflows` | Orchestrate multiple Antigravity skills through guided workflows for SaaS MVP delivery, security audits, AI agent builds, and browser QA. |
| `api-endpoint-builder` | Builds production-ready REST API endpoints with validation, error handling, authentication, and documentation. Follows best practices for security ... |
| `api-fuzzing-bug-bounty` | This skill should be used when the user asks to "test API security", "fuzz APIs", "find IDOR vulnerabilities", "test REST API", "test GraphQL", "AP... |
| `api-security-best-practices` | Implement secure API design patterns including authentication, authorization, input validation, rate limiting, and protection against common API vu... |
| `api-security-testing` | API security testing workflow for REST and GraphQL APIs covering authentication, authorization, rate limiting, input validation, and security best ... |
| `apify-influencer-discovery` | Find and evaluate influencers for brand partnerships, verify authenticity, and track collaboration performance across Instagram, Facebook, YouTube,... |
| `attack-tree-construction` | Build comprehensive attack trees to visualize threat paths. Use when mapping attack scenarios, identifying defense gaps, or communicating security ... |
| `audit-context-building` | Enables ultra-granular, line-by-line code analysis to build deep architectural context before vulnerability or bug finding. |
| `audit-skills` | Expert security auditor for AI Skills and Bundles. Performs non-intrusive static analysis to identify malicious patterns, data leaks, system stabil... |
| `auth-implementation-patterns` | Master authentication and authorization patterns including JWT, OAuth2, session management, and RBAC to build secure, scalable access control syste... |
| `avoid-ai-writing` | Audit and rewrite content to remove 21 categories of AI writing patterns with a 43-entry replacement table |
| `aws-compliance-checker` | Automated compliance checking against CIS, PCI-DSS, HIPAA, and SOC 2 benchmarks |
| `aws-iam-best-practices` | IAM policy review, hardening, and least privilege implementation |
| `aws-penetration-testing` | This skill should be used when the user asks to "pentest AWS", "test AWS security", "enumerate IAM", "exploit cloud infrastructure", "AWS privilege... |
| `aws-secrets-rotation` | Automate AWS secrets rotation for RDS, API keys, and credentials |
| `aws-security-audit` | Comprehensive AWS security posture assessment using AWS CLI and security best practices |
| `azure-communication-common-java` | Azure Communication Services common utilities for Java. Use when working with CommunicationTokenCredential, user identifiers, token refresh, or sha... |
| `azure-cosmos-db-py` | Build Azure Cosmos DB NoSQL services with Python/FastAPI following production-grade patterns. Use when implementing database client setup with dual... |
| `azure-identity-dotnet` | Azure Identity SDK for .NET. Authentication library for Azure SDK clients using Microsoft Entra ID. Use for DefaultAzureCredential, managed identit... |
| `azure-identity-java` | Azure Identity Java SDK for authentication with Azure services. Use when implementing DefaultAzureCredential, managed identity, service principal, ... |
| `azure-identity-py` | Azure Identity SDK for Python authentication. Use for DefaultAzureCredential, managed identity, service principals, and token caching. |
| `azure-identity-rust` | Azure Identity SDK for Rust authentication. Use for DeveloperToolsCredential, ManagedIdentityCredential, ClientSecretCredential, and token-based au... |
| `azure-identity-ts` | Authenticate to Azure services using Azure Identity SDK for JavaScript (@azure/identity). Use when configuring authentication with DefaultAzureCred... |
| `azure-keyvault-py` | Azure Key Vault SDK for Python. Use for secrets, keys, and certificates management with secure storage. |
| `azure-keyvault-secrets-rust` | Azure Key Vault Secrets SDK for Rust. Use for storing and retrieving secrets, passwords, and API keys. Triggers: "keyvault secrets rust", "SecretCl... |
| `azure-keyvault-secrets-ts` | Manage secrets using Azure Key Vault Secrets SDK for JavaScript (@azure/keyvault-secrets). Use when storing and retrieving application secrets or c... |
| `azure-security-keyvault-keys-dotnet` | Azure Key Vault Keys SDK for .NET. Client library for managing cryptographic keys in Azure Key Vault and Managed HSM. Use for key creation, rotatio... |
| `azure-security-keyvault-keys-java` | Azure Key Vault Keys Java SDK for cryptographic key management. Use when creating, managing, or using RSA/EC keys, performing encrypt/decrypt/sign/... |
| `azure-security-keyvault-secrets-java` | Azure Key Vault Secrets Java SDK for secret management. Use when storing, retrieving, or managing passwords, API keys, connection strings, or other... |
| `backend-security-coder` | Expert in secure backend coding practices specializing in input validation, authentication, and API security. Use PROACTIVELY for backend security ... |
| `broken-authentication` | This skill should be used when the user asks to "test for broken authentication vulnerabilities", "assess session management security", "perform cr... |
| `burp-suite-testing` | This skill should be used when the user asks to "intercept HTTP traffic", "modify web requests", "use Burp Suite for testing", "perform web vulnera... |
| `burpsuite-project-parser` | Searches and explores Burp Suite project files (.burp) from the command line. Use when searching response headers or bodies with regex patterns, ex... |
| `cc-skill-security-review` | Use this skill when adding authentication, handling user input, working with secrets, creating API endpoints, or implementing payment/sensitive fea... |
| `claude-settings-audit` | Analyze a repository to generate recommended Claude Code settings.json permissions. Use when setting up a new project, auditing existing settings, ... |
| `clerk-auth` | Expert patterns for Clerk auth implementation, middleware, organizations, webhooks, and user sync Use when: adding authentication, clerk auth, user... |
| `cloud-penetration-testing` | This skill should be used when the user asks to "perform cloud penetration testing", "assess Azure or AWS or GCP security", "enumerate cloud resour... |
| `code-review-checklist` | Comprehensive checklist for conducting thorough code reviews covering functionality, security, performance, and maintainability |
| `codebase-audit-pre-push` | Deep audit before GitHub push: removes junk files, dead code, security holes, and optimization issues. Checks every file line-by-line for productio... |
| `codebase-cleanup-deps-audit` | You are a dependency security expert specializing in vulnerability scanning, license compliance, and supply chain security. Analyze project depende... |
| `comfyui-gateway` | REST API gateway for ComfyUI servers. Workflow management, job queuing, webhooks, caching, auth, rate limiting, and image delivery (URL + base64). |
| `constant-time-analysis` | Detects timing side-channel vulnerabilities in cryptographic code. Use when implementing or reviewing crypto code, encountering division on secrets... |
| `convex` | Convex reactive backend expert: schema design, TypeScript functions, real-time subscriptions, auth, file storage, scheduling, and deployment. |
| `cred-omega` | CISO operacional enterprise para gestao total de credenciais e segredos. Descobre, classifica, protege e governa TODAS as API keys, tokens, secrets... |
| `crypto-bd-agent` | Autonomous crypto business development patterns — multi-chain token discovery, 100-point scoring with wallet forensics, x402 micropayments, ERC-800... |
| `customs-trade-compliance` | Codified expertise for customs documentation, tariff classification, duty optimisation, restricted party screening, and regulatory compliance acros... |
| `debug-buttercup` | Debugs the Buttercup CRS (Cyber Reasoning System) running on Kubernetes. Use when diagnosing pod crashes, restart loops, Redis failures, resource p... |
| `dependency-management-deps-audit` | You are a dependency security expert specializing in vulnerability scanning, license compliance, and supply chain security. Analyze project depende... |
| `deployment-pipeline-design` | Design multi-stage CI/CD pipelines with approval gates, security checks, and deployment orchestration. Use when architecting deployment workflows, ... |
| `differential-review` | Performs security-focused differential review of code changes (PRs, commits, diffs). Adapts analysis depth to codebase size, uses git history for c... |
| `django-perf-review` | Django performance code review. Use when asked to "review Django performance", "find N+1 queries", "optimize Django", "check queryset performance",... |
| `doc-coauthoring` | Guide users through a structured workflow for co-authoring documentation. Use when user wants to write documentation, proposals, technical specs, d... |
| `docker-expert` | Docker containerization expert with deep knowledge of multi-stage builds, image optimization, container security, Docker Compose orchestration, and... |
| `dotnet-backend` | Build ASP.NET Core 8+ backend services with EF Core, auth, background jobs, and production API patterns. |
| `ethical-hacking-methodology` | This skill should be used when the user asks to "learn ethical hacking", "understand penetration testing lifecycle", "perform reconnaissance", "con... |
| `fastapi-router-py` | Create FastAPI routers with CRUD operations, authentication dependencies, and proper response models. Use when building REST API endpoints, creatin... |
| `fda-food-safety-auditor` | Expert AI auditor for FDA Food Safety (FSMA), HACCP, and PCQI compliance. Reviews food facility records and preventive controls. |
| `fda-medtech-compliance-auditor` | Expert AI auditor for Medical Device (SaMD) compliance, IEC 62304, and 21 CFR Part 820. Reviews DHFs, technical files, and software validation. |
| `ffuf-web-fuzzing` | Expert guidance for ffuf web fuzzing during penetration testing, including authenticated fuzzing with raw requests, auto-calibration, and result an... |
| `file-path-traversal` | This skill should be used when the user asks to "test for directory traversal", "exploit path traversal vulnerabilities", "read arbitrary files thr... |
| `find-bugs` | Find bugs, security vulnerabilities, and code quality issues in local branch changes. Use when asked to review changes, find bugs, security review,... |
| `firebase` | Firebase gives you a complete backend in minutes - auth, database, storage, functions, hosting. But the ease of setup hides real complexity. Securi... |
| `firmware-analyst` | Expert firmware analyst specializing in embedded systems, IoT security, and hardware reverse engineering. |
| `fix-review` | Verify fix commits address audit findings without new bugs |
| `fixing-accessibility` | Audit and fix HTML accessibility issues including ARIA labels, keyboard navigation, focus management, color contrast, and form errors. Use when add... |
| `fixing-metadata` | Audit and fix HTML metadata including page titles, meta descriptions, canonical URLs, Open Graph tags, Twitter cards, favicons, JSON-LD structured ... |
| `fixing-motion-performance` | Audit and fix animation performance issues including layout thrashing, compositor properties, scroll-linked motion, and blur effects. Use when anim... |
| `frontend-mobile-security-xss-scan` | You are a frontend security specialist focusing on Cross-Site Scripting (XSS) vulnerability detection and prevention. Analyze React, Vue, Angular, ... |
| `frontend-security-coder` | Expert in secure frontend coding practices specializing in XSS prevention, output sanitization, and client-side security patterns. |
| `gha-security-review` | GitHub Actions security review for workflow exploitation vulnerabilities. Use when asked to "review GitHub Actions", "audit workflows", "check CI s... |
| `graphql-architect` | Master modern GraphQL with federation, performance optimization, and enterprise security. Build scalable schemas, implement advanced caching, and d... |
| `growth-engine` | Motor de crescimento para produtos digitais -- growth hacking, SEO, ASO, viral loops, email marketing, CRM, referral programs e aquisicao organica.... |
| `html-injection-testing` | This skill should be used when the user asks to "test for HTML injection", "inject HTML into web pages", "perform HTML injection attacks", "deface ... |
| `hubspot-integration` | Expert patterns for HubSpot CRM integration including OAuth authentication, CRM objects, associations, batch operations, webhooks, and custom objec... |
| `hugging-face-jobs` | This skill should be used when users want to run any workload on Hugging Face Jobs infrastructure. Covers UV scripts, Docker-based jobs, hardware s... |
| `hugging-face-paper-publisher` | Publish and manage research papers on Hugging Face Hub. Supports creating paper pages, linking papers to models/datasets, claiming authorship, and ... |
| `idor-testing` | This skill should be used when the user asks to "test for insecure direct object references," "find IDOR vulnerabilities," "exploit broken access c... |
| `k8s-manifest-generator` | Create production-ready Kubernetes manifests for Deployments, Services, ConfigMaps, and Secrets following best practices and security standards. Us... |
| `k8s-security-policies` | Implement Kubernetes security policies including NetworkPolicy, PodSecurityPolicy, and RBAC for production-grade security. Use when securing Kubern... |
| `laravel-expert` | Senior Laravel Engineer role for production-grade, maintainable, and idiomatic Laravel solutions. Focuses on clean architecture, security, performa... |
| `laravel-security-audit` | Security auditor for Laravel applications. Analyzes code for vulnerabilities, misconfigurations, and insecure practices using OWASP standards and L... |
| `leiloeiro-edital` | Analise e auditoria de editais de leilao judicial e extrajudicial. Riscos ocultos, clausulas perigosas, debitos, ocupante e classificacao da oportu... |
| `linkerd-patterns` | Implement Linkerd service mesh patterns for lightweight, security-focused service mesh deployments. Use when setting up Linkerd, configuring traffi... |
| `linux-privilege-escalation` | This skill should be used when the user asks to "escalate privileges on Linux", "find privesc vectors on Linux systems", "exploit sudo misconfigura... |
| `local-legal-seo-audit` | Audit and improve local SEO for law firms, attorneys, forensic experts and legal/professional services sites with local presence, focusing on GBP, ... |
| `loki-mode` | Multi-agent autonomous startup system for Claude Code. Triggers on "Loki Mode". Orchestrates 100+ specialized agents across engineering, QA, DevOps... |
| `m365-agents-dotnet` | Microsoft 365 Agents SDK for .NET. Build multichannel agents for Teams/M365/Copilot Studio with ASP.NET Core hosting, AgentApplication routing, and... |
| `m365-agents-py` | Microsoft 365 Agents SDK for Python. Build multichannel agents for Teams/M365/Copilot Studio with aiohttp hosting, AgentApplication routing, stream... |
| `malware-analyst` | Expert malware analyst specializing in defensive malware research, threat intelligence, and incident response. Masters sandbox analysis, behavioral... |
| `memory-forensics` | Master memory forensics techniques including memory acquisition, process analysis, and artifact extraction using Volatility and related tools. Use ... |
| `metasploit-framework` | This skill should be used when the user asks to "use Metasploit for penetration testing", "exploit vulnerabilities with msfconsole", "create payloa... |
| `micro-saas-launcher` | Expert in launching small, focused SaaS products fast - the indie hacker approach to building profitable software. Covers idea validation, MVP deve... |
| `microsoft-azure-webjobs-extensions-authentication-events-dotnet` | Microsoft Entra Authentication Events SDK for .NET. Azure Functions triggers for custom authentication extensions. |
| `mobile-security-coder` | Expert in secure mobile coding practices specializing in input validation, WebView security, and mobile-specific security patterns. |
| `nestjs-expert` | Nest.js framework expert specializing in module architecture, dependency injection, middleware, guards, interceptors, testing with Jest/Supertest, ... |
| `network-engineer` | Expert network engineer specializing in modern cloud networking, security architectures, and performance optimization. |
| `nextjs-supabase-auth` | Expert integration of Supabase Auth with Next.js App Router Use when: supabase auth next, authentication next.js, login supabase, auth middleware, ... |
| `nodejs-backend-patterns` | Build production-ready Node.js backend services with Express/Fastify, implementing middleware patterns, error handling, authentication, database in... |
| `nodejs-best-practices` | Node.js development principles and decision-making. Framework selection, async patterns, security, and architecture. Teaches thinking, not copying. |
| `notebooklm` | Use this skill to query your Google NotebookLM notebooks directly from Claude Code for source-grounded, citation-backed answers from Gemini. Browse... |
| `odoo-l10n-compliance` | Country-specific Odoo localization: tax configuration, e-invoicing (CFDI, FatturaPA, SAF-T), fiscal reporting, and country chart of accounts setup. |
| `odoo-rpc-api` | Expert on Odoo's external JSON-RPC and XML-RPC APIs. Covers authentication, model calls, record CRUD, and real-world integration examples in Python... |
| `odoo-security-rules` | Expert in Odoo access control: ir.model.access.csv, record rules (ir.rule), groups, and multi-company security patterns. |
| `openapi-spec-generation` | Generate and maintain OpenAPI 3.1 specifications from code, design-first specs, and validation patterns. Use when creating API documentation, gener... |
| `payment-integration` | Integrate Stripe, PayPal, and payment processors. Handles checkout flows, subscriptions, webhooks, and PCI compliance. Use PROACTIVELY when impleme... |
| `pci-compliance` | Implement PCI DSS compliance requirements for secure handling of payment card data and payment systems. Use when securing payment processing, achie... |
| `pentest-checklist` | This skill should be used when the user asks to "plan a penetration test", "create a security assessment checklist", "prepare for penetration testi... |
| `pentest-commands` | This skill should be used when the user asks to "run pentest commands", "scan with nmap", "use metasploit exploits", "crack passwords with hydra or... |
| `plaid-fintech` | Expert patterns for Plaid API integration including Link token flows, transactions sync, identity verification, Auth for ACH, balance checks, webho... |
| `privilege-escalation-methods` | This skill should be used when the user asks to "escalate privileges", "get root access", "become administrator", "privesc techniques", "abuse sudo... |
| `production-code-audit` | Autonomously deep-scan entire codebase line-by-line, understand architecture and patterns, then systematically transform it to production-grade, co... |
| `python-fastapi-development` | Python FastAPI backend development with async patterns, SQLAlchemy, Pydantic, authentication, and production API patterns. |
| `red-team-tactics` | Red team tactics principles based on MITRE ATT&CK. Attack phases, detection evasion, reporting. |
| `red-team-tools` | This skill should be used when the user asks to "follow red team methodology", "perform bug bounty hunting", "automate reconnaissance", "hunt for X... |
| `returns-reverse-logistics` | Codified expertise for returns authorisation, receipt and inspection, disposition decisions, refund processing, fraud detection, and warranty claim... |
| `saas-mvp-launcher` | Use when planning or building a SaaS MVP from scratch. Provides a structured roadmap covering tech stack, architecture, auth, payments, and launch ... |
| `sast-configuration` | Configure Static Application Security Testing (SAST) tools for automated vulnerability detection in application code. Use when setting up security ... |
| `scanning-tools` | This skill should be used when the user asks to "perform vulnerability scanning", "scan networks for open ports", "assess web application security"... |
| `secrets-management` | Implement secure secrets management for CI/CD pipelines using Vault, AWS Secrets Manager, or native platform solutions. Use when handling sensitive... |
| `security-audit` | Comprehensive security auditing workflow covering web application testing, API security, penetration testing, vulnerability scanning, and security ... |
| `security-auditor` | Expert security auditor specializing in DevSecOps, comprehensive cybersecurity, and compliance frameworks. |
| `security-bluebook-builder` | Create or refine a concise, normative security policy ("Blue Book") for sensitive applications. Use when users need a threat model, data classifica... |
| `security-compliance-compliance-check` | You are a compliance expert specializing in regulatory requirements for software systems including GDPR, HIPAA, SOC2, PCI-DSS, and other industry s... |
| `security-requirement-extraction` | Derive security requirements from threat models and business context. Use when translating threats into actionable requirements, creating security ... |
| `security-scanning-security-dependencies` | You are a security expert specializing in dependency vulnerability analysis, SBOM generation, and supply chain security. Scan project dependencies ... |
| `security-scanning-security-hardening` | Coordinate multi-layer security scanning and hardening across application, infrastructure, and compliance controls. |
| `security-scanning-security-sast` | Static Application Security Testing (SAST) for code vulnerability analysis across multiple languages and frameworks |
| `semgrep-rule-creator` | Creates custom Semgrep rules for detecting security vulnerabilities, bug patterns, and code patterns. Use when writing Semgrep rules or building cu... |
| `seo-audit` | Diagnose and audit SEO issues affecting crawlability, indexation, rankings, and organic performance. |
| `seo-authority-builder` | Analyzes content for E-E-A-T signals and suggests improvements to build authority and trust. Identifies missing credibility elements. Use PROACTIVE... |
| `seo-content-auditor` | Analyzes provided content for quality, E-E-A-T signals, and SEO best practices. Scores content and provides improvement recommendations based on es... |
| `seo-forensic-incident-response` | Investigate sudden drops in organic traffic or rankings and run a structured forensic SEO incident response with triage, root-cause analysis and re... |
| `service-mesh-expert` | Expert service mesh architect specializing in Istio, Linkerd, and cloud-native networking patterns. Masters traffic management, security policies, ... |
| `skill-scanner` | Scan agent skills for security issues. Use when asked to "scan a skill", "audit a skill", "review skill security", "check skill for injection", "va... |
| `skill-sentinel` | Auditoria e evolucao do ecossistema de skills. Qualidade de codigo, seguranca, custos, gaps, duplicacoes, dependencias e relatorios de saude. |
| `slack-bot-builder` | Build Slack apps using the Bolt framework across Python, JavaScript, and Java. Covers Block Kit for rich UIs, interactive components, slash command... |
| `smtp-penetration-testing` | This skill should be used when the user asks to "perform SMTP penetration testing", "enumerate email users", "test for open mail relays", "grab SMT... |
| `solidity-security` | Master smart contract security best practices to prevent common vulnerabilities and implement secure Solidity patterns. Use when writing smart cont... |
| `spec-to-code-compliance` | Verifies code implements exactly what documentation specifies for blockchain audits. Use when comparing code against whitepapers, finding gaps betw... |
| `sql-injection-testing` | This skill should be used when the user asks to "test for SQL injection vulnerabilities", "perform SQLi attacks", "bypass authentication using SQL ... |
| `sqlmap-database-pentesting` | This skill should be used when the user asks to "automate SQL injection testing," "enumerate database structure," "extract database credentials usi... |
| `ssh-penetration-testing` | This skill should be used when the user asks to "pentest SSH services", "enumerate SSH configurations", "brute force SSH credentials", "exploit SSH... |
| `stride-analysis-patterns` | Apply STRIDE methodology to systematically identify threats. Use when analyzing system security, conducting threat modeling sessions, or creating s... |
| `supply-chain-risk-auditor` | Identifies dependencies at heightened risk of exploitation or takeover. Use when assessing supply chain attack surface, evaluating dependency healt... |
| `telegram-mini-app` | Expert in building Telegram Mini Apps (TWA) - web apps that run inside Telegram with native-like experience. Covers the TON ecosystem, Telegram Web... |
| `threat-mitigation-mapping` | Map identified threats to appropriate security controls and mitigations. Use when prioritizing security investments, creating remediation plans, or... |
| `threat-modeling-expert` | Expert in threat modeling methodologies, security architecture review, and risk assessment. Masters STRIDE, PASTA, attack trees, and security requi... |
| `top-web-vulnerabilities` | This skill should be used when the user asks to "identify web application vulnerabilities", "explain common security flaws", "understand vulnerabil... |
| `ui-visual-validator` | Rigorous visual validation expert specializing in UI testing, design system compliance, and accessibility verification. |
| `using-neon` | Guides and best practices for working with Neon Serverless Postgres. Covers getting started, local development with Neon, choosing a connection met... |
| `variant-analysis` | Find similar vulnerabilities and bugs across codebases using pattern-based analysis. Use when hunting bug variants, building CodeQL/Semgrep queries... |
| `varlock` | Secure environment variable management with Varlock. Use when handling secrets, API keys, credentials, or any sensitive configuration. Ensures secr... |
| `varlock-claude-skill` | Secure environment variable management ensuring secrets are never exposed in Claude sessions, terminals, logs, or git commits |
| `vibe-code-auditor` | Audit rapidly generated or AI-produced code for structural flaws, fragility, and production risks. |
| `vulnerability-scanner` | Advanced vulnerability analysis principles. OWASP 2025, Supply Chain Security, attack surface mapping, risk prioritization. |
| `wcag-audit-patterns` | Conduct WCAG 2.2 accessibility audits with automated testing, manual verification, and remediation guidance. Use when auditing websites for accessi... |
| `web-design-guidelines` | Review UI code for Web Interface Guidelines compliance. Use when asked to "review my UI", "check accessibility", "audit design", "review UX", or "c... |
| `web-security-testing` | Web application security testing workflow for OWASP Top 10 vulnerabilities including injection, XSS, authentication flaws, and access control issues. |
| `windows-privilege-escalation` | This skill should be used when the user asks to "escalate privileges on Windows," "find Windows privesc vectors," "enumerate Windows for privilege ... |
| `wordpress` | Complete WordPress development workflow covering theme development, plugin creation, WooCommerce integration, performance optimization, and securit... |
| `wordpress-penetration-testing` | This skill should be used when the user asks to "pentest WordPress sites", "scan WordPress for vulnerabilities", "enumerate WordPress users, themes... |
| `wordpress-plugin-development` | WordPress plugin development workflow covering plugin architecture, hooks, admin interfaces, REST API, and security best practices. |
| `xss-html-injection` | This skill should be used when the user asks to "test for XSS vulnerabilities", "perform cross-site scripting attacks", "identify HTML injection fl... |
| `zeroize-audit` | Detects missing zeroization of sensitive data in source code and identifies zeroization removed by compiler optimizations, with assembly-level anal... |

[↑ Voltar ao Índice](#indice-de-categorias)

---

## 📁 Web & Frontend

| ID | Descrição |
| :--- | :--- |
| `00-andruia-consultant` | Arquitecto de Soluciones Principal y Consultor Tecnológico de Andru.ia. Diagnostica y traza la hoja de ruta óptima para proyectos de IA en español. |
| `20-andruia-niche-intelligence` | Estratega de Inteligencia de Dominio de Andru.ia. Analiza el nicho específico de un proyecto para inyectar conocimientos, regulaciones y estándares... |
| `3d-web-experience` | Expert in building 3D experiences for the web - Three.js, React Three Fiber, Spline, WebGL, and interactive 3D scenes. Covers product configurators... |
| `ab-test-setup` | Structured guide for setting up A/B tests with mandatory gates for hypothesis, metrics, and execution readiness. |
| `algorithmic-art` | Creating algorithmic art using p5.js with seeded randomness and interactive parameter exploration. Use this when users request creating art using c... |
| `android-jetpack-compose-expert` | Expert guidance for building modern Android UIs with Jetpack Compose, covering state management, navigation, performance, and Material Design 3. |
| `angular` | Modern Angular (v20+) expert with deep knowledge of Signals, Standalone Components, Zoneless applications, SSR/Hydration, and reactive patterns. |
| `angular-best-practices` | Angular performance optimization and best practices guide. Use when writing, reviewing, or refactoring Angular code for optimal performance, bundle... |
| `angular-migration` | Migrate from AngularJS to Angular using hybrid mode, incremental component rewriting, and dependency injection updates. Use when upgrading AngularJ... |
| `angular-state-management` | Master modern Angular state management with Signals, NgRx, and RxJS. Use when setting up global state, managing component stores, choosing between ... |
| `animejs-animation` | Advanced JavaScript animation library skill for creating complex, high-performance web animations. |
| `antigravity-design-expert` | Core UI/UX engineering skill for building highly interactive, spatial, weightless, and glassmorphism-based web interfaces using GSAP and 3D CSS. |
| `ask-questions-if-underspecified` | Clarify requirements before implementing. Use when serious doubts arise. |
| `avalonia-layout-zafiro` | Guidelines for modern Avalonia UI layout using Zafiro.Avalonia, emphasizing shared styles, generic components, and avoiding XAML redundancy. |
| `avalonia-zafiro-development` | Mandatory skills, conventions, and behavioral rules for Avalonia UI development using the Zafiro toolkit. |
| `bash-defensive-patterns` | Master defensive Bash programming techniques for production-grade scripts. Use when writing robust shell scripts, CI/CD pipelines, or system utilit... |
| `bats-testing-patterns` | Master Bash Automated Testing System (Bats) for comprehensive shell script testing. Use when writing tests for shell scripts, CI/CD pipelines, or r... |
| `bazel-build-optimization` | Optimize Bazel builds for large-scale monorepos. Use when configuring Bazel, implementing remote execution, or optimizing build performance for ent... |
| `blockchain-developer` | Build production-ready Web3 applications, smart contracts, and decentralized systems. Implements DeFi protocols, NFT platforms, DAOs, and enterpris... |
| `blog-writing-guide` | Write, review, and improve blog posts for the Sentry engineering blog following Sentry's specific writing standards, voice, and quality bar. Use th... |
| `brainstorming` | Use before creative or constructive work (features, architecture, behavior). Transforms vague ideas into validated designs through disciplined reas... |
| `brand-guidelines` | Write copy following Sentry brand guidelines. Use when writing UI text, error messages, empty states, onboarding flows, 404 pages, documentation, m... |
| `browser-extension-builder` | Expert in building browser extensions that solve real problems - Chrome, Firefox, and cross-browser extensions. Covers extension architecture, mani... |
| `build` | build |
| `building-native-ui` | Complete guide for building beautiful apps with Expo Router. Covers fundamentals, styling, components, navigation, animations, patterns, and native... |
| `busybox-on-windows` | How to use a Win32 build of BusyBox to run many of the standard UNIX command line tools on Windows. |
| `canvas-design` | Create beautiful visual art in .png and .pdf documents using design philosophy. You should use this skill when the user asks to create a poster, pi... |
| `cc-skill-frontend-patterns` | Frontend development patterns for React, Next.js, state management, performance optimization, and UI best practices. |
| `cc-skill-project-guidelines-example` | Project Guidelines Skill (Example) |
| `chat-widget` | Build a real-time support chat system with a floating widget for users and an admin dashboard for support staff. Use when the user wants live chat,... |
| `chrome-extension-developer` | Expert in building Chrome Extensions using Manifest V3. Covers background scripts, service workers, content scripts, and cross-context communication. |
| `code-refactoring-refactor-clean` | You are a code refactoring expert specializing in clean code principles, SOLID design patterns, and modern software engineering best practices. Ana... |
| `codebase-cleanup-refactor-clean` | You are a code refactoring expert specializing in clean code principles, SOLID design patterns, and modern software engineering best practices. Ana... |
| `design-md` | Analyze Stitch projects and synthesize a semantic design system into DESIGN.md files |
| `design-spells` | Curated micro-interactions and design details that add "magic" and personality to websites and apps. |
| `docs-architect` | Creates comprehensive technical documentation from existing codebases. Analyzes architecture, design patterns, and implementation details to produc... |
| `electron-development` | Master Electron desktop app development with secure IPC, contextIsolation, preload scripts, multi-process architecture, electron-builder packaging,... |
| `emergency-card` | 生成紧急情况下快速访问的医疗信息摘要卡片。当用户需要旅行、就诊准备、紧急情况或询问"紧急信息"、"医疗卡片"、"急救信息"时使用此技能。提取关键信息（过敏、用药、急症、植入物），支持多格式输出（JSON、文本、二维码），用于急救或快速就医。 |
| `environment-setup-guide` | Guide developers through setting up development environments with proper tools, dependencies, and configurations |
| `error-handling-patterns` | Master error handling patterns across languages including exceptions, Result types, error propagation, and graceful degradation to build resilient ... |
| `expo-dev-client` | Build and distribute Expo development clients locally or via TestFlight |
| `expo-tailwind-setup` | Set up Tailwind CSS v4 in Expo with react-native-css and NativeWind v5 for universal styling |
| `expo-ui-jetpack-compose` | expo-ui-jetpack-compose |
| `expo-ui-swift-ui` | expo-ui-swift-ui |
| `fp-either-ref` | Quick reference for Either type. Use when user needs error handling, validation, or operations that can fail with typed errors. |
| `fp-refactor` | Comprehensive guide for refactoring imperative TypeScript code to fp-ts functional patterns |
| `fp-types-ref` | Quick reference for fp-ts types. Use when user asks which type to use, needs Option/Either/Task decision help, or wants fp-ts imports. |
| `frontend-design` | Create distinctive, production-grade frontend interfaces with intentional aesthetics, high craft, and non-generic visual identity. Use when buildin... |
| `frontend-developer` | Build React components, implement responsive layouts, and handle client-side state management. Masters React 19, Next.js 15, and modern frontend ar... |
| `frontend-slides` | Create stunning, animation-rich HTML presentations from scratch or by converting PowerPoint files. |
| `game-audio` | Game audio principles. Sound design, music integration, adaptive audio systems. |
| `game-design` | Game design principles. GDD structure, balancing, player psychology, progression. |
| `godot-gdscript-patterns` | Master Godot 4 GDScript patterns including signals, scenes, state machines, and optimization. Use when building Godot games, implementing game syst... |
| `health-trend-analyzer` | 分析一段时间内健康数据的趋势和模式。关联药物、症状、生命体征、化验结果和其他健康指标的变化。识别令人担忧的趋势、改善情况，并提供数据驱动的洞察。当用户询问健康趋势、模式、随时间的变化或"我的健康状况有什么变化？"时使用。支持多维度分析（体重/BMI、症状、药物依从性、化验结果、情绪睡眠），相关... |
| `hig-components-content` | Apple Human Interface Guidelines for content display components. |
| `hig-components-dialogs` | Apple HIG guidance for presentation components including alerts, action sheets, popovers, sheets, and digit entry views. |
| `hig-components-layout` | Apple Human Interface Guidelines for layout and navigation components. |
| `hig-components-menus` | Apple HIG guidance for menu and button components including menus, context menus, dock menus, edit menus, the menu bar, toolbars, action buttons, p... |
| `hig-components-search` | Apple HIG guidance for navigation-related components including search fields, page controls, and path controls. |
| `hig-components-status` | Apple HIG guidance for status and progress UI components including progress indicators, status bars, and activity rings. |
| `hig-components-system` | Apple HIG guidance for system experience components: widgets, live activities, notifications, complications, home screen quick actions, top shelf, ... |
| `hig-foundations` | Apple Human Interface Guidelines design foundations. |
| `hig-inputs` | Apple HIG guidance for input methods and interaction patterns: gestures, Apple Pencil, keyboards, game controllers, pointers, Digital Crown, eye tr... |
| `hig-patterns` | Apple Human Interface Guidelines interaction and UX patterns. |
| `hig-platforms` | Apple Human Interface Guidelines for platform-specific design. |
| `hig-project-context` | Create or update a shared Apple design context document that other HIG skills use to tailor guidance. |
| `incident-runbook-templates` | Create structured incident response runbooks with step-by-step procedures, escalation paths, and recovery actions. Use when building runbooks, resp... |
| `interactive-portfolio` | Expert in building portfolios that actually land jobs and clients - not just showing work, but creating memorable experiences. Covers developer por... |
| `kaizen` | Guide for continuous improvement, error proofing, and standardization. Use this skill when the user wants to improve code quality, refactor, or dis... |
| `leiloeiro-avaliacao` | Avaliacao pericial de imoveis em leilao. Valor de mercado, liquidacao forcada, ABNT NBR 14653, metodos comparativo/renda/custo, CUB e margem de seg... |
| `machine-learning-ops-ml-pipeline` | Design and implement a complete ML pipeline for: $ARGUMENTS |
| `makepad-animation` | CRITICAL: Use for Makepad animation system. Triggers on: makepad animation, makepad animator, makepad hover, makepad state, makepad transition, "fr... |
| `makepad-basics` | CRITICAL: Use for Makepad getting started and app structure. Triggers on: makepad, makepad getting started, makepad tutorial, live_design!, app_mai... |
| `makepad-skills` | Makepad UI development skills for Rust apps: setup, patterns, shaders, packaging, and troubleshooting. |
| `makepad-widgets` | CRITICAL: Use for Makepad widgets and UI components. Triggers on: makepad widget, makepad View, makepad Button, makepad Label, makepad Image, makep... |
| `matematico-tao` | Matemático ultra-avançado inspirado em Terence Tao. Análise rigorosa de código e arquitetura com teoria matemática profunda: teoria da informação, ... |
| `microservices-patterns` | Design microservices architectures with service boundaries, event-driven communication, and resilience patterns. Use when building distributed syst... |
| `monorepo-architect` | Expert in monorepo architecture, build systems, and dependency management at scale. Masters Nx, Turborepo, Bazel, and Lerna for efficient multi-pro... |
| `monorepo-management` | Master monorepo management with Turborepo, Nx, and pnpm workspaces to build efficient, scalable multi-package repositories with optimized builds an... |
| `n8n-validation-expert` | Interpret validation errors and guide fixing them. Use when encountering validation errors, validation warnings, false positives, operator structur... |
| `nx-workspace-patterns` | Configure and optimize Nx monorepo workspaces. Use when setting up Nx, configuring project boundaries, optimizing build caching, or implementing af... |
| `obsidian-clipper-template-creator` | Guide for creating templates for the Obsidian Web Clipper. Use when you want to create a new clipping template, understand available variables, or ... |
| `odoo-hr-payroll-setup` | Expert guide for Odoo HR and Payroll: salary structures, payslip rules, leave policies, employee contracts, and payroll journal entries. |
| `odoo-qweb-templates` | Expert in Odoo QWeb templating for PDF reports, email templates, and website pages. Covers t-if, t-foreach, t-field, and report actions. |
| `odoo-sales-crm-expert` | Expert guide for Odoo Sales and CRM: pipeline stages, quotation templates, pricelists, sales teams, lead scoring, and forecasting. |
| `personal-tool-builder` | Expert in building custom tools that solve your own problems first. The best products often start as personal tools - scratch your own itch, build ... |
| `planning-with-files` | Implements Manus-style file-based planning for complex tasks. Creates task_plan.md, findings.md, and progress.md. Use when starting complex multi-s... |
| `pricing-strategy` | Design pricing, packaging, and monetization strategies based on value, customer willingness to pay, and growth objectives. |
| `radix-ui-design-system` | Build accessible design systems with Radix UI primitives. Headless component customization, theming strategies, and compound component patterns for... |
| `react-best-practices` | React and Next.js performance optimization guidelines from Vercel Engineering. This skill should be used when writing, reviewing, or refactoring Re... |
| `react-modernization` | Upgrade React applications to latest versions, migrate from class components to hooks, and adopt concurrent features. Use when modernizing React co... |
| `react-patterns` | Modern React patterns and principles. Hooks, composition, performance, TypeScript best practices. |
| `remotion-best-practices` | Best practices for Remotion - Video creation in React |
| `requesting-code-review` | Use when completing tasks, implementing major features, or before merging to verify work meets requirements |
| `risk-metrics-calculation` | Calculate portfolio risk metrics including VaR, CVaR, Sharpe, Sortino, and drawdown analysis. Use when measuring portfolio risk, implementing risk ... |
| `scroll-experience` | Expert in building immersive scroll-driven experiences - parallax storytelling, scroll animations, interactive narratives, and cinematic web experi... |
| `search-specialist` | Expert web researcher using advanced search techniques and |
| `shadcn` | Manages shadcn/ui components and projects, providing context, documentation, and usage patterns for building modern design systems. |
| `software-architecture` | Guide for quality focused software architecture. This skill should be used when users want to write code, design architecture, analyze code, in any... |
| `swiftui-expert-skill` | Write, review, or improve SwiftUI code following best practices for state management, view composition, performance, and iOS 26+ Liquid Glass adopt... |
| `team-composition-analysis` | This skill should be used when the user asks to \\"plan team structure", "determine hiring needs", "design org chart", "calculate compensation", "p... |
| `test-fixing` | Run tests and systematically fix all failing tests using smart error grouping. Use when user asks to fix failing tests, mentions test failures, run... |
| `theme-factory` | Toolkit for styling artifacts with a theme. These artifacts can be slides, docs, reportings, HTML landing pages, etc. There are 10 pre-set themes w... |
| `threejs-animation` | Three.js animation - keyframe animation, skeletal animation, morph targets, animation mixing. Use when animating objects, playing GLTF animations, ... |
| `threejs-fundamentals` | Three.js scene setup, cameras, renderer, Object3D hierarchy, coordinate systems. Use when setting up 3D scenes, creating cameras, configuring rende... |
| `threejs-geometry` | Three.js geometry creation - built-in shapes, BufferGeometry, custom geometry, instancing. Use when creating 3D shapes, working with vertices, buil... |
| `threejs-interaction` | Three.js interaction - raycasting, controls, mouse/touch input, object selection. Use when handling user input, implementing click detection, addin... |
| `threejs-lighting` | Three.js lighting - light types, shadows, environment lighting. Use when adding lights, configuring shadows, setting up IBL, or optimizing lighting... |
| `threejs-materials` | Three.js materials - PBR, basic, phong, shader materials, material properties. Use when styling meshes, working with textures, creating custom shad... |
| `threejs-textures` | Three.js textures - texture types, UV mapping, environment maps, texture settings. Use when working with images, UV coordinates, cubemaps, HDR envi... |
| `turborepo-caching` | Configure Turborepo for efficient monorepo builds with local and remote caching. Use when setting up Turborepo, optimizing build pipelines, or impl... |
| `typescript-advanced-types` | Master TypeScript's advanced type system including generics, conditional types, mapped types, template literals, and utility types for building typ... |
| `ui-ux-pro-max` | UI/UX design intelligence. 50 styles, 21 palettes, 50 font pairings, 20 charts, 9 stacks (React, Next.js, Vue, Svelte, SwiftUI, React Native, Flutt... |
| `unreal-engine-cpp-pro` | Expert guide for Unreal Engine 5.x C++ development, covering UObject hygiene, performance patterns, and best practices. |
| `using-superpowers` | Use when starting any conversation - establishes how to find and use skills, requiring Skill tool invocation before ANY response including clarifyi... |
| `verification-before-completion` | Use when about to claim work is complete, fixed, or passing, before committing or creating PRs - requires running verification commands and confirm... |
| `viral-generator-builder` | Expert in building shareable generator tools that go viral - name generators, quiz makers, avatar creators, personality tests, and calculator tools... |
| `vr-ar` | VR/AR development principles. Comfort, interaction, performance requirements. |
| `web-games` | Web browser game development principles. Framework selection, WebGPU, optimization, PWA. |
| `web-performance-optimization` | Optimize website and web application performance including loading speed, Core Web Vitals, bundle size, caching strategies, and runtime performance |
| `web3-testing` | Test smart contracts comprehensively using Hardhat and Foundry with unit tests, integration tests, and mainnet forking. Use when testing Solidity c... |
| `webapp-testing` | Toolkit for interacting with and testing local web applications using Playwright. Supports verifying frontend functionality, debugging UI behavior,... |
| `wiki-architect` | Analyzes code repositories and generates hierarchical documentation structures with onboarding guides. Use when the user wants to create a wiki, ge... |
| `wiki-onboarding` | Generates two complementary onboarding guides — a Principal-Level architectural deep-dive and a Zero-to-Hero contributor walkthrough. Use when the ... |
| `wiki-vitepress` | Packages generated wiki Markdown into a VitePress static site with dark theme, dark-mode Mermaid diagrams with click-to-zoom, and production build ... |
| `writing-plans` | Use when you have a spec or requirements for a multi-step task, before touching code |
| `zod-validation-expert` | Expert in Zod — TypeScript-first schema validation. Covers parsing, custom errors, refinements, type inference, and integration with React Hook For... |
| `zustand-store-ts` | Create Zustand stores with TypeScript, subscribeWithSelector middleware, and proper state/action separation. Use when building React state manageme... |

[↑ Voltar ao Índice](#indice-de-categorias)

---

