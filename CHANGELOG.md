# Histórico de versões

## v6.0 — 2026-11-21

### Novo
- **Análise de Documentação → EAP (PMBOK)**, agora a primeira aba: lê PDF, DOCX e TXT (proposta técnica, contrato, termo de referência, edital) e gera a EAP:
  - análise automática offline por regras, análise com IA (Claude, chave do usuário) e modo manual "copiar prompt / importar JSON";
  - ramo de Gerenciamento do Projeto, fases a partir da estrutura do documento, marcos de aprovação, durações distribuídas pelo prazo contratual e predecessoras;
  - editor da EAP (renomear, tipo, duração, predecessoras por EDT, mover, mudar nível, incluir e excluir);
  - verificação de boas práticas (regra dos 100%, orientação a entregas, filho único, níveis, marcos, rede aberta, ciclos, prazo);
  - dicionário da EAP e declaração de escopo extraída (objeto, prazos, premissas, restrições, exclusões, critérios de aceitação, riscos);
  - exportação **MS Project XML** (abre direto no MS Project, sem helper), `.mpp` via helper, Excel, CSV e JSON;
  - salvamento automático no navegador.
- **Design system "Aço & Azul"**: novo visual corporativo técnico (azul-marinho, laranja de segurança, IBM Plex), cabeçalho e navegação reorganizados.
- **Codificação**: exportar e importar a base de códigos em JSON e restaurar a base padrão.
- **Exportações**: CSV no formato do MS Project para EAP-Base e cronograma; Lista de Documentos em Excel com modelo próprio (abas *Informações do Projeto* e *LD*).

### Alterado
- Projeto de demonstração: base de códigos e EAP de referência genéricas e fictícias; campos da codificação com nomes genéricos (Fase, Área, Subárea, Local, Disciplina, Sistema, Tipo).
- Combinações sem limite fixo continuam permitidas, mas acima de 5.000 o app mostra o total e pede confirmação.
- Bibliotecas e fontes incluídas em `vendor/` para funcionar offline; dados movidos para `data/`.
- Helper `.mpp` 2.1: aceita marcos (duração 0), anotações, EDT em Texto1 e data de início do projeto; compatível com as chamadas anteriores.

### Corrigido
- "Exportar tudo (.xlsx)" falhava quando havia cronograma.
