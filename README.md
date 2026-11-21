# Gerador EAP — Planejamento de Projetos de Engenharia

Aplicação web (HTML + JavaScript, sem servidor) que transforma a **documentação de referência de um projeto** — proposta técnica, contrato, termo de referência ou edital — em uma **Estrutura Analítica do Projeto (EAP) segundo as boas práticas do PMBOK**, pronta para abrir no **Microsoft Project**.

Também gera **Lista de Documentos (LD)** codificada, **EAP-Base** e **cronograma** a partir de uma base de codificação por segmento (Metrô, Aeroportuário, Edificações, Ferroviário, Rodoviário).

> **Projeto de demonstração.** A base de códigos e a EAP de referência são fictícias. Substitua-os pelos padrões da sua organização na aba *Codificação*.

Tudo roda no navegador: os dados ficam no seu computador. Nada é enviado a servidores, exceto se você optar pela análise com IA.

---

## Como usar

1. Baixe o repositório (**Code › Download ZIP**) e extraia a pasta, ou clone com `git clone`.
2. Abra o arquivo **`index.html`** no Chrome ou no Edge (dois cliques).

Não é preciso instalar nada para gerar a EAP, a LD, os arquivos Excel e o XML do MS Project.

> Mantenha a estrutura de pastas (`vendor/`, `data/`, `js/`, `css/`): o `index.html` carrega os arquivos delas.

---

## Análise de Documentação → EAP (PMBOK)

1. Arraste os arquivos (PDF com texto, DOCX, TXT/MD) ou cole o texto.
2. Informe a data de início e, se quiser, o nome do projeto.
3. Escolha o modo de análise:
   - **Automática (offline):** regras identificam seções de escopo, entregáveis, fases, prazo contratual, premissas, restrições, exclusões, critérios de aceitação e riscos. Nada sai do computador.
   - **Com IA (opcional):** usa a API da Anthropic (Claude) com a **sua** chave. Mais precisa em documentos longos ou pouco estruturados. *O texto dos documentos é enviado à API — confirme se a confidencialidade do contrato permite.*
   - **Manual com IA (sem chave):** *Copiar prompt* → cole no Claude → copie o JSON da resposta → *Importar JSON*.
4. Revise e edite a EAP: nomes, tipo (pacote ou marco), duração, predecessoras (pelo código EDT), nível e ordem.
5. Confira a **verificação de boas práticas**, o **dicionário da EAP** e a **declaração de escopo** extraída.
6. Exporte.

### Boas práticas aplicadas (PMBOK / Practice Standard for WBS)

| Prática | Como é aplicada |
|---|---|
| Orientação a entregas | Elementos nomeados como substantivos; verbos iniciais ("Elaborar…") são removidos e sinalizados |
| Regra dos 100% | Ramo de **Gerenciamento do Projeto** sempre presente (iniciação, planejamento, monitoramento e controle, encerramento) + todo o escopo técnico encontrado; exclusões ficam fora e são listadas |
| Decomposição | Resumos com 2+ filhos; pacotes de trabalho no nível mais baixo; alerta acima de 6 níveis |
| Marcos | Kick-off, aprovação ao final de cada fase e encerramento (duração zero) |
| Rede de dependências | Fases em sequência término-início; atividades sem sucessora ligadas ao aceite final |
| Prazo | Prazo contratual identificado é distribuído entre as fases; alerta se a duração estimada exceder o prazo |
| Dicionário da EAP | Descrição, critério de aceitação e fonte (seção do documento) por pacote e marco |

### Exportações

| Formato | Uso |
|---|---|
| **MS Project (.xml)** | Abre direto no MS Project (*Arquivo › Abrir › Formato XML*): hierarquia, EDT, durações, marcos, predecessoras e dicionário (em *Anotações*). Depois, *Salvar como* `.mpp`. |
| **.mpp via helper** | Gera o `.mpp` nativo automaticamente (requer Windows + MS Project + helper local). |
| **Excel (.xlsx)** | Abas *EAP*, *Dicionário da EAP* e *Declaração de Escopo*. |
| **CSV** | Tabela completa (separador `;`, UTF-8). |
| **JSON** | Salva o projeto para continuar depois (*Abrir projeto salvo*). |

O trabalho em andamento também é salvo automaticamente no navegador.

---

## Lista de Documentos, EAP-Base e Cronograma

1. Aba **Configuração da LD**: informe centro de custo, ordem de serviço e segmento; marque os itens da codificação.
2. **Gerar combinações** e **Gerar documentos**: o sistema cria os códigos no formato `CC-OS-FASE-ÁREA-SUBÁREA-LOCAL-DISCIPLINA-SISTEMA-TIPO-NNNN`, a LD, a EAP-Base e o cronograma.
3. Aba **Exportações**: Excel, CSV no formato do MS Project ou `.mpp` via helper.
4. Aba **Codificação**: edite, exporte e importe a base de códigos (JSON) ou restaure a base de demonstração.

---

## Helper local para arquivos .mpp (opcional)

O navegador não grava `.mpp` sozinho. O helper em Python automatiza o MS Project:

1. Instale o **Python 3** (marque *Add Python to PATH*) e tenha o **Microsoft Project** instalado.
2. Na pasta `helper/`, execute **`INICIAR_HELPER_MPP.bat`** (instala o `pywin32` e sobe o servidor em `http://127.0.0.1:8765`).
3. No app, aba **Exportações › Verificar helper local**.
4. Em caso de problema, rode `DIAGNOSTICO_HELPER.bat`.

O helper escuta apenas em `127.0.0.1` (a sua própria máquina).

---

## Estrutura

```
index.html                 Aplicação
css/tema.css               Design system "Aço & Azul" (tokens, componentes)
css/analise-eap.css        Estilos da aba Análise de Documentação
js/analise-eap.js          Análise de documentação → EAP PMBOK e exportações
js/melhorias-v6.js         Backup e restauração da base de códigos
data/seg-codigos.js        Base de códigos de demonstração por segmento
data/eap-referencia.js     EAP de referência de demonstração
helper/                    Helper Python para gerar .mpp
vendor/                    Bibliotecas e fontes de terceiros (uso offline)
```

## Design system

Visual corporativo técnico: azul-marinho `#0B2545`/`#0B3D91`, laranja de segurança `#E8590C`/`#C2410C`, neutros em tons de aço, tipografia **IBM Plex Sans** e **IBM Plex Mono** (códigos EDT, datas e números). Cores e tamanhos atendem ao contraste mínimo WCAG AA.

## Bibliotecas e fontes de terceiros

| Componente | Uso | Licença |
|---|---|---|
| [SheetJS (xlsx) 0.18.5](https://sheetjs.com) | Leitura e escrita de Excel | Apache-2.0 |
| [PDF.js 3.11](https://mozilla.github.io/pdf.js/) | Leitura de PDF | Apache-2.0 |
| [mammoth.js 1.8](https://github.com/mwilliamson/mammoth.js) | Leitura de DOCX | BSD-2-Clause |
| [IBM Plex](https://github.com/IBM/plex) | Tipografia | SIL OFL 1.1 |

## Limitações conhecidas

- PDFs digitalizados (imagem) precisam de OCR antes; o app avisa quando não há texto.
- Arquivos `.doc` antigos não são lidos — salve como `.docx` ou PDF.
- A análise automática depende da estrutura do documento (títulos numerados, listas). Para documentos pouco estruturados, prefira o modo com IA e sempre revise a EAP.
- As datas exibidas são uma prévia em dias úteis (seg–sex, sem feriados); o MS Project recalcula pelo calendário do projeto.

## Licença

[MIT](LICENSE) © 2026 Cristiano Lelis. Histórico de versões em [CHANGELOG.md](CHANGELOG.md).
