# Referência dos módulos

## Entrada e configuração

| Arquivo | Responsabilidade |
|---|---|
| `main.py` | Configura o logger e inicia a interface. Não executa nada quando apenas importado. |
| `app/core/config.py` | Caminhos, colunas obrigatórias, limites e identidade visual. |
| `app/core/version.py` | Nome e versão exibidos pelo aplicativo e usados no build. |
| `app/core/exceptions.py` | Exceções de negócio apresentáveis ao usuário. |
| `app/core/logger.py` | Logger rotativo no diretório de dados do usuário. |

## Domínio

| Arquivo | Responsabilidade / API principal |
|---|---|
| `app/domain/models.py` | `RegistroFuncionario`, registro imutável com horas armazenadas em minutos. |
| `app/domain/time.py` | `para_minutos` e `formatar_horas`; suporta negativos e valores acima de 24 horas. |
| `app/domain/identity.py` | `normalizar_matricula`; preserva a matrícula completa como texto. |
| `app/domain/__init__.py` | Fachada pública do domínio. |

## Serviços

| Arquivo | Responsabilidade / API principal |
|---|---|
| `reader_service.py` | Converte CSV UTF-8 com BOM e delimitador `;` em workbook. |
| `validator_service.py` | Valida extensão, tamanho, cabeçalho e resultado mínimo. |
| `filter_service.py` | Lista departamentos e remove linhas fora do filtro escolhido. |
| `calculator_service.py` | Soma Banco Total/Saldo com as funções do domínio e conta pessoas. |
| `writer_service.py` | Escreve TOTAL, legenda e destaques na aba principal. |
| `worksheet_formatting_service.py` | Formata cabeçalhos, bordas, alinhamentos e larguras. |
| `workbook_pipeline_service.py` | Caso de uso central: leitura, filtro, registros, privacidade, abas e gravação. |
| `file_service.py` | Nomes de saída, extensão, slug e nome curto. |
| `background_task_runner.py` | Executa uma tarefa serial em thread e entrega eventos à thread da interface. |
| `preferences_service.py` | Preferências em JSON com escrita atômica. |
| `history_service.py` | Histórico sanitizado e limitado. |
| `audit_service.py` | Auditoria sanitizada e limitada. |
| `rhid_credentials_service.py` | Cofre Windows/DPAPI para credenciais lembradas. |
| `analytics_service.py` | Extrai do XLSX um snapshot analítico sem CPF e calcula indicadores derivados. |

## Relatórios Excel

| Arquivo | Responsabilidade |
|---|---|
| `app/reports/saldo.py` | Aba SALDO com matrícula completa e dados de banco. |
| `app/reports/ranking.py` | Rankings de maiores devedores e maiores saldos extras. |
| `app/reports/resumo.py` | Totais de Banco Saldo por departamento. |
| `app/reports/styles.py` | Objetos de estilo compartilhados pelo resumo. |

Os geradores recebem registros já convertidos; não reinterpretam o CSV nem
implementam um segundo cálculo de horas.

## Integração RHiD

| Arquivo | Responsabilidade / API principal |
|---|---|
| `app/integrations/rhid_client.py` | Cliente HTTPS, login, clientes, empresas, setores, pessoas, job, polling e download CSV. |
| `app/integrations/rhid_report_service.py` | Valida o CSV consolidado, cria temporário e o envia à pipeline comum. |

## Integração Power BI

| Arquivo | Responsabilidade / API principal |
|---|---|
| `app/integrations/powerbi_client.py` | Login Microsoft interativo, criação/reuso do modelo Push e envio em lotes. |
| `app/integrations/powerbi_destination.py` | Contrato de publicação e adaptador do modelo Push, separando o transporte para a migração Fabric. |
| `app/services/analytics_service.py` | Converte o Excel salvo em linhas analíticas identificadas por `IDRelatorio`. |
| `app/services/powerbi_desktop_service.py` | Cria um PBIR fino ligado ao modelo publicado e solicita sua abertura no Power BI Desktop. |
| `app/services/powerbi_send_registry.py` | Calcula a impressão digital do conteúdo e registra envios para evitar duplicidade acidental. |

## Controller e interface

| Arquivo | Responsabilidade |
|---|---|
| `app/controllers/main_controller.py` | Fachada estável consumida pela janela e composição dos fluxos especializados. |
| `app/controllers/csv_workflow.py` | Seleção, preflight, lote, progresso, cancelamento e resultados CSV. |
| `app/controllers/rhid_workflow.py` | Login, catálogo, escopo e geração direta pelo RHiD. |
| `app/controllers/powerbi_workflow.py` | Preparação, deduplicação, envio e abertura no Power BI Desktop. |
| `app/controllers/connection_diagnostics_workflow.py` | Testes independentes de RHiD, Microsoft, workspace e modelo sem gerar relatório. |
| `app/controllers/workflow_state.py` | Estado mínimo compartilhado entre os três fluxos. |
| `app/ui/main_window.py` | Janela, cabeçalho, navegação, atalhos e fachada usada pelo controller. |
| `app/ui/navigation.py` | Enum e regras de retorno entre páginas. |
| `app/ui/view_state.py` | Máquina de estados e configuração dos controles. |
| `app/ui/report_pages.py` | Páginas inicial, diagnóstico, CSV, processamento e sucesso. |
| `app/ui/rhid_page.py` | Login, escolha de cliente/setores, período, opções e progresso RHiD. |
| `app/ui/responsive.py` | Perfis puros de densidade normal, compacta e densa para manter as páginas sem rolagem. |

## Build e recursos

| Arquivo | Responsabilidade |
|---|---|
| `main.spec` | Receita PyInstaller. |
| `version_info.txt` | Metadados do executável Windows. |
| `build_tools/FASJornada.iss` | Receita do instalador Inno Setup. |
| `build_tools/build_release.ps1` | Build local validado, assinatura opcional, instalador, ZIP e hashes. |
| `build_tools/verify_release_version.py` | Impede release com tag ou metadados de versão divergentes. |
| `.github/workflows/ci.yml` | Valida pushes e pull requests para `main`. |
| `.github/workflows/release.yml` | Gera e publica artefatos imutáveis quando uma tag de versão é criada. |
| `app/assets/` | Única fonte de ícone, logos e padrão gráfico. |
