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

## Controller e interface

| Arquivo | Responsabilidade |
|---|---|
| `app/controllers/main_controller.py` | Orquestra diálogos, preflight, lotes, integração, histórico e estados. |
| `app/ui/main_window.py` | Janela, cabeçalho, navegação, atalhos e fachada usada pelo controller. |
| `app/ui/navigation.py` | Enum e regras de retorno entre páginas. |
| `app/ui/view_state.py` | Máquina de estados e configuração dos controles. |
| `app/ui/report_pages.py` | Páginas inicial, CSV, processamento e sucesso. |
| `app/ui/rhid_page.py` | Login, escolha de cliente/setores, período, opções e progresso RHiD. |

## Build e recursos

| Arquivo | Responsabilidade |
|---|---|
| `main.spec` | Receita PyInstaller. |
| `version_info.txt` | Metadados do executável Windows. |
| `build_tools/FASJornada.iss` | Receita do instalador Inno Setup. |
| `build_tools/*.bat` | Automação de EXE, instalador e release versionada. |
| `app/assets/` | Única fonte de ícone, logos e padrão gráfico. |
