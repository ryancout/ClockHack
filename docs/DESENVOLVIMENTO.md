# Desenvolvimento e testes

## Ambiente

- Windows 10/11;
- Python 3.11 ou superior;
- dependências de execução em `requirements.txt`;
- dependências de desenvolvimento em `requirements-dev.txt`;
- Inno Setup 6 apenas para o instalador.

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements-dev.txt
python main.py
```

## Validação padrão

```powershell
python -m pytest -q
pyright app
python -m compileall -q app main.py
python -m pip check
git diff --check
```

## Organização dos testes

| Teste | Cobertura principal |
|---|---|
| `test_domain.py`, `test_time_service.py` | identidade, sinais e formatação de horas |
| `test_csv_contract.py` | fixture CSV, filtros, abas, totais, matrícula e ausência de CPF |
| `test_reports.py` | SALDO, RANKING e RESUMO |
| `test_rhid_client.py` | protocolo HTTP, catálogo, payload, polling e erros |
| `test_rhid_report_service.py` | contrato consolidado e ponte para a pipeline |
| `test_rhid_credentials_service.py` | WinCred, DPAPI e falhas seguras |
| `test_main_controller_states.py` | preflight, lote, progresso, cancelamento e falhas parciais |
| `test_background_task_runner.py` | isolamento de thread e prevenção de reentrada |
| `test_report_pages.py`, `test_rhid_page.py` | lógica da interface sem janela real |
| `test_persistence.py` | schemas e escrita dos arquivos auxiliares |
| `test_product_identity.py` | nome, versão, AppId e diretório legado |

## Convenções

- regras puras pertencem ao domínio; widgets não entram em serviços;
- chamadas Tkinter ocorrem somente na thread principal;
- processamento OpenPyXL é serial para limitar memória e preservar ordem;
- funções internas começam com `_`; contratos entre camadas usam dataclasses imutáveis;
- edições de cálculo exigem teste de caracterização antes da alteração;
- arquivos reais com dados de funcionários não entram no repositório;
- fixtures devem ser anônimas e claramente sintéticas.

## Mudanças de versão

Atualize `CHANGELOG.md`, versão do app, metadados do Windows e instalador. Para
qualquer correção posterior a uma release publicada, crie uma nova versão em vez
de substituir silenciosamente os binários associados à tag anterior.
