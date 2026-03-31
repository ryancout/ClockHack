# Processador de Planilhas FAS

Aplicativo desktop em Python para tratar planilhas Excel/CSV de banco de horas, com filtro por departamento, linha TOTAL, destaque visual de saldos críticos e geração automática das abas **RANKING** e **RESUMO**.

## Recursos

- processamento em lote de arquivos `.xlsx` e `.csv`
- filtro por departamento antes do cálculo
- cálculo de **Banco Total** e **Banco Saldo**
- linha **TOTAL** na aba principal preservando a planilha tratada
- destaque visual para saldos menores que `-8:00` e maiores que `8:00`
- aba **RANKING** com top devedores e top horas extras
- aba **RESUMO** com total por departamento
- interface desktop com CustomTkinter
- preferências, logs e histórico gravados fora da pasta do projeto

## Estrutura

- `main.py` — ponto de entrada
- `app/` — código-fonte do app
- `build_tools/` — scripts de build e instalador
- `tests/` — testes unitários básicos
- `main.spec` — build do PyInstaller

## Como rodar em desenvolvimento

```bash
pip install -r requirements.txt
python main.py
```

## Como gerar o EXE

```bat
build_tools\gerar_exe.bat
```

Saída esperada:

```text
dist\ProcessadorPlanilhasFAS.exe
```

## Como gerar o instalador

1. Gere o EXE primeiro.
2. Abra o `build_tools\ProcessadorPlanilhasFAS.iss` no Inno Setup.
3. Compile o instalador.

## Observações de segurança

- logs, histórico, auditoria e preferências são gravados em pasta de usuário
- o projeto não envia dados para internet
- o zip de distribuição não inclui `.git`, `build`, `dist`, `logs` nem `data`
