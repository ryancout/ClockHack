# Processador de Planilhas FAS

Aplicativo desktop em Python para tratamento de planilhas Excel/CSV de banco de horas, com cálculo automático, filtros, destaques visuais e geração de análises.

---

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
- barra de progresso durante processamento
- tempo de execução ao final
- validação de arquivos de entrada
- mensagens de erro amigáveis

---

## Estrutura

- `main.py` — ponto de entrada
- `app/` — código-fonte do app
- `build_tools/` — scripts de build e instalador
- `tests/` — testes unitários básicos
- `main.spec` — build do PyInstaller
- `version_info.txt` — propriedades do executável

---

## Como rodar em desenvolvimento

```bash
pip install -r requirements.txt
python main.py
```

---

## Build automático (recomendado)

```bat
build_tools\build_release_auto_version.bat
```

O script:

- solicita a versão
- atualiza `app/core/version.py`
- atualiza `version_info.txt`
- limpa builds anteriores
- gera o executável (.exe)
- cria o arquivo `.zip`
- envia automaticamente para o GitHub

---

## Como gerar o EXE manualmente

```bat
build_tools\gerar_exe.bat
```

Saída esperada:

```
dist\ProcessadorPlanilhasFAS.exe
```

---

## Como gerar o instalador

1. Gere o EXE primeiro
2. Abra o arquivo:

```
build_tools\ProcessadorPlanilhasFAS.iss
```

3. Compile no Inno Setup

---

## Distribuição

Os arquivos finais são gerados em:

```
releases/
```

Sempre contendo apenas a versão mais recente:

```
ProcessadorPlanilhasFAS_vX.X.X.zip
```

---

## Versionamento

A versão do aplicativo é controlada em:

- `app/core/version.py` → versão exibida no app
- `version_info.txt` → versão do executável (Windows)

Ambos são atualizados automaticamente pelo script de build.

---

## Identidade visual

O aplicativo utiliza um ícone único (`app/assets/icon.ico`) aplicado em:

- executável (.exe)
- janela do aplicativo
- instalador
- atalhos do sistema

---

## Observações de segurança

- logs, histórico, auditoria e preferências são gravados em pasta do usuário
- o aplicativo não envia dados para a internet
- o pacote de distribuição não inclui:
  - `.git`
  - `build`
  - `dist`
  - `logs`
  - `data`

---

## Tecnologias utilizadas

- Python 3.11+
- CustomTkinter
- OpenPyXL
- PyInstaller
- Inno Setup

---

## Licença

Uso interno — FAS
