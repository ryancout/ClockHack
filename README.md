# Processador de Planilhas FAS

Aplicativo desktop em Python para tratamento de relatórios CSV de banco de horas, com cálculo automático, filtros, destaques visuais e geração de análises em Excel.

---

## Recursos

- processamento em lote de arquivos de entrada `.csv`
- conexão direta ao RHiD para escolher empresa, setor e período e gerar o Excel sem download manual
- filtro por departamento antes do cálculo
- cálculo de **Banco Total** e **Banco Saldo**
- linha **TOTAL** na aba principal preservando a planilha tratada
- identifica homônimos pelos 3 últimos dígitos da matrícula na aba principal e nas abas **SALDO** e **RANKING**
- destaque visual para saldos menores que `-8:00` e maiores que `8:00`
- aba **RANKING** com top devedores e top horas extras
- aba **RESUMO** com total por departamento
- interface desktop com CustomTkinter
- processamento pesado em segundo plano, mantendo a janela responsiva
- cancelamento seguro do lote: o arquivo em andamento termina e os próximos não são iniciados
- preferências, logs e histórico gravados fora da pasta do projeto
- barra animada e bloqueio de ações duplicadas durante o processamento
- tempo de execução ao final
- validação de arquivos de entrada
- mensagens de erro amigáveis

---

## Estrutura

- `main.py` — ponto de entrada
- `app/domain/` — modelos e políticas de negócio sem dependência da interface
- `app/reports/` — geradores independentes das abas SALDO, RANKING e RESUMO
- `app/services/` — leitura, validação, cálculo, formatação e orquestração
- `app/controllers/` — coordenação entre a interface e os casos de uso
- `app/ui/` — interface desktop CustomTkinter
- `build_tools/` — scripts de build e instalador
- `tests/` — testes unitários e de contrato do CSV/Excel
- `main.spec` — build do PyInstaller
- `version_info.txt` — propriedades do executável

Os limites entre as camadas estão documentados em [docs/ARQUITETURA.md](docs/ARQUITETURA.md).

---

## Como rodar em desenvolvimento

```bash
pip install -r requirements.txt
python main.py
```

### Cancelamento do processamento

Durante um lote, o botão **Cancelar processamento** interrompe o trabalho de forma cooperativa. O arquivo que já começou continua até ser salvo com segurança; os arquivos seguintes não são iniciados. A tela informa quantos arquivos foram salvos e mantém disponíveis os botões para abrir o resultado e a pasta de saída.

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
- a integração opcional comunica-se somente com os serviços HTTPS oficiais do RHiD/Control iD
- a senha do RHiD não é salva; o token permanece apenas na memória durante a execução
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
