# FAS Jornada

**Relatório e Análise de Jornada.** Aplicativo desktop em Python para tratamento de relatórios de banco de horas, com cálculo automático, filtros, destaques visuais e geração de análises em Excel.

**Versão atual: 8.1 — Windows 10/11**

[Baixar a versão mais recente](https://github.com/ryancout/ClockHack/releases/latest)

---

## Instalação

Na página de releases, escolha uma das opções:

- `FASJornada_Setup_8.1.exe` — instalador recomendado, com atalhos e desinstalação pelo Windows
- `FASJornada_v8.1_portable.zip` — versão portátil; extraia o ZIP e execute `FASJornada.exe`

O Windows pode exibir uma confirmação de segurança porque o aplicativo ainda não possui assinatura digital comercial. Confira se o arquivo foi baixado diretamente deste repositório antes de executá-lo.

---

## Recursos

- processamento em lote de arquivos de entrada `.csv`
- conexão direta ao RHiD para escolher empresa, um ou vários setores e período, sem download manual
- opção **Todos os setores**, mantendo apenas setores com funcionários ativos
- datas em `DD/MM/AAAA`, calendário e sem limite artificial de período no aplicativo
- filtro por departamento antes do cálculo
- cálculo de **Banco Total** e **Banco Saldo**
- linha **TOTAL** na aba principal preservando a planilha tratada
- identifica homônimos pela matrícula completa na aba principal e nas abas **SALDO** e **RANKING**
- remove a coluna CPF de todos os arquivos Excel gerados e não a solicita na integração RHiD
- destaque visual para saldos menores que `-8:00` e maiores que `8:00`
- aba **RANKING** com top devedores e top horas extras
- aba **RESUMO** com total por departamento
- navegação responsiva em uma única janela, sem barras de rolagem; a densidade visual se ajusta ao tamanho disponível
- processamento pesado em segundo plano, mantendo a janela responsiva
- cancelamento seguro do lote: o arquivo em andamento termina e os próximos não são iniciados
- preferências, logs e histórico gravados fora da pasta do projeto
- barra animada e bloqueio de ações duplicadas durante o processamento
- tempo de execução ao final
- validação de arquivos de entrada
- mensagens de erro amigáveis
- armazenamento opcional das credenciais RHiD no Gerenciador de Credenciais do Windows ou protegido por DPAPI
- envio opcional do relatório tratado ao Power BI, com login Microsoft no navegador, dados pré-tratados sem CPF e abertura do modelo no Power BI Desktop
- prevenção de reenvio acidental do mesmo conteúdo ao Power BI, com confirmação explícita quando a repetição for intencional
- tela **Verificar conexões** para testar RHiD, conta Microsoft, workspace e modelo Power BI sem gerar relatório
- fronteira de publicação preparada para substituir o modelo Push por Fabric Lakehouse/OneLake

---

## Estrutura

- `main.py` — ponto de entrada
- `app/domain/` — modelos e políticas de negócio sem dependência da interface
- `app/reports/` — geradores independentes das abas SALDO, RANKING e RESUMO
- `app/services/` — leitura, validação, cálculo, formatação e orquestração
- `app/controllers/` — fachada da janela e fluxos independentes de CSV, RHiD e Power BI
- `app/ui/` — interface desktop CustomTkinter
- `build_tools/` — scripts de build e instalador
- `tests/` — testes unitários e de contrato do CSV/Excel
- `main.spec` — build do PyInstaller
- `version_info.txt` — propriedades do executável

A documentação completa está no [índice técnico](docs/README.md), incluindo
arquitetura, referência de todos os módulos, fluxos CSV/RHiD/Power BI, segurança, testes
e distribuição.

---

## Como rodar em desenvolvimento

```bash
pip install -r requirements.txt
python main.py
```

Para desenvolver, testar e empacotar:

```bash
pip install -r requirements-dev.txt
```

### Cancelamento do processamento

Durante um lote, o botão **Cancelar processamento** interrompe o trabalho de forma cooperativa. O arquivo que já começou continua até ser salvo com segurança; os arquivos seguintes não são iniciados. A tela informa quantos arquivos foram salvos e mantém disponíveis os botões para abrir o resultado e a pasta de saída.

---

## Build versionado

```bat
build_tools\build_release_auto_version.bat
```

O script:

- confirma que `app/core/version.py`, `version_info.txt` e o instalador possuem a mesma versão
- executa testes, análise estática, compilação e verificação de dependências
- gera o executável, o instalador, o ZIP portátil e `SHA256SUMS.txt`
- assina os binários quando um certificado Authenticode estiver configurado
- não altera versão, commit, tag nem repositório remoto

Pushes e pull requests para `main` são verificados pelo GitHub Actions. Uma tag
`vX.X` ou `vX.X.X` dispara a criação da release somente quando a tag coincide
com todos os metadados de versão. Uma release existente nunca é sobrescrita.

---

## Como gerar o EXE manualmente

```bat
build_tools\gerar_exe.bat
```

Saída esperada:

```
dist\FASJornada.exe
```

---

## Como gerar o instalador

1. Gere o EXE primeiro
2. Abra o arquivo:

```
build_tools\FASJornada.iss
```

3. Compile no Inno Setup

---

## Distribuição

Os arquivos finais são gerados em:

```
dist/
```

Incluindo:

```
FASJornada_Setup_X.X.X.exe
FASJornada_vX.X.X_portable.zip
SHA256SUMS.txt
```

---

## Versionamento

A versão do aplicativo é controlada em:

- `app/core/version.py` → versão exibida no app
- `version_info.txt` → versão do executável (Windows)
- `build_tools/FASJornada.iss` → versão e nome do instalador

Os três devem ser atualizados na mesma alteração. O build e o workflow de release
interrompem a execução se encontrarem qualquer divergência.

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
- o diretório legado `ProcessadorPlanilhasFAS` é preservado para manter preferências e dados de versões anteriores
- a integração opcional comunica-se somente com os serviços HTTPS oficiais do RHiD/Control iD
- o acesso ao Power BI usa login Microsoft interativo, sem segredo de aplicativo e sem armazenar a senha
- lembrar o acesso RHiD é opcional; quando marcado, usuário, domínio e senha ficam no Gerenciador de Credenciais do Windows ou em arquivo binário protegido por DPAPI — nunca em JSON ou texto puro
- sem a opção de lembrar, a senha permanece apenas na memória durante a sessão
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
- MSAL
- PyInstaller
- Inno Setup

---

## Licença

Uso interno — FAS
