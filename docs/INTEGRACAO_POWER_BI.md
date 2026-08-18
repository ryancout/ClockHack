# Integração Power BI

## Objetivo

Depois de gerar um Excel por CSV ou RHiD, a tela de sucesso permite enviar os
dados tratados ao workspace **FAS Jornada Analytics**. Cada envio recebe um
`IDRelatorio` UUID próprio e acrescenta novas linhas ao modelo, sem alterar o
arquivo Excel independente que foi salvo pelo usuário.

O aplicativo não usa banco de dados próprio. O histórico analítico enviado fica
armazenado no modelo semântico do Power BI.

## Autenticação

A integração é um aplicativo público/desktop do Microsoft Entra ID:

- usa MSAL e login interativo no navegador;
- solicita apenas o escopo delegado `Dataset.ReadWrite.All`;
- usa o URI de redirecionamento `http://localhost`;
- não possui client secret;
- não recebe nem armazena a senha Microsoft;
- mantém o access token somente na memória da execução.

O usuário autenticado precisa ter acesso de gravação ao workspace configurado.
Em outra máquina, o executável funciona com o mesmo registro de aplicativo, mas
cada usuário deve entrar com uma conta corporativa autorizada.

## Modelo analítico

O aplicativo cria ou reutiliza o modelo Push `FAS Jornada Analytics v2` e sua tabela
`Jornada`. A tabela é desnormalizada para que o analista possa montar os visuais
sem criar relacionamentos antes do primeiro uso.

O nome do modelo é versionado quando o contrato de colunas muda. Isso permite
criar a estrutura nova sem apagar automaticamente modelos ou históricos antigos
do workspace.

Dimensões principais:

- `IDRelatorio`, `GeradoEm`, `PeriodoInicial`, `PeriodoFinal`;
- `Origem`, `Empresa`, `SetoresSelecionados`, `Arquivo`;
- `Matricula`, `Funcionario`, `Departamento`.

Indicadores principais:

- horas normais, trabalhadas, previstas, extras, abono e falta/atraso;
- Banco Total e Banco Saldo;
- faltas em dias;
- valores de tempo em minutos e horas decimais;
- percentual trabalhado e classificação do saldo.

O CPF não faz parte do esquema e nenhum valor dessa coluna é enviado.

## Fluxo

1. O usuário gera ou importa e trata um relatório normalmente.
2. Na tela **Relatório salvo**, escolhe **Enviar ao Power BI**.
3. O app extrai as linhas da aba principal do XLSX, ignora TOTAL e linhas sem
   identidade, converte horas pelas mesmas funções do domínio e cria o snapshot.
4. Uma impressão digital do conteúdo é comparada aos envios anteriores. Se o
   conteúdo já existir, o usuário precisa confirmar o reenvio.
5. O navegador realiza o login Microsoft.
6. O app localiza ou cria o modelo no workspace e envia as linhas em lotes.
7. A tela mostra a quantidade enviada e o `IDRelatorio`.
8. O app cria um relatório fino conectado ao modelo e o abre no Power BI
   Desktop. O botão passa a exibir **Abrir no Power BI Desktop**, sem repetir o
   envio das linhas.

## Verificação de conexões

A página **Verificar conexões** executa somente leituras:

1. valida a sessão RHiD e lista a quantidade de empresas;
2. realiza o login Microsoft, quando necessário;
3. confirma acesso ao workspace consultando seus modelos;
4. localiza e valida a tabela do modelo configurado.

O diagnóstico não cria modelo, não adiciona linhas, não baixa CSV e não gera
Excel. Se o acesso RHiD ainda não estiver ativo, usa apenas as credenciais já
digitadas ou lembradas de forma segura na tela RHiD.

## Proteção contra duplicidade

O arquivo `data/powerbi_sends.json`, no perfil local do usuário, guarda somente
a impressão digital, IDs técnicos, data, nome simples do arquivo e quantidade de
linhas. Ele não contém funcionários, matrículas ou valores de jornada. A
impressão digital ignora `IDRelatorio`, horário de geração e nome do arquivo,
permitindo reconhecer o mesmo conteúdo mesmo após gerar outro XLSX.

A proteção evita repetição acidental, mas não bloqueia a operação: quando houver
uma necessidade real de reenviar, o usuário pode confirmá-la explicitamente.

## Power BI Desktop

O arquivo `definition.pbir` gerado fica na pasta local de dados do FAS Jornada.
Ele não contém dados, senha ou token: guarda somente a referência ao workspace e
ao modelo semântico. Ao abrir, o Power BI Desktop pode pedir o login corporativo.

A conexão é ao vivo e somente leitura para o modelo. O analista pode montar e
salvar seus gráficos normalmente; para enxergar o modelo, a conta precisa ter a
permissão **Build** no conjunto semântico.

## Limites e ciclo de vida

- o envio acrescenta linhas; ele não substitui execuções anteriores;
- a API Push só permite apagar todas as linhas de uma tabela, não apenas um
  `IDRelatorio`; exclusão seletiva exige reconstruir o modelo a partir dos
  arquivos arquivados ou adotar armazenamento persistente;
- a política `None` do modelo suporta até 5 milhões de linhas por tabela;
- a Microsoft informa que a criação de novos modelos em tempo real, incluindo
  modelos Push, permanece disponível até 31/10/2027; os modelos existentes não
  são afetados por esse marco;
- o FAS Jornada envia lotes independentes, não eventos contínuos. Por isso, o
  destino planejado é Fabric Lakehouse/OneLake, preservando um arquivo por
  relatório. Consulte [o plano de migração](MIGRACAO_POWER_BI_FABRIC.md).

## Configuração técnica

Os identificadores não secretos ficam em `app/core/config.py`:

- `POWER_BI_CLIENT_ID`;
- `POWER_BI_TENANT_ID`;
- `POWER_BI_WORKSPACE_ID`;
- `POWER_BI_WORKSPACE_NAME`;
- `POWER_BI_DATASET_NAME`;
- `POWER_BI_TABLE_NAME`.

Eles identificam o aplicativo, diretório e workspace, mas não concedem acesso
sem autenticação e permissão da conta Microsoft.
