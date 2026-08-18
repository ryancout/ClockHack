# Integração RHiD

Esta integração foi implementada no próprio FAS Jornada. Ela é um cliente dos
serviços HTTPS do RHiD/Control iD; o projeto não hospeda uma API intermediária e
não exige banco de dados.

## Autenticação

- login por e-mail, senha e, quando a conta possui mais de um cliente, domínio;
- token Bearer mantido somente na instância do cliente;
- identificação do cliente enviada no cabeçalho esperado pelo RHiD;
- “Esqueci minha senha” abre `https://www.rhid.com.br/v2/#/forgot_password`.

Os endpoints e formatos usados foram derivados da documentação Swagger oficial e
do comportamento do aplicativo web do RHiD. Como parte desses caminhos é interna
ao produto, alterações do fornecedor podem exigir atualização do cliente.

## Catálogo e escopo

O cliente carrega empresas, departamentos e pessoas. A lista exibida contém
somente setores relacionados a funcionários ativos. IDs são usados internamente
para não confundir homônimos, embora a interface mostre apenas os nomes.

Na geração, os IDs das pessoas ativas do escopo escolhido formam `listIdStr`.
Nenhuma seleção vazia é enviada como se significasse “todos do cliente”.

## Relatório remoto

O app solicita um extrato em CSV agrupado por pessoa (`agrupamento=person`), com
período no formato `yyyyMMdd`. As propriedades das colunas são obtidas do catálogo
do próprio tenant, priorizando `className + headerText`; isso evita confundir
“Previsto” com “Horas Previstas” ou extras por percentual com “Extras Total”.

Sequência remota:

1. POST de criação do relatório;
2. leitura de `guid`, `numPeople` e erros;
3. polling periódico do status do GUID;
4. download binário do CSV ao atingir 100%;
5. validação local antes de produzir Excel.

## Contrato e proteção do cálculo

O retorno deve conter as colunas operacionais esperadas e uma matrícula não vazia
e única por linha. Duplicidade bloqueia a geração: o aplicativo não soma snapshots
diários nem tenta adivinhar qual saldo é o correto. CPF não é solicitado e também
não faz parte do contrato do retorno.

O cliente HTTP não usa OpenPyXL. `rhid_report_service.py` faz a ponte para
`processar_arquivo`, garantindo que CSV manual e RHiD compartilhem exatamente o
mesmo cálculo.

## Diagnóstico

Os logs registram apenas metadados úteis, como presença de GUID, quantidade de
pessoas, percentual e tamanho baixado. Tokens, senhas, CPF e conteúdo das linhas
não devem ser registrados.
