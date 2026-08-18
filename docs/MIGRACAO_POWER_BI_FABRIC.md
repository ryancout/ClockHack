# Migração do Power BI Push para Microsoft Fabric

## Decisão arquitetural

O FAS Jornada continuará usando o modelo semântico Push enquanto a migração é
preparada. A Microsoft informa que a criação de novos modelos em tempo real,
incluindo Push, permanece habilitada até **31/10/2027**; os modelos existentes
não são afetados por essa data.

O fluxo do aplicativo é de lote: cada relatório tratado representa um snapshot
independente. Por isso, o destino escolhido para a migração é um **Fabric
Lakehouse em OneLake**, e não Eventstream. Cada envio deverá criar um arquivo
novo e imutável, permitindo auditoria, reprocessamento e construção posterior de
uma tabela Delta e de um modelo semântico Direct Lake.

Referências oficiais:

- https://learn.microsoft.com/power-bi/connect-data/service-real-time-streaming
- https://learn.microsoft.com/fabric/onelake/onelake-access-api
- https://learn.microsoft.com/fabric/data-engineering/lakehouse-api

## Estado atual do código

- `analytics_service.py` produz um snapshot desnormalizado e sem CPF;
- `PowerBiDestination` isola o transporte do restante do fluxo;
- `PushSemanticModelDestination` mantém o comportamento atual;
- cada resultado registra `powerbi_destination` e `powerbi_resource_id`, além
  das chaves legadas necessárias ao Power BI Desktop;
- a tela de diagnóstico confirma autenticação, workspace e modelo sem criar
  recursos ou enviar dados.

## Destino planejado

Estrutura sugerida no OneLake:

```text
Files/fas-jornada/
  ano=2026/
    mes=08/
      <IDRelatorio>.jsonl
```

Cada arquivo deve conter as mesmas colunas analíticas atuais, uma linha por
funcionário. O nome usa `IDRelatorio`, nunca nome, matrícula ou CPF. O upload deve
usar criação exclusiva; uma colisão de ID é erro e não sobrescrita.

## Etapas

1. **Provisionamento:** confirmar capacidade Fabric e criar um Lakehouse de
   desenvolvimento no workspace corporativo.
2. **Permissões:** registrar os escopos delegados mínimos exigidos pelas APIs
   Fabric e OneLake, preservando login interativo sem segredo no desktop.
3. **Adaptador:** implementar `FabricLakehouseDestination` usando a API Fabric
   para descobrir o Lakehouse e a API ADLS compatível do OneLake para gravar um
   arquivo independente por snapshot.
4. **Tabela:** carregar os arquivos em uma tabela Delta com chave composta por
   `IDRelatorio` e `Matricula`; rejeitar duplicidade em vez de somar novamente.
5. **Modelo:** criar um modelo semântico Direct Lake e validar os gráficos com os
   mesmos indicadores atuais.
6. **Execução paralela:** por um período controlado, comparar contagem, totais e
   saldos entre Push e Lakehouse, sem expor CPF.
7. **Corte:** trocar a fábrica de destino, manter o Push somente para consulta e
   remover sua criação automática antes de 31/10/2027.

## Critérios de aceite

- um arquivo novo por integração, sem sobrescrita;
- mesma quantidade de funcionários e mesmos totais do Excel de origem;
- sinais negativos e horas acima de 24 preservados;
- nenhuma coluna ou valor de CPF;
- reenvio idêntico bloqueado pela impressão digital atual;
- falha de rede não deixa arquivo parcial visível;
- diagnóstico confirma workspace e Lakehouse antes de habilitar o corte;
- reversão para o adaptador Push durante a fase paralela sem mudar a pipeline.
