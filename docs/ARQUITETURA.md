# Arquitetura

O aplicativo separa regras de negócio, geração de relatórios, infraestrutura e interface.

```text
UI (CustomTkinter)
        |
        v
Controller
        |
        v
Executor em segundo plano
   (fila + Tk after)
        |
        v
Pipeline de processamento
   |          |          |
   v          v          v
Domínio    Serviços    Relatórios
             |
             v
          CSV / XLSX
```

## Responsabilidades

- `app/domain/`: modelos imutáveis e políticas puras de identidade e tempo. Não conhece Tkinter, OpenPyXL ou arquivos.
- `app/reports/`: transforma registros de domínio nas abas auxiliares do Excel.
- `app/services/`: leitura, validação, cálculo, escrita, formatação e orquestração.
- `app/integrations/`: autenticação e geração/download do CSV nos serviços oficiais do RHiD; não conhece widgets.
- `app/controllers/`: traduz ações da interface em chamadas aos serviços.
- `app/ui/`: constrói widgets e apresenta o estado ao usuário.

## Fluxo do processamento

1. O CSV é validado, carregado e tem seu cabeçalho lido da primeira linha.
2. O filtro de departamento é aplicado.
3. Cada linha vira um `RegistroFuncionario`, mantendo horas em minutos.
4. A matrícula é reduzida aos três dígitos finais somente na saída.
5. Os totais da aba principal são calculados pelos serviços existentes.
6. Os geradores criam `SALDO`, `RANKING` e `RESUMO` sem recalcular horas.
7. A aba principal é formatada e o workbook é salvo em XLSX.

Na entrada direta pelo RHiD, a integração solicita o `extrato` CSV agrupado por
funcionário, com exatamente as 13 colunas do contrato, acompanha o processamento
remoto por GUID e valida que exista somente uma linha por matrícula antes de gravar
um CSV temporário.
Esse arquivo entra no mesmo `processar_arquivo`; portanto não existe um segundo
cálculo de horas. O temporário é removido ao final.

## Regras de dependência

- O domínio não importa módulos de outras camadas.
- Relatórios podem depender do domínio e de funções puras de formatação.
- A pipeline orquestra; regras de layout ficam nos módulos de relatório.
- A interface não lê CSV nem manipula workbooks diretamente.
- Diálogos e widgets permanecem na thread principal; somente a pipeline pesada roda no worker.
- O worker publica eventos em uma fila, consumida pelo loop do Tkinter com `after`.
- Alterações estruturais devem manter os testes de contrato verdes.

## Processamento assíncrono

Antes de iniciar o worker, o controlador cria um plano imutável com os arquivos,
opções e destinos. Sobrescritas e colisões de nomes são verificadas nessa
etapa, portanto nenhum arquivo é processado se o preflight for cancelado.

O `BackgroundTaskRunner` usa uma única thread serial para o OpenPyXL. Progresso,
conclusão e falhas atravessam uma fila; a interface consome esses eventos apenas
na thread principal. Isso evita chamadas Tkinter a partir do worker, impede duas
execuções simultâneas e mantém a janela responsiva. A thread não é daemon e a
janela recusa fechamento enquanto houver gravação, protegendo o XLSX em andamento.

## Compatibilidade protegida

A API pública usada pelo controlador continua sendo:

- `obter_departamentos(caminho_arquivo)`
- `processar_arquivo(caminho_arquivo, caminho_saida, ...)`

A fixture anônima em `tests/fixtures/` protege cabeçalhos, filtros, totais, abas, homônimos e privacidade da matrícula.
