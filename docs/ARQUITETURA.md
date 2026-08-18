# Arquitetura

O aplicativo separa regras de negócio, geração de relatórios, infraestrutura e interface.

```text
UI (CustomTkinter)
        |
        v
MainController (fachada)
        |
        v
Fluxos CSV / RHiD / Power BI
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
- a saída opcional para Power BI passa por um adaptador analítico sem CPF e por
  um cliente REST; ambos executam fora da thread da interface.
- `app/controllers/`: o `MainController` mantém a API da janela e delega para
  `CsvWorkflow`, `RhidWorkflow`, `PowerBiWorkflow` e o diagnóstico de conexões.
  Os fluxos compartilham apenas preferências, seleção e último resultado por
  meio de `WorkflowState`.
- `app/ui/`: constrói widgets e apresenta o estado ao usuário.

## Interface responsiva

A janela usa páginas fixas, sem `CTkScrollableFrame`. O módulo
`app/ui/responsive.py` seleciona um perfil normal, compacto ou denso conforme a
área cliente disponível. Cada página reduz escala e espaçamentos de modo
coordenado, mantendo todas as ações visíveis sem alterar o fluxo de navegação.
O redimensionamento é tratado com debounce para evitar reconstruções excessivas.

## Fluxo do processamento

1. O CSV é validado, carregado e tem seu cabeçalho lido da primeira linha.
2. O filtro de departamento é aplicado.
3. Cada linha vira um `RegistroFuncionario`, mantendo horas em minutos.
4. A matrícula completa é preservada como texto e o CPF é removido da saída.
5. Os totais da aba principal são calculados pelos serviços existentes.
6. Os geradores criam `SALDO`, `RANKING` e `RESUMO` sem recalcular horas.
7. A aba principal é formatada e o workbook é salvo em XLSX.

Na entrada direta pelo RHiD, a integração solicita o `extrato` CSV agrupado por
funcionário, sem solicitar CPF, acompanha o processamento
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

A fixture anônima em `tests/fixtures/` protege cabeçalhos, filtros, totais, abas, homônimos e a remoção do CPF.

## Estado e persistência

`EstadoInterface` centraliza as transições visuais: vazio, pronto,
processando, concluído e erro. A janela traduz o estado em textos e habilitação
dos controles; o controller não altera widgets diretamente.

Preferências, histórico e auditoria ficam fora do repositório e são tratados como
efeitos auxiliares. O resultado principal é o XLSX: uma falha ao registrar
histórico ou auditoria é registrada no log, mas não transforma um arquivo salvo
com sucesso em aparente falha.

O envio ao Power BI calcula uma impressão digital SHA-256 somente dos dados
analíticos estáveis. IDs, horário de geração e nome do arquivo não participam da
comparação. Um registro local permite detectar o mesmo conteúdo em outra sessão;
o usuário ainda pode confirmar um reenvio intencional.

O preparo do snapshot não conhece o transporte. `PowerBiDestination` define a
fronteira de publicação; hoje ela é implementada por
`PushSemanticModelDestination`. Um futuro destino Lakehouse poderá receber as
mesmas linhas sem alterar cálculo, privacidade, deduplicação ou interface.

## Pontos de extensão

- uma nova origem de dados deve produzir o mesmo contrato CSV ou registros
  equivalentes e entrar pela pipeline comum;
- uma nova aba deve receber `RegistroFuncionario` e não reler a worksheet;
- uma nova tela deve ser adicionada ao roteamento de `navigation.py` e exposta
  pela fachada de `MainWindow`;
- novos dados sensíveis exigem revisão de saída, logs, histórico e temporários.
