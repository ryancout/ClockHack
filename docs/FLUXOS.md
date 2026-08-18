# Fluxos de processamento

## Origem CSV

1. O usuário seleciona um ou vários arquivos `.csv`.
2. O controller valida cada entrada e carrega os departamentos do primeiro arquivo.
3. Antes da thread, define destinos, detecta colisões e confirma sobrescritas.
4. Um plano imutável captura arquivos, filtro e abas opcionais.
5. O worker processa os arquivos em sequência; a interface recebe progresso por fila.
6. Cada sucesso é registrado no histórico e o último resultado fica disponível para abertura.

Contrato de entrada:

- UTF-8 com BOM aceito;
- delimitador `;`;
- cabeçalho na primeira linha;
- colunas mínimas: Nome do funcionário, Número de matrícula, Nome do departamento,
  Banco Total e Banco Saldo;
- horários `HH:MM` ou `-HH:MM`, inclusive acima de 24 horas.

## Pipeline Excel comum

```text
CSV -> validar -> carregar -> mapear colunas -> filtrar departamento
    -> converter linhas em registros/minutos -> calcular totais
    -> normalizar matrícula -> remover CPF -> criar abas opcionais
    -> formatar aba principal -> salvar XLSX
```

Regras importantes:

- a matrícula completa é mantida como texto, inclusive zeros iniciais;
- CPF pode existir no CSV manual, mas é removido antes da gravação do Excel;
- vazio nas colunas de banco mantém o comportamento histórico de zero;
- a linha TOTAL usa o mesmo parser do domínio;
- SALDO, RANKING e RESUMO recebem os mesmos registros já calculados;
- filtros e abas opcionais não alteram a função de conversão de horas.

## Origem RHiD

1. A página autentica com e-mail, senha e domínio quando necessário.
2. O cliente carrega empresas, departamentos e pessoas ativas.
3. O usuário escolhe empresa, todos ou vários setores, datas e abas.
4. O diálogo de salvamento define o XLSX antes do trabalho remoto.
5. O RHiD cria um job de relatório; o cliente acompanha seu GUID até 100%.
6. O CSV baixado é validado como consolidado: cabeçalhos exigidos e uma linha por matrícula.
7. Um temporário local entra na mesma pipeline do CSV manual e é apagado no `finally`.

## Cancelamento e falha parcial

O cancelamento é cooperativo. O arquivo que já entrou na gravação termina para não
corromper o XLSX; próximos itens não começam. Se houver sucesso antes de uma falha,
o aplicativo informa quantos arquivos foram salvos e permite abrir o último deles.
Histórico, auditoria e preferências são auxiliares: falhas neles não invalidam um
XLSX já salvo.
