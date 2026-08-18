# Histórico de versões

## 8.2 — 2026-08-18

- diagnóstico somente leitura das conexões RHiD, Microsoft, workspace e modelo
- isolamento do destino Power BI para permitir a futura migração do modelo Push
  para Fabric Lakehouse/OneLake sem alterar a preparação dos indicadores
- automação de CI/release, hashes e assinatura Authenticode opcional
- perfis visuais responsivos sem barras de rolagem

## 8.1 — 2026-08-18

- identificação de homônimos pela matrícula completa
- remoção do CPF dos arquivos gerados e da solicitação feita ao RHiD
- limpeza de arquivos e camadas antigas sem alteração das regras de cálculo
- documentação técnica completa de arquitetura, módulos, fluxos, integração,
  segurança, testes e distribuição
- separação entre dependências de execução e de desenvolvimento

## 8.0 — 2026-08-18

### Principais novidades

- nova identidade do produto: **FAS Jornada**
- interface redesenhada, responsiva e organizada em um único fluxo de navegação
- geração direta de relatórios pela integração com o RHiD
- escolha de empresa, múltiplos setores ou todos os setores com funcionários ativos
- período em formato brasileiro, com calendário compacto
- opção segura de lembrar o acesso RHiD pelo Gerenciador de Credenciais do Windows ou DPAPI
- fluxo alternativo por arquivos CSV mantido com a mesma importância
- tela de processamento em segundo plano, progresso e cancelamento cooperativo
- tela de conclusão com acesso ao arquivo, pasta de saída e geração de novo relatório

### Relatórios e segurança

- preservação do cálculo existente de Banco Total e Banco Saldo
- abas SALDO, RESUMO e RANKING selecionadas por padrão
- validação do contrato CSV retornado pelo RHiD antes do processamento
- gravação transacional dos arquivos para reduzir risco de saída incompleta
- credenciais nunca são armazenadas em JSON ou texto puro

### Arquitetura

- separação entre domínio, integrações, relatórios, serviços, controlador e interface
- processamento pesado retirado da thread da interface
- suíte de caracterização para proteger os cálculos durante refatorações

## 6.4

- inclusão da aba SALDO
- ajustes nas abas RANKING e RESUMO
- melhorias visuais na planilha principal e no processo de build
