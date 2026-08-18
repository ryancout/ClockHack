# Histórico de versões

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
- identificação de homônimos pelos três últimos dígitos da matrícula
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
