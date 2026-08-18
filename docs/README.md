# Documentação técnica

Este diretório descreve o código do FAS Jornada sem duplicar os arquivos-fonte.
Os documentos refletem o estado atual da branch de desenvolvimento.

- [Arquitetura](ARQUITETURA.md): camadas, dependências e concorrência.
- [Referência dos módulos](REFERENCIA_MODULOS.md): responsabilidade de cada arquivo Python.
- [Fluxos de processamento](FLUXOS.md): CSV, RHiD, cálculo e saída Excel.
- [Integração RHiD](INTEGRACAO_RHID.md): autenticação, catálogo, geração e download.
- [Segurança e privacidade](SEGURANCA_E_PRIVACIDADE.md): credenciais, CPF, temporários e dados locais.
- [Desenvolvimento e testes](DESENVOLVIMENTO.md): ambiente, comandos e estratégia de testes.
- [Instalador e empacotamento](INSTALADOR.md): EXE, instalador e artefatos.

## Regra de atualização

Toda mudança de contrato, módulo, endpoint, fluxo ou build deve atualizar o
documento correspondente no mesmo commit. A documentação de versões antigas não
é mantida como instrução ativa; o histórico permanece no Git e no `CHANGELOG.md`.
