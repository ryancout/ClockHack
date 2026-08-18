# Segurança e privacidade

## Dados presentes no Excel

- o Excel gerado mantém nome, matrícula completa, setor e informações de jornada;
- CPF é removido da aba principal e nunca é inserido nas abas auxiliares;
- a integração RHiD não solicita `Person.cpf`;
- matrícula é necessária para distinguir homônimos e, portanto, continua sendo
  dado pessoal que exige controle de acesso ao arquivo final.

## Credenciais RHiD

Salvar credenciais é opcional. No Windows, a preferência é o Gerenciador de
Credenciais. Se ele não estiver acessível, o fallback é um arquivo binário cifrado
pela DPAPI para o usuário do Windows. Senha nunca é gravada em JSON ou texto puro,
e é omitida de representações de objetos.

Sem “lembrar acesso”, a senha permanece apenas na memória da sessão. O repositório
não contém credenciais, tokens nem arquivos `.env` necessários ao funcionamento.

## Arquivos locais

Dados da aplicação ficam sob o perfil do usuário, no diretório legado
`ProcessadorPlanilhasFAS`, preservado para compatibilidade:

- `data/preferences.json`;
- `data/history.json`;
- `data/audit.json`;
- `data/powerbi_sends.json`, com hashes e metadados técnicos, sem conteúdo de jornada;
- `logs/app.log` e rotações;
- credencial DPAPI, somente quando necessária.

Preferências usam escrita atômica. CSVs temporários do RHiD são criados pelo
sistema operacional, processados e removidos em bloco `finally`. Falha de remoção
é registrada sem expor o conteúdo.

## Rede e distribuição

- a integração usa HTTPS com os serviços oficiais configurados no cliente;
- o login do Power BI ocorre no navegador pela Microsoft Identity Platform; a
  senha Microsoft não passa pelo FAS Jornada e o token fica somente em memória;
- o modelo analítico recebe nome, matrícula e dados de jornada, mas nunca CPF;
- IDs de aplicativo, diretório e workspace não são segredos e não concedem
  acesso sem uma conta corporativa autorizada;
- não existe servidor próprio nem banco de dados do FAS Jornada;
- builds e instaladores não incluem logs, dados locais, `.git`, testes ou caches;
- releases sem assinatura digital comercial podem gerar alerta do Windows.
- dependências reutilizáveis dos workflows GitHub Actions são fixadas pelo hash
  completo do commit, reduzindo risco de troca indireta da ação;
- cada release inclui `SHA256SUMS.txt` para verificação dos downloads;
- a assinatura Authenticode é opcional e usa secrets do GitHub somente nas
  etapas de assinatura; o certificado não é armazenado no repositório.

## Checklist antes de publicar

1. Executar testes e análise estática.
2. Confirmar ausência de credenciais, relatórios reais e dados pessoais no diff.
3. Gerar artefatos a partir de commit limpo e versionado.
4. Publicar hashes SHA-256 junto dos downloads.
5. Não substituir silenciosamente artefatos de uma tag já divulgada.
6. Quando houver certificado corporativo, confirmar a assinatura Authenticode do
   executável e do instalador antes da publicação.
