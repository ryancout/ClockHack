# Instalador e empacotamento

O projeto inclui:

- `build_tools/gerar_exe.bat`
- `build_tools/gerar_instalador.bat`
- `build_tools/FASJornada.iss`
- `app/assets/icon.ico`
- `main.spec`

Instale primeiro as dependências de desenvolvimento:

```bat
python -m pip install -r requirements-dev.txt
```

## Como gerar o EXE

Execute:

```bat
build_tools\gerar_exe.bat
```

Apos a geracao, o executavel sera criado em:

```bat
dist\FASJornada.exe
```

## Como gerar o instalador

1. Instale o Inno Setup 6.
2. Gere o EXE primeiro com:

```bat
build_tools\gerar_exe.bat
```

3. Depois execute:

```bat
build_tools\gerar_instalador.bat
```

O instalador padrao do Windows sera gerado em:

```bat
dist\FASJornada_Setup_X.X.X.exe
```

## Release completa

Execute:

```powershell
powershell -ExecutionPolicy Bypass -File build_tools\build_release.ps1
```

O atalho `build_tools\build_release_auto_version.bat` chama o mesmo processo.
Ele não edita versões e não executa comandos Git. Antes do empacotamento, o
script confirma que `app/core/version.py`, `version_info.txt` e
`build_tools/FASJornada.iss` estão sincronizados e executa toda a validação.

Os arquivos são produzidos em `dist/`:

- `FASJornada_Setup_X.X.X.exe`;
- `FASJornada_vX.X.X_portable.zip`;
- `SHA256SUMS.txt`.

### Assinatura opcional

Para assinar localmente com Authenticode, instale o Windows SDK e defina:

```powershell
$env:FAS_SIGN_CERTIFICATE_PATH = 'C:\caminho\certificado.pfx'
$env:FAS_SIGN_CERTIFICATE_PASSWORD = 'senha-do-certificado'
```

A senha não deve entrar no repositório. Sem certificado, o build é concluído
com um aviso e os binários permanecem sem assinatura.

### Release automática

O workflow `.github/workflows/release.yml` é acionado por uma tag `vX.X` ou
`vX.X.X`. Ele valida a correspondência da tag, executa testes, cria instalador,
ZIP e hashes, e publica uma nova release sem substituir uma existente.

No GitHub, a assinatura é ativada pelos secrets
`WINDOWS_CERTIFICATE_BASE64` e `WINDOWS_CERTIFICATE_PASSWORD`. Nunca inclua
`data`, `logs`, relatórios reais, `.git`, caches ou credenciais no pacote.
