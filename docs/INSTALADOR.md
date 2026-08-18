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

`build_tools\build_release_auto_version.bat` sincroniza versão, gera o EXE e o
ZIP portátil e prepara o envio do código. Antes de criar tag/release, valide os
artefatos e gere `SHA256SUMS.txt`. Nunca inclua `data`, `logs`, relatórios reais,
`.git`, caches ou credenciais no pacote.
