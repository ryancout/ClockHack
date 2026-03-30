# Instalador e empacotamento

O projeto inclui:

- `build_tools/gerar_exe.bat`
- `build_tools/gerar_instalador.bat`
- `build_tools/ProcessadorPlanilhasFAS.iss`
- `app/assets/icon.ico`
- `main.spec`

## Como gerar o EXE

Execute:

```bat
build_tools\gerar_exe.bat
```

Apos a geracao, o executavel sera criado em:

```bat
dist\ProcessadorPlanilhasFAS.exe
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
dist\ProcessadorPlanilhasFAS_Setup.exe
```
