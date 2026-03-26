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

# ClockHack
Advanced workforce time analyzer that processes Excel/CSV files, calculates balances, highlights anomalies, and generates automated insights like rankings and department analytics.

## TimeForge is a desktop application designed to transform raw workforce time data into actionable insights.

It processes Excel and CSV files exported from time tracking systems, automatically calculates time balances, highlights critical situations (such as excessive overtime or negative balances), and generates structured outputs ready for analysis and reporting.

Key features include:

Batch processing of multiple files
Department-based filtering before processing
Automatic calculation of "Banco de Horas"
Visual highlighting of anomalies:
Negative balances (debt) in red
Excessive overtime in yellow
Generation of a final Excel file preserving original formatting
Automatic creation of:
Ranking sheet (top debtors and top overtime)
Summary sheet with department aggregation
Built-in charts for quick visualization
Clean and intuitive desktop interface (no browser required)
This tool bridges the gap between raw time tracking exports and real workforce analytics, eliminating manual Excel work and enabling faster, more reliable decision-making.
