"""Cria e abre um relatório Power BI Desktop ligado ao modelo publicado."""

from __future__ import annotations

import json
import os
from collections.abc import Callable
from pathlib import Path
from uuid import UUID

from app.core.config import (
    ASSETS_DIR,
    DATA_DIR,
    POWER_BI_DATASET_NAME,
    POWER_BI_WORKSPACE_NAME,
)


class PowerBiDesktopError(Exception):
    """Falha ao preparar ou abrir o relatório no Power BI Desktop."""


_PAGE_ID = "b8c5fb8d635f898326c6"
_PROJECT_NAME = POWER_BI_DATASET_NAME


def _validar_id_modelo(dataset_id: object) -> str:
    try:
        return str(UUID(str(dataset_id).strip()))
    except (AttributeError, TypeError, ValueError) as erro:
        raise PowerBiDesktopError("O identificador do modelo Power BI é inválido.") from erro


def _validar_nome(valor: object, campo: str) -> str:
    texto = str(valor).strip()
    if not texto or any(caractere in texto for caractere in ('"', ";", "\r", "\n")):
        raise PowerBiDesktopError(f"O nome de {campo} do Power BI é inválido.")
    return texto


def _salvar_json(caminho: Path, conteudo: dict) -> None:
    caminho.parent.mkdir(parents=True, exist_ok=True)
    temporario = caminho.with_suffix(caminho.suffix + ".tmp")
    try:
        temporario.write_text(
            json.dumps(conteudo, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        os.replace(temporario, caminho)
    except OSError as erro:
        raise PowerBiDesktopError(
            "Não foi possível criar o atalho para o Power BI Desktop."
        ) from erro


def _carregar_tema_padrao() -> dict:
    caminho = ASSETS_DIR / "powerbi" / "CY24SU10.json"
    try:
        conteudo = json.loads(caminho.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as erro:
        raise PowerBiDesktopError(
            "O tema padrão do relatório Power BI não está disponível."
        ) from erro
    if not isinstance(conteudo, dict):
        raise PowerBiDesktopError("O tema padrão do Power BI é inválido.")
    return conteudo


def criar_relatorio_powerbi_desktop(
    dataset_id: object,
    *,
    pasta_base: str | os.PathLike[str] | None = None,
    workspace_name: str = POWER_BI_WORKSPACE_NAME,
    dataset_name: str = POWER_BI_DATASET_NAME,
) -> Path:
    """Gera um PBIR fino, sem dados locais nem credenciais gravadas."""

    modelo_id = _validar_id_modelo(dataset_id)
    workspace = _validar_nome(workspace_name, "workspace")
    modelo = _validar_nome(dataset_name, "modelo")
    raiz = Path(pasta_base) if pasta_base is not None else DATA_DIR / "powerbi"
    report_dir = raiz / f"{_PROJECT_NAME}.Report"
    definition_dir = report_dir / "definition"
    page_dir = definition_dir / "pages" / _PAGE_ID
    theme_path = (
        report_dir
        / "StaticResources"
        / "SharedResources"
        / "BaseThemes"
        / "CY24SU10.json"
    )

    conexao = (
        'Data Source="powerbi://api.powerbi.com/v1.0/myorg/'
        f'{workspace}";initial catalog={modelo};access mode=readonly;'
        f'integrated security=ClaimsToken;semanticmodelid={modelo_id}'
    )
    _salvar_json(
        report_dir / "definition.pbir",
        {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
            "version": "4.0",
            "datasetReference": {
                "byConnection": {"connectionString": conexao}
            },
        },
    )
    _salvar_json(
        definition_dir / "version.json",
        {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/versionMetadata/1.0.0/schema.json",
            "version": "2.0.0",
        },
    )
    _salvar_json(
        definition_dir / "report.json",
        {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/1.2.0/schema.json",
            "themeCollection": {
                "baseTheme": {
                    "name": "CY24SU10",
                    "reportVersionAtImport": "5.61",
                    "type": "SharedResources",
                }
            },
            "layoutOptimization": "None",
            "objects": {
                "section": [
                    {
                        "properties": {
                            "verticalAlignment": {
                                "expr": {"Literal": {"Value": "'Top'"}}
                            }
                        }
                    }
                ]
            },
            "resourcePackages": [
                {
                    "name": "SharedResources",
                    "type": "SharedResources",
                    "items": [
                        {
                            "name": "CY24SU10",
                            "path": "BaseThemes/CY24SU10.json",
                            "type": "BaseTheme",
                        }
                    ],
                }
            ],
            "settings": {
                "useStylableVisualContainerHeader": True,
                "defaultDrillFilterOtherVisuals": True,
                "allowChangeFilterTypes": True,
                "useEnhancedTooltips": True,
                "useDefaultAggregateDisplayName": True,
            },
        },
    )
    _salvar_json(theme_path, _carregar_tema_padrao())
    _salvar_json(
        definition_dir / "pages" / "pages.json",
        {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.0.0/schema.json",
            "pageOrder": [_PAGE_ID],
            "activePageName": _PAGE_ID,
        },
    )
    _salvar_json(
        page_dir / "page.json",
        {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/1.3.0/schema.json",
            "name": _PAGE_ID,
            "displayName": "Visão geral",
            "displayOption": "FitToPage",
            "height": 720,
            "width": 1280,
        },
    )
    return report_dir / "definition.pbir"


def abrir_relatorio_powerbi_desktop(
    dataset_id: object,
    *,
    abridor: Callable[[str], object] | None = None,
) -> Path:
    """Cria o relatório conectado e solicita sua abertura no Desktop."""

    caminho = criar_relatorio_powerbi_desktop(dataset_id)
    abrir = abridor or getattr(os, "startfile", None)
    if abrir is None:
        raise PowerBiDesktopError(
            "A abertura automática do Power BI Desktop só está disponível no Windows."
        )
    try:
        abrir(str(caminho))
    except OSError as erro:
        raise PowerBiDesktopError(
            "Não foi possível abrir o Power BI Desktop. Verifique se ele está instalado."
        ) from erro
    return caminho


__all__ = [
    "PowerBiDesktopError",
    "abrir_relatorio_powerbi_desktop",
    "criar_relatorio_powerbi_desktop",
]
