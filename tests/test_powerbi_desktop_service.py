import json

import pytest

from app.services.powerbi_desktop_service import (
    PowerBiDesktopError,
    criar_relatorio_powerbi_desktop,
)


DATASET_ID = "59c84d55-3ba5-4244-947e-8f5f1189f8f5"


def test_cria_relatorio_fino_conectado_sem_credenciais(tmp_path):
    pbir = criar_relatorio_powerbi_desktop(DATASET_ID, pasta_base=tmp_path)

    definicao = json.loads(pbir.read_text(encoding="utf-8"))
    conexao = definicao["datasetReference"]["byConnection"]["connectionString"]
    assert f"semanticmodelid={DATASET_ID}" in conexao
    assert "FAS Jornada Analytics v2" in conexao
    assert "ClaimsToken" in conexao
    assert "access_token" not in pbir.read_text(encoding="utf-8").lower()

    definition_dir = pbir.parent / "definition"
    assert (definition_dir / "report.json").is_file()
    assert (definition_dir / "version.json").is_file()
    assert (definition_dir / "pages" / "pages.json").is_file()
    tema = (
        pbir.parent
        / "StaticResources"
        / "SharedResources"
        / "BaseThemes"
        / "CY24SU10.json"
    )
    assert json.loads(tema.read_text(encoding="utf-8"))["name"] == "CY24SU10"
    report = json.loads((definition_dir / "report.json").read_text(encoding="utf-8"))
    assert report["themeCollection"]["baseTheme"]["name"] == "CY24SU10"
    paginas = json.loads(
        (definition_dir / "pages" / "pages.json").read_text(encoding="utf-8")
    )
    pagina_id = paginas["activePageName"]
    pagina = json.loads(
        (definition_dir / "pages" / pagina_id / "page.json").read_text(
            encoding="utf-8"
        )
    )
    assert pagina["displayName"] == "Visão geral"
    assert pagina["displayOption"] == "FitToPage"


def test_rejeita_id_de_modelo_invalido(tmp_path):
    with pytest.raises(PowerBiDesktopError, match="identificador"):
        criar_relatorio_powerbi_desktop("id-inválido", pasta_base=tmp_path)


def test_rejeita_nome_que_poderia_alterar_connection_string(tmp_path):
    with pytest.raises(PowerBiDesktopError, match="workspace"):
        criar_relatorio_powerbi_desktop(
            DATASET_ID,
            pasta_base=tmp_path,
            workspace_name="Workspace;Password=indevido",
        )
