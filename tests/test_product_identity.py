from pathlib import Path

from app.core import config
from app.core.version import APP_NAME, APP_VERSION


ROOT = Path(__file__).resolve().parents[1]


def test_identidade_publica_do_produto():
    assert APP_NAME == "FAS Jornada"
    assert APP_VERSION == "8.2"
    assert config.APP_TITLE == f"FAS Jornada - V{APP_VERSION}"


def test_diretorio_de_dados_legado_e_preservado(monkeypatch, tmp_path):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    monkeypatch.delenv("APPDATA", raising=False)

    assert config._user_data_base_dir() == tmp_path / "ProcessadorPlanilhasFAS"


def test_metadados_e_instalador_usam_nova_identidade():
    spec = (ROOT / "main.spec").read_text(encoding="utf-8")
    versao = (ROOT / "version_info.txt").read_text(encoding="utf-8")
    instalador = (ROOT / "build_tools" / "FASJornada.iss").read_text(
        encoding="utf-8"
    )

    assert "name='FASJornada'" in spec
    assert "StringStruct('ProductName', 'FAS Jornada')" in versao
    assert "StringStruct('FileDescription', 'Relatório e Análise de Jornada')" in versao
    assert "StringStruct('InternalName', 'FASJornada')" in versao
    assert "StringStruct('ProductVersion', '8.2')" in versao
    assert "#define MyAppExeName \"FASJornada.exe\"" in instalador
    assert "#define MyAppVersion \"8.2\"" in instalador
    assert "AppId={{A8E7A9B6-9C32-49D2-A0C9-7E4C11223344}" in instalador
