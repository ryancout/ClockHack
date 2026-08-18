from pathlib import Path

import pytest

from build_tools.verify_release_version import (
    read_app_version,
    validate_release_metadata,
    version_tuple,
)


def _release_tree(tmp_path: Path, version: str = "8.2") -> Path:
    (tmp_path / "app/core").mkdir(parents=True)
    (tmp_path / "build_tools").mkdir()
    (tmp_path / "app/core/version.py").write_text(
        f'APP_NAME = "FAS Jornada"\nAPP_VERSION = "{version}"\n',
        encoding="utf-8",
    )
    numeric = ",".join(str(item) for item in version_tuple(version))
    (tmp_path / "version_info.txt").write_text(
        f"filevers=({numeric})\n"
        f"prodvers=({numeric})\n"
        f"StringStruct('FileVersion', '{version}')\n"
        f"StringStruct('ProductVersion', '{version}')\n",
        encoding="utf-8",
    )
    (tmp_path / "build_tools/FASJornada.iss").write_text(
        f'#define MyAppVersion "{version}"\n', encoding="utf-8"
    )
    return tmp_path


def test_metadados_consistentes_com_a_tag(tmp_path):
    root = _release_tree(tmp_path)

    assert read_app_version(root) == "8.2"
    assert validate_release_metadata(root, tag="v8.2") == "8.2"


def test_tag_divergente_impede_release(tmp_path):
    root = _release_tree(tmp_path)

    with pytest.raises(ValueError, match="não corresponde"):
        validate_release_metadata(root, tag="v8.1")


def test_metadado_windows_divergente_impede_release(tmp_path):
    root = _release_tree(tmp_path)
    (root / "version_info.txt").write_text("incompatível", encoding="utf-8")

    with pytest.raises(ValueError, match="filevers"):
        validate_release_metadata(root)
