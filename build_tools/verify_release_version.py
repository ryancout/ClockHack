"""Valida que todos os metadados publicados usam a mesma versão."""

from __future__ import annotations

import argparse
import re
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
VERSION_PATTERN = re.compile(r'^APP_VERSION\s*=\s*["\']([^"\']+)["\']', re.MULTILINE)


def read_app_version(root: Path = ROOT) -> str:
    content = (root / "app/core/version.py").read_text(encoding="utf-8")
    match = VERSION_PATTERN.search(content)
    if match is None:
        raise ValueError("APP_VERSION não foi encontrado em app/core/version.py.")
    version = match.group(1)
    if not re.fullmatch(r"\d+\.\d+(?:\.\d+){0,2}", version):
        raise ValueError(f"Versão inválida: {version!r}.")
    return version


def version_tuple(version: str) -> tuple[int, int, int, int]:
    parts = [int(item) for item in version.split(".")]
    return tuple((parts + [0, 0, 0, 0])[:4])  # type: ignore[return-value]


def validate_release_metadata(root: Path = ROOT, tag: str | None = None) -> str:
    version = read_app_version(root)
    if tag is not None and tag != f"v{version}":
        raise ValueError(
            f"A tag {tag!r} não corresponde à versão v{version} do aplicativo."
        )

    info = (root / "version_info.txt").read_text(encoding="utf-8")
    expected_tuple = ",".join(str(item) for item in version_tuple(version))
    compact_info = re.sub(r"\s+", "", info)
    for field in ("filevers", "prodvers"):
        if f"{field}=({expected_tuple})" not in compact_info:
            raise ValueError(f"{field} não corresponde a {version}.")
    for field in ("FileVersion", "ProductVersion"):
        if not re.search(
            rf"StringStruct\(['\"]{field}['\"],\s*['\"]{re.escape(version)}['\"]\)",
            info,
        ):
            raise ValueError(f"{field} não corresponde a {version}.")

    installer = (root / "build_tools/FASJornada.iss").read_text(encoding="utf-8")
    if f'#define MyAppVersion "{version}"' not in installer:
        raise ValueError("A versão do instalador não corresponde ao aplicativo.")
    return version


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--tag")
    args = parser.parse_args()
    version = validate_release_metadata(tag=args.tag)
    print(f"Metadados consistentes: FAS Jornada v{version}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
