"""Registro local para evitar reenvios acidentais ao Power BI."""

from __future__ import annotations

import hashlib
import json
import os
import tempfile
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from app.core.config import POWERBI_SENDS_FILE
from app.services.analytics_service import PowerBiSnapshot


_CAMPOS_VARIAVEIS = frozenset({"IDRelatorio", "GeradoEm", "Arquivo"})
_MAX_REGISTROS = 500


@dataclass(frozen=True, slots=True)
class PowerBiSendRecord:
    fingerprint: str
    report_id: str
    dataset_id: str
    sent_at: str
    source_file: str
    row_count: int


def calcular_fingerprint_snapshot(snapshot: PowerBiSnapshot) -> str:
    """Identifica o conteúdo analítico, ignorando IDs e horários de execução."""

    linhas = []
    for linha in snapshot.rows:
        estavel = {
            chave: valor
            for chave, valor in linha.items()
            if chave not in _CAMPOS_VARIAVEIS
        }
        linhas.append(estavel)
    linhas.sort(
        key=lambda item: json.dumps(
            item,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
            default=str,
        )
    )
    conteudo = json.dumps(
        linhas,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
        default=str,
    ).encode("utf-8")
    return hashlib.sha256(conteudo).hexdigest()


class PowerBiSendRegistry:
    def __init__(self, path: str | Path = POWERBI_SENDS_FILE) -> None:
        self.path = Path(path)

    def find(self, fingerprint: str) -> PowerBiSendRecord | None:
        for registro in self._load():
            if registro.fingerprint == fingerprint:
                return registro
        return None

    def register(
        self,
        *,
        fingerprint: str,
        report_id: str,
        dataset_id: str,
        source_file: str,
        row_count: int,
        sent_at: datetime | None = None,
    ) -> PowerBiSendRecord:
        registro = PowerBiSendRecord(
            fingerprint=fingerprint,
            report_id=report_id,
            dataset_id=dataset_id,
            sent_at=(sent_at or datetime.now().astimezone()).isoformat(),
            source_file=Path(source_file).name,
            row_count=int(row_count),
        )
        dados = [
            asdict(item)
            for item in self._load()
            if item.fingerprint != fingerprint
        ]
        dados.insert(0, asdict(registro))
        self._atomic_write(dados[:_MAX_REGISTROS])
        return registro

    def _load(self) -> list[PowerBiSendRecord]:
        if not self.path.exists():
            return []
        try:
            bruto: Any = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, UnicodeDecodeError, json.JSONDecodeError):
            return []
        if not isinstance(bruto, list):
            return []

        registros = []
        for item in bruto:
            if not isinstance(item, dict):
                continue
            try:
                registros.append(
                    PowerBiSendRecord(
                        fingerprint=str(item["fingerprint"]),
                        report_id=str(item["report_id"]),
                        dataset_id=str(item["dataset_id"]),
                        sent_at=str(item["sent_at"]),
                        source_file=Path(str(item["source_file"])).name,
                        row_count=int(item["row_count"]),
                    )
                )
            except (KeyError, TypeError, ValueError):
                continue
        return registros

    def _atomic_write(self, data: list[dict[str, Any]]) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        fd, temp_path = tempfile.mkstemp(
            prefix=f"{self.path.stem}_",
            suffix=".tmp",
            dir=str(self.path.parent),
        )
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as arquivo:
                json.dump(data, arquivo, ensure_ascii=False, indent=2)
                arquivo.flush()
                os.fsync(arquivo.fileno())
            os.replace(temp_path, self.path)
        finally:
            if os.path.exists(temp_path):
                os.remove(temp_path)


__all__ = [
    "PowerBiSendRecord",
    "PowerBiSendRegistry",
    "calcular_fingerprint_snapshot",
]
