"""Perfis puros de densidade para a interface sem barras de rolagem."""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum


class LayoutDensity(str, Enum):
    DENSE = "dense"
    COMPACT = "compact"
    NORMAL = "normal"


@dataclass(frozen=True, slots=True)
class LayoutProfile:
    density: LayoutDensity
    widget_scaling: float
    header_padding: int
    content_padding: int
    footer_height: int


DENSE_PROFILE = LayoutProfile(LayoutDensity.DENSE, 0.78, 8, 6, 34)
COMPACT_PROFILE = LayoutProfile(LayoutDensity.COMPACT, 0.88, 14, 10, 48)
NORMAL_PROFILE = LayoutProfile(LayoutDensity.NORMAL, 1.0, 24, 18, 68)


def choose_layout_profile(width: int, height: int) -> LayoutProfile:
    """Escolhe densidade pela área cliente disponível, sem alterar navegação."""

    safe_width = max(1, int(width))
    safe_height = max(1, int(height))
    if safe_width < 900 or safe_height < 680:
        return DENSE_PROFILE
    if safe_width < 1500 or safe_height < 900:
        return COMPACT_PROFILE
    return NORMAL_PROFILE


__all__ = [
    "COMPACT_PROFILE",
    "DENSE_PROFILE",
    "LayoutDensity",
    "LayoutProfile",
    "NORMAL_PROFILE",
    "choose_layout_profile",
]
