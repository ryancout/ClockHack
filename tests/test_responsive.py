from pathlib import Path

from app.ui.responsive import LayoutDensity, choose_layout_profile


def test_densidade_normal_em_tela_grande():
    profile = choose_layout_profile(1920, 1000)

    assert profile.density is LayoutDensity.NORMAL
    assert profile.widget_scaling == 1.0


def test_densidade_compacta_no_notebook_padrao():
    profile = choose_layout_profile(1366, 768)

    assert profile.density is LayoutDensity.COMPACT
    assert profile.footer_height < 68


def test_densidade_maxima_em_area_cliente_baixa():
    profile = choose_layout_profile(820, 620)

    assert profile.density is LayoutDensity.DENSE
    assert profile.widget_scaling < 0.8


def test_interface_principal_nao_usa_frames_com_rolagem():
    ui_dir = Path(__file__).resolve().parents[1] / "app/ui"
    source = "\n".join(
        path.read_text(encoding="utf-8")
        for path in (ui_dir / "main_window.py", ui_dir / "report_pages.py", ui_dir / "rhid_page.py")
    )

    assert "CTkScrollableFrame" not in source
