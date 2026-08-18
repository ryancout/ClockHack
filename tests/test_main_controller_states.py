from pathlib import Path

import pytest

from app.controllers import main_controller as controller_module
from app.core.exceptions import AppError
from app.integrations.rhid_client import RhidCompany, RhidDepartment
from app.services.analytics_service import PowerBiSnapshot
from app.ui.view_state import EstadoInterface


class FakeView:
    def __init__(self):
        self.estados = []
        self.status = []
        self.arquivos = []
        self.departamentos = []
        self.metricas = []
        self.progressos = []
        self.pastas_saida = []
        self.tempos = []
        self.historicos = []
        self.botao_abrir_habilitado = []
        self.botao_pasta_habilitado = []
        self.versao_atualizada = False
        self.rhid_ocupado = []
        self.rhid_progressos = []
        self.rhid_sucessos = []
        self.rhid_erros = []
        self.powerbi_ocupado = []
        self.powerbi_progressos = []
        self.powerbi_sucessos = []
        self.powerbi_erros = []

    def definir_estado(self, estado, total_arquivos=0):
        self.estados.append((estado, total_arquivos))

    def atualizar_status(self, mensagem, tipo):
        self.status.append((mensagem, tipo))

    def atualizar_arquivo(self, texto):
        self.arquivos.append(texto)

    def atualizar_departamentos(self, departamentos, selecionado="Todos"):
        self.departamentos.append((departamentos, selecionado))

    def atualizar_metricas(self, *metricas):
        self.metricas.append(metricas)

    def atualizar_progresso(self, progresso):
        self.progressos.append(progresso)

    def atualizar_pasta_saida(self, pasta):
        self.pastas_saida.append(pasta)

    def atualizar_tempo_execucao(self, tempo):
        self.tempos.append(tempo)

    def atualizar_versao(self):
        self.versao_atualizada = True

    def habilitar_botao_abrir(self, habilitado):
        self.botao_abrir_habilitado.append(habilitado)

    def habilitar_botao_abrir_pasta(self, habilitado):
        self.botao_pasta_habilitado.append(habilitado)

    def renderizar_historico(self, historico):
        self.historicos.append(historico)

    def agendar_na_interface(self, _atraso, callback):
        callback()

    def definir_geracao_rhid_ocupada(self, ocupado):
        self.rhid_ocupado.append(ocupado)

    def atualizar_progresso_rhid(self, valor, mensagem=""):
        self.rhid_progressos.append((valor, mensagem))

    def exibir_sucesso_rhid(self, mensagem):
        self.rhid_sucessos.append(mensagem)

    def exibir_erro_rhid(self, mensagem):
        self.rhid_erros.append(mensagem)

    def definir_powerbi_ocupado(self, ocupado):
        self.powerbi_ocupado.append(ocupado)

    def definir_powerbi_enviado(self, enviado):
        self.powerbi_enviado = bool(enviado)

    def atualizar_progresso_powerbi(self, valor, mensagem=""):
        self.powerbi_progressos.append((valor, mensagem))

    def exibir_sucesso_powerbi(self, mensagem):
        self.powerbi_sucessos.append(mensagem)

    def exibir_erro_powerbi(self, mensagem):
        self.powerbi_erros.append(mensagem)


class TaskRunnerImediato:
    def __init__(self):
        self.ativo = False
        self.cancelamento_solicitado = False

    def cancelar(self):
        return False

    def executar(self, trabalho, ao_progresso, ao_concluir, ao_erro):
        if self.ativo:
            return False

        self.ativo = True
        try:
            resultado = trabalho(ao_progresso)
        except Exception as erro:
            self.ativo = False
            ao_erro(erro)
        else:
            self.ativo = False
            ao_concluir(resultado)
        return True


class TaskRunnerPendente:
    def __init__(self):
        self.ativo = False
        self.cancelamento_solicitado = False
        self._trabalho = None
        self._ao_progresso = None
        self._ao_concluir = None
        self._ao_erro = None

    def executar(self, trabalho, ao_progresso, ao_concluir, ao_erro):
        if self.ativo:
            return False
        self.ativo = True
        self._trabalho = trabalho
        self._ao_progresso = ao_progresso
        self._ao_concluir = ao_concluir
        self._ao_erro = ao_erro
        return True

    def cancelar(self):
        if not self.ativo:
            return False
        self.cancelamento_solicitado = True
        return True

    def concluir(self):
        try:
            resultado = self._trabalho(self._ao_progresso)
        except Exception as erro:
            self.ativo = False
            self._ao_erro(erro)
        else:
            self.ativo = False
            self._ao_concluir(resultado)


@pytest.fixture
def controller_isolado(monkeypatch, tmp_path):
    preferencias = {
        "last_open_dir": "",
        "last_save_dir": "",
        "last_department": "Todos",
    }
    eventos = []
    historicos = []

    monkeypatch.setattr(controller_module, "carregar_preferencias", lambda: preferencias.copy())
    monkeypatch.setattr(controller_module, "salvar_preferencias", lambda _preferencias: None)
    monkeypatch.setattr(controller_module, "ultimos_processamentos", lambda: historicos.copy())
    monkeypatch.setattr(controller_module, "registrar_historico", historicos.append)
    monkeypatch.setattr(
        controller_module,
        "registrar_evento",
        lambda nome, dados: eventos.append((nome, dados)),
    )
    monkeypatch.setattr(controller_module.messagebox, "showinfo", lambda *args, **kwargs: None)
    monkeypatch.setattr(controller_module.messagebox, "showwarning", lambda *args, **kwargs: None)
    monkeypatch.setattr(controller_module.messagebox, "showerror", lambda *args, **kwargs: None)
    monkeypatch.setattr(controller_module.messagebox, "askyesno", lambda *args, **kwargs: False)

    view = FakeView()
    controller = controller_module.MainController(
        view,
        task_runner=TaskRunnerImediato(),
        powerbi_registry=controller_module.PowerBiSendRegistry(
            tmp_path / "powerbi_sends.json"
        ),
    )
    return controller, view, eventos, historicos


def test_fluxo_vazio_pronto_processando_concluido(
    monkeypatch,
    tmp_path,
    controller_isolado,
):
    controller, view, eventos, historicos = controller_isolado
    arquivo_entrada = tmp_path / "relatorio.csv"
    arquivo_saida = tmp_path / "relatorio_tratado.xlsx"

    monkeypatch.setattr(
        controller_module.filedialog,
        "askopenfilenames",
        lambda **_kwargs: (str(arquivo_entrada),),
    )
    monkeypatch.setattr(controller_module, "validar_arquivo_entrada", lambda _caminho: None)
    monkeypatch.setattr(
        controller_module,
        "obter_departamentos",
        lambda _caminho: ["Todos", "Financeiro"],
    )
    monkeypatch.setattr(
        controller_module.filedialog,
        "asksaveasfilename",
        lambda **_kwargs: str(arquivo_saida),
    )

    resultado = {
        "caminho_saida": str(arquivo_saida),
        "tipo_entrada": "CSV",
        "quantidade_funcionarios": 2,
        "banco_total": "10:30",
        "banco_saldo": "01:15",
        "departamento": "Financeiro",
        "gerou_saldo": True,
        "gerou_resumo": True,
        "gerou_ranking": True,
    }
    monkeypatch.setattr(
        controller_module,
        "processar_arquivo",
        lambda *_args, **_kwargs: resultado,
    )

    controller.iniciar()
    controller.selecionar_arquivos()
    controller.processar("Financeiro")

    assert view.estados == [
        (EstadoInterface.VAZIO, 0),
        (EstadoInterface.PRONTO, 1),
        (EstadoInterface.PROCESSANDO, 1),
        (EstadoInterface.CONCLUIDO, 1),
    ]
    assert view.arquivos[-1] == "1 arquivo selecionado: relatorio.csv"
    assert controller.ultimo_resultado == resultado
    assert view.metricas[-1] == (2, "10:30", "1:15")
    assert view.progressos[-1] == 1.0
    assert view.botao_abrir_habilitado[-1] is True
    assert view.botao_pasta_habilitado[-1] is True
    assert view.status[-1][1] == "success"
    assert "1 arquivo foi salvo" in view.status[-1][0]
    assert len(historicos) == 1
    assert [nome for nome, _dados in eventos] == [
        "arquivos_selecionados",
        "processamento_lote",
    ]


def test_falha_no_processamento_transiciona_de_processando_para_erro(
    monkeypatch,
    tmp_path,
    controller_isolado,
):
    controller, view, eventos, historicos = controller_isolado
    arquivo_entrada = tmp_path / "relatorio.csv"
    arquivo_saida = tmp_path / "relatorio_tratado.xlsx"
    controller.arquivos_selecionados = [str(arquivo_entrada)]

    monkeypatch.setattr(
        controller_module.filedialog,
        "asksaveasfilename",
        lambda **_kwargs: str(arquivo_saida),
    )

    def falhar_processamento(*_args, **_kwargs):
        raise AppError("Arquivo de teste inválido")

    monkeypatch.setattr(controller_module, "processar_arquivo", falhar_processamento)

    controller.processar("Todos")

    assert view.estados == [
        (EstadoInterface.PROCESSANDO, 1),
        (EstadoInterface.ERRO, 1),
    ]
    assert controller.ultimo_resultado is None
    assert view.status[-1] == ("Arquivo de teste inválido", "error")
    assert historicos == []
    assert [nome for nome, _dados in eventos] == ["arquivo_ignorado"]


def test_limpar_selecao_retorna_ao_estado_vazio(controller_isolado):
    controller, view, _eventos, _historicos = controller_isolado
    controller.arquivos_selecionados = [str(Path("relatorio.csv"))]
    controller.ultimo_resultado = {"caminho_saida": "relatorio_tratado.xlsx"}

    controller.limpar_selecao()

    assert controller.arquivos_selecionados == []
    assert controller.ultimo_resultado is None
    assert view.estados == [(EstadoInterface.VAZIO, 0)]
    assert view.arquivos[-1] == "Nenhum arquivo selecionado"
    assert view.botao_abrir_habilitado[-1] is False
    assert view.botao_pasta_habilitado[-1] is False


def test_processamento_e_despachado_e_bloqueia_reentrada(
    monkeypatch,
    tmp_path,
    controller_isolado,
):
    controller, view, _eventos, _historicos = controller_isolado
    runner = TaskRunnerPendente()
    controller._task_runner = runner
    arquivo_original = tmp_path / "relatorio.csv"
    arquivo_saida = tmp_path / "relatorio_tratado.xlsx"
    controller.arquivos_selecionados = [str(arquivo_original)]
    dialogos_salvar = []
    chamadas_pipeline = []

    def escolher_saida(**_kwargs):
        dialogos_salvar.append("aberto")
        return str(arquivo_saida)

    monkeypatch.setattr(
        controller_module.filedialog,
        "asksaveasfilename",
        escolher_saida,
    )

    def processar_em_teste(entrada, saida, departamento, **opcoes):
        chamadas_pipeline.append((entrada, saida, departamento, opcoes))
        return {
            "caminho_saida": saida,
            "tipo_entrada": "CSV",
            "quantidade_funcionarios": 1,
            "banco_total": "30:15",
            "banco_saldo": "-2:30",
            "departamento": departamento,
            "gerou_saldo": opcoes["gerar_saldo"],
            "gerou_resumo": opcoes["gerar_resumo"],
            "gerou_ranking": opcoes["gerar_ranking"],
        }

    monkeypatch.setattr(controller_module, "processar_arquivo", processar_em_teste)

    controller.processar(
        "Todos",
        gerar_saldo=False,
        gerar_resumo=True,
        gerar_ranking=False,
    )

    assert runner.ativo is True
    assert controller.processamento_em_andamento is True
    assert chamadas_pipeline == []
    assert view.estados[-1] == (EstadoInterface.PROCESSANDO, 1)

    controller.processar("Todos")
    assert dialogos_salvar == ["aberto"]
    assert view.status[-1] == ("Já existe um processamento em andamento.", "warning")

    controller.arquivos_selecionados = [str(tmp_path / "outro.csv")]
    runner.concluir()

    assert controller.processamento_em_andamento is False
    assert len(chamadas_pipeline) == 1
    entrada, saida, departamento, opcoes = chamadas_pipeline[0]
    assert entrada == str(arquivo_original)
    assert saida == str(arquivo_saida)
    assert departamento == "Todos"
    assert opcoes == {
        "gerar_saldo": False,
        "gerar_ranking": False,
        "gerar_resumo": True,
    }
    assert view.estados[-1] == (EstadoInterface.CONCLUIDO, 1)
    assert view.metricas[-1] == (1, "30:15", "-2:30")


def test_colisao_de_destinos_impede_inicio_do_lote(
    monkeypatch,
    tmp_path,
    controller_isolado,
):
    controller, view, _eventos, _historicos = controller_isolado
    runner = TaskRunnerPendente()
    controller._task_runner = runner
    controller.arquivos_selecionados = [
        str(tmp_path / "origem_a" / "relatorio.csv"),
        str(tmp_path / "origem_b" / "relatorio.csv"),
    ]
    chamadas_pipeline = []

    monkeypatch.setattr(
        controller_module.filedialog,
        "askdirectory",
        lambda **_kwargs: str(tmp_path / "saida"),
    )
    monkeypatch.setattr(
        controller_module,
        "processar_arquivo",
        lambda *args, **kwargs: chamadas_pipeline.append((args, kwargs)),
    )

    controller.processar("Todos")

    assert runner.ativo is False
    assert chamadas_pipeline == []
    assert view.estados[-1] == (EstadoInterface.PRONTO, 2)
    assert view.status[-1][1] == "error"
    assert "mesma saída" in view.status[-1][0]


def test_cancelar_uma_sobrescrita_impede_todo_o_lote(
    monkeypatch,
    tmp_path,
    controller_isolado,
):
    controller, view, _eventos, _historicos = controller_isolado
    runner = TaskRunnerPendente()
    controller._task_runner = runner
    controller.arquivos_selecionados = [
        str(tmp_path / "primeiro.csv"),
        str(tmp_path / "segundo.csv"),
    ]
    pasta_saida = tmp_path / "saida"
    pasta_saida.mkdir()
    (pasta_saida / "primeiro_todos_tratado.xlsx").write_text("existente")
    (pasta_saida / "segundo_todos_tratado.xlsx").write_text("existente")
    respostas = iter([True, False])
    confirmacoes = []
    chamadas_pipeline = []

    monkeypatch.setattr(
        controller_module.filedialog,
        "askdirectory",
        lambda **_kwargs: str(pasta_saida),
    )

    def confirmar(*_args, **_kwargs):
        confirmacoes.append("perguntado")
        return next(respostas)

    monkeypatch.setattr(controller_module.messagebox, "askyesno", confirmar)
    monkeypatch.setattr(
        controller_module,
        "processar_arquivo",
        lambda *args, **kwargs: chamadas_pipeline.append((args, kwargs)),
    )

    controller.processar("Todos")

    assert confirmacoes == ["perguntado", "perguntado"]
    assert chamadas_pipeline == []
    assert runner.ativo is False
    assert view.estados[-1] == (EstadoInterface.PRONTO, 2)
    assert view.status[-1][1] == "warning"


def test_lote_assincrono_agrega_totais_e_preserva_ordem(
    monkeypatch,
    tmp_path,
    controller_isolado,
):
    controller, view, eventos, historicos = controller_isolado
    primeiro = tmp_path / "primeiro.csv"
    segundo = tmp_path / "segundo.csv"
    pasta_saida = tmp_path / "saida"
    controller.arquivos_selecionados = [str(primeiro), str(segundo)]
    chamadas = []

    monkeypatch.setattr(
        controller_module.filedialog,
        "askdirectory",
        lambda **_kwargs: str(pasta_saida),
    )

    def processar_em_teste(entrada, saida, departamento, **opcoes):
        chamadas.append((entrada, saida, departamento, opcoes))
        primeiro_arquivo = entrada == str(primeiro)
        return {
            "caminho_saida": saida,
            "tipo_entrada": "CSV",
            "quantidade_funcionarios": 3 if primeiro_arquivo else 4,
            "banco_total": "30:00" if primeiro_arquivo else "-5:30",
            "banco_saldo": "-2:00" if primeiro_arquivo else "27:45",
            "departamento": departamento,
            "gerou_saldo": opcoes["gerar_saldo"],
            "gerou_resumo": opcoes["gerar_resumo"],
            "gerou_ranking": opcoes["gerar_ranking"],
        }

    monkeypatch.setattr(controller_module, "processar_arquivo", processar_em_teste)

    controller.processar("Todos")

    assert [Path(chamada[0]).name for chamada in chamadas] == [
        "primeiro.csv",
        "segundo.csv",
    ]
    assert [Path(chamada[1]).name for chamada in chamadas] == [
        "primeiro_todos_tratado.xlsx",
        "segundo_todos_tratado.xlsx",
    ]
    assert view.estados == [
        (EstadoInterface.PROCESSANDO, 2),
        (EstadoInterface.CONCLUIDO, 2),
    ]
    assert view.metricas[-1] == (7, "24:30", "25:45")
    assert "2 arquivos foram salvos" in view.status[-1][0]
    assert len(historicos) == 2
    assert Path(controller.ultimo_resultado["caminho_saida"]).name == (
        "segundo_todos_tratado.xlsx"
    )
    assert [nome for nome, _dados in eventos] == ["processamento_lote"]


def test_falha_do_historico_nao_transforma_arquivo_salvo_em_erro(
    monkeypatch,
    tmp_path,
    controller_isolado,
):
    controller, view, _eventos, _historicos = controller_isolado
    entrada = tmp_path / "relatorio.csv"
    saida = tmp_path / "relatorio_tratado.xlsx"
    controller.arquivos_selecionados = [str(entrada)]

    monkeypatch.setattr(
        controller_module.filedialog,
        "asksaveasfilename",
        lambda **_kwargs: str(saida),
    )
    monkeypatch.setattr(
        controller_module,
        "registrar_historico",
        lambda _item: (_ for _ in ()).throw(OSError("histórico indisponível")),
    )
    monkeypatch.setattr(
        controller_module,
        "processar_arquivo",
        lambda *_args, **_kwargs: {
            "caminho_saida": str(saida),
            "tipo_entrada": "CSV",
            "quantidade_funcionarios": 1,
            "banco_total": "1:00",
            "banco_saldo": "0:30",
            "departamento": "Todos",
            "gerou_saldo": True,
            "gerou_resumo": True,
            "gerou_ranking": True,
        },
    )

    controller.processar("Todos")

    assert view.estados[-1] == (EstadoInterface.CONCLUIDO, 1)
    assert view.status[-1][1] == "success"
    assert controller.ultimo_resultado["caminho_saida"] == str(saida)


def test_falha_parcial_informa_e_expoe_arquivo_ja_salvo(
    monkeypatch,
    tmp_path,
    controller_isolado,
):
    controller, view, eventos, historicos = controller_isolado
    primeiro = tmp_path / "primeiro.csv"
    segundo = tmp_path / "segundo.csv"
    pasta_saida = tmp_path / "saida"
    controller.arquivos_selecionados = [str(primeiro), str(segundo)]

    monkeypatch.setattr(
        controller_module.filedialog,
        "askdirectory",
        lambda **_kwargs: str(pasta_saida),
    )

    def processar_em_teste(entrada, saida, departamento, **_opcoes):
        if entrada == str(segundo):
            raise AppError("Segundo arquivo inválido")
        return {
            "caminho_saida": saida,
            "tipo_entrada": "CSV",
            "quantidade_funcionarios": 2,
            "banco_total": "28:00",
            "banco_saldo": "-3:15",
            "departamento": departamento,
            "gerou_saldo": True,
            "gerou_resumo": True,
            "gerou_ranking": True,
        }

    monkeypatch.setattr(controller_module, "processar_arquivo", processar_em_teste)

    controller.processar("Todos")

    assert view.estados[-1] == (EstadoInterface.ERRO, 2)
    assert "1 arquivo foi salvo antes da falha" in view.status[-1][0]
    assert view.metricas[-1] == (2, "28:00", "-3:15")
    assert view.botao_abrir_habilitado[-1] is True
    assert view.botao_pasta_habilitado[-1] is True
    assert len(historicos) == 1
    assert Path(controller.ultimo_resultado["caminho_saida"]).name == (
        "primeiro_todos_tratado.xlsx"
    )
    assert [nome for nome, _dados in eventos] == ["arquivo_ignorado"]


def test_cancelamento_interrompe_lote_antes_do_proximo_arquivo(
    monkeypatch,
    tmp_path,
    controller_isolado,
):
    controller, view, _eventos, _historicos = controller_isolado
    runner = TaskRunnerPendente()
    controller._task_runner = runner
    primeiro = tmp_path / "primeiro.csv"
    segundo = tmp_path / "segundo.csv"
    controller.arquivos_selecionados = [str(primeiro), str(segundo)]
    chamadas = []

    monkeypatch.setattr(
        controller_module.filedialog,
        "askdirectory",
        lambda **_kwargs: str(tmp_path / "saida"),
    )
    monkeypatch.setattr(
        controller_module,
        "processar_arquivo",
        lambda entrada, saida, departamento, **_opcoes: chamadas.append(entrada)
        or {
            "caminho_saida": saida,
            "tipo_entrada": "CSV",
            "quantidade_funcionarios": 1,
            "banco_total": "1:00",
            "banco_saldo": "0:30",
            "departamento": departamento,
            "gerou_saldo": True,
            "gerou_resumo": True,
            "gerou_ranking": True,
        },
    )

    controller.processar("Todos")
    assert view.estados[-1] == (EstadoInterface.PROCESSANDO, 2)
    assert controller.cancelar_processamento() is None
    assert view.estados[-1] == (EstadoInterface.CANCELANDO, 2)

    runner.concluir()

    assert chamadas == []
    assert view.estados[-1] == (EstadoInterface.CANCELADO, 2)
    assert "cancelado" in view.status[-1][0].lower()


def test_relatorio_rhid_gera_excel_e_expoe_resultado(
    monkeypatch,
    tmp_path,
    controller_isolado,
):
    controller, view, eventos, historicos = controller_isolado
    saida = tmp_path / "relatorio_rhid.xlsx"
    cliente = type("Cliente", (), {"autenticado": True})()
    controller._rhid_client = cliente
    controller._rhid_empresas = (RhidCompany(7, "Projeto A", "Projeto A"),)
    controller._rhid_departamentos = (RhidDepartment(12, "Operações", 7),)

    monkeypatch.setattr(
        controller_module.filedialog,
        "asksaveasfilename",
        lambda **_kwargs: str(saida),
    )
    monkeypatch.setattr(
        controller_module,
        "BackgroundTaskRunner",
        lambda _agendar: TaskRunnerImediato(),
    )

    def processar(_cliente, plano, reportar):
        assert plano.company_id == 7
        assert plano.department_id == 12
        assert plano.gerar_saldo is True
        assert plano.gerar_resumo is True
        assert plano.gerar_ranking is True
        reportar(50)
        return {
            "caminho_saida": str(saida),
            "tipo_entrada": "RHID",
            "quantidade_funcionarios": 4,
            "banco_total": "24:15",
            "banco_saldo": "0:00",
            "departamento": "Operações",
            "gerou_saldo": True,
            "gerou_resumo": True,
            "gerou_ranking": True,
        }

    monkeypatch.setattr(controller_module, "processar_relatorio_rhid", processar)

    controller.gerar_relatorio_rhid(7, 12, "2026-08-01", "2026-08-11")

    assert view.rhid_ocupado == [True, False]
    assert view.rhid_progressos[0][0] == 0
    assert view.rhid_progressos[-1][0] == 100
    assert view.rhid_sucessos == [f"Salvo em: {saida}"]
    assert view.metricas[-1] == (4, "24:15", "0:00")
    assert view.estados[-1] == (EstadoInterface.CONCLUIDO, 1)
    assert controller.ultimo_resultado["tipo_entrada"] == "RHID"
    assert len(historicos) == 1
    assert [nome for nome, _dados in eventos] == ["processamento_rhid"]
    assert eventos[0][1]["departamento_ids"] == [12]
    assert "departamento_id" not in eventos[0][1]


@pytest.mark.parametrize("departamento_ids", [[12, 13], (12, 13)])
def test_relatorio_rhid_aceita_multiplos_setores_e_periodo_extenso(
    monkeypatch,
    tmp_path,
    controller_isolado,
    departamento_ids,
):
    controller, view, eventos, _historicos = controller_isolado
    saida = tmp_path / "relatorio_rhid_multissetor.xlsx"
    cliente = type("Cliente", (), {"autenticado": True})()
    controller._rhid_client = cliente
    controller._rhid_empresas = (RhidCompany(7, "Projeto A", "Projeto A"),)
    controller._rhid_departamentos = (
        RhidDepartment(12, "Operações", 7),
        RhidDepartment(13, "Financeiro", 7),
    )

    monkeypatch.setattr(
        controller_module.filedialog,
        "asksaveasfilename",
        lambda **_kwargs: str(saida),
    )
    monkeypatch.setattr(
        controller_module,
        "BackgroundTaskRunner",
        lambda _agendar: TaskRunnerImediato(),
    )

    def processar(_cliente, plano, _reportar):
        assert plano.company_id == 7
        assert plano.department_id == (12, 13)
        assert plano.department_label == "2 setores selecionados"
        assert plano.gerar_saldo is False
        assert plano.gerar_resumo is True
        assert plano.gerar_ranking is False
        return {
            "caminho_saida": str(saida),
            "tipo_entrada": "RHID",
            "quantidade_funcionarios": 8,
            "banco_total": "40:00",
            "banco_saldo": "2:00",
            "departamento": plano.department_label,
            "gerou_saldo": plano.gerar_saldo,
            "gerou_resumo": plano.gerar_resumo,
            "gerou_ranking": plano.gerar_ranking,
        }

    monkeypatch.setattr(controller_module, "processar_relatorio_rhid", processar)

    controller.gerar_relatorio_rhid(
        7,
        departamento_ids,
        "2026-01-01",
        "2026-08-11",
        gerar_saldo=False,
        gerar_resumo=True,
        gerar_ranking=False,
    )

    assert view.rhid_erros == []
    assert view.arquivos[-1].startswith(
        "RHiD: Projeto A / 2 setores selecionados /"
    )
    assert eventos[0][1]["departamento_ids"] == [12, 13]
    assert "departamento_id" not in eventos[0][1]


def test_relatorio_rhid_recusa_setor_de_outra_empresa(
    monkeypatch,
    controller_isolado,
):
    controller, view, _eventos, _historicos = controller_isolado
    cliente = type("Cliente", (), {"autenticado": True})()
    controller._rhid_client = cliente
    controller._rhid_empresas = (
        RhidCompany(7, "Projeto A", "Projeto A"),
        RhidCompany(8, "Projeto B", "Projeto B"),
    )
    controller._rhid_departamentos = (
        RhidDepartment(12, "Operações", 7),
        RhidDepartment(20, "Comercial", 8),
    )
    dialogos = []
    monkeypatch.setattr(
        controller_module.filedialog,
        "asksaveasfilename",
        lambda **_kwargs: dialogos.append(True),
    )

    controller.gerar_relatorio_rhid(
        7,
        [12, 20],
        "2026-08-01",
        "2026-08-11",
    )

    assert dialogos == []
    assert view.rhid_erros == ["Selecione um setor válido para essa empresa."]


def test_relatorio_rhid_aceita_setor_global_com_funcionarios(
    monkeypatch,
    controller_isolado,
):
    controller, view, _eventos, _historicos = controller_isolado
    controller._rhid_client = type("Cliente", (), {"autenticado": True})()
    controller._rhid_empresas = (RhidCompany(7, "Projeto A", "Projeto A"),)
    controller._rhid_departamentos = (
        RhidDepartment(12, "Setor compartilhado", 0),
    )
    dialogos = []
    monkeypatch.setattr(
        controller_module.filedialog,
        "asksaveasfilename",
        lambda **_kwargs: dialogos.append(True) or "",
    )

    controller.gerar_relatorio_rhid(
        7,
        12,
        "2026-08-01",
        "2026-08-11",
    )

    assert dialogos == [True]
    assert view.rhid_erros == []


def test_envia_ultimo_relatorio_ao_powerbi_em_background(
    monkeypatch,
    tmp_path,
    controller_isolado,
):
    controller, view, eventos, _historicos = controller_isolado
    caminho = tmp_path / "jornada.xlsx"
    caminho.write_bytes(b"xlsx")
    controller.ultimo_resultado = {
        "caminho_saida": str(caminho),
        "tipo_entrada": "CSV",
    }
    snapshot = PowerBiSnapshot(
        "relatorio-123",
        ({"IDRelatorio": "relatorio-123", "Funcionario": "Ana"},),
        str(caminho),
    )

    class ClienteFalso:
        autenticado = False

        def login_interativo(self):
            self.autenticado = True
            return None

        def obter_ou_criar_dataset(self):
            return "dataset-456"

        def enviar_linhas(self, _dataset_id, linhas, ao_progresso):
            ao_progresso(len(linhas), len(linhas))
            return len(linhas)

    monkeypatch.setattr(
        controller_module,
        "preparar_snapshot_powerbi",
        lambda *_args, **_kwargs: snapshot,
    )
    monkeypatch.setattr(controller_module, "PowerBiClient", ClienteFalso)
    abertos = []
    monkeypatch.setattr(
        controller_module,
        "abrir_relatorio_powerbi_desktop",
        lambda dataset_id: abertos.append(dataset_id),
    )
    monkeypatch.setattr(
        controller_module,
        "BackgroundTaskRunner",
        lambda _agendar: TaskRunnerImediato(),
    )

    controller.enviar_ultimo_resultado_powerbi()

    assert view.powerbi_ocupado == [True, False]
    assert view.powerbi_erros == []
    assert any("1 funcionário enviado" in item for item in view.powerbi_sucessos)
    assert "Power BI Desktop aberto" in view.powerbi_sucessos[-1]
    assert controller.ultimo_resultado["powerbi_report_id"] == "relatorio-123"
    assert controller.ultimo_resultado["powerbi_dataset_id"] == "dataset-456"
    assert view.powerbi_enviado is True
    assert abertos == ["dataset-456"]
    assert eventos[-1][0] == "envio_powerbi"
    assert eventos[-1][1]["quantidade_linhas"] == 1
