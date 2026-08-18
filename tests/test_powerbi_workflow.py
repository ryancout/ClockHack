from app.controllers.powerbi_workflow import PowerBiWorkflow
from app.controllers.workflow_state import WorkflowState
from app.services.analytics_service import PowerBiSnapshot
from app.services.powerbi_send_registry import (
    PowerBiSendRegistry,
    calcular_fingerprint_snapshot,
)


class ImmediateRunner:
    ativo = False

    def __init__(self, _schedule):
        pass

    def executar(self, work, on_progress, on_complete, on_error):
        self.ativo = True
        try:
            result = work(on_progress)
        except Exception as error:
            self.ativo = False
            on_error(error)
        else:
            self.ativo = False
            on_complete(result)
        return True


class FakeView:
    def __init__(self):
        self.busy = []
        self.progress = []
        self.status = []
        self.errors = []
        self.success = []

    def agendar_na_interface(self, _delay, callback):
        callback()

    def definir_powerbi_ocupado(self, value):
        self.busy.append(value)

    def atualizar_progresso_powerbi(self, value, message=""):
        self.progress.append((value, message))

    def atualizar_status(self, message, kind):
        self.status.append((message, kind))

    def exibir_erro_powerbi(self, message):
        self.errors.append(message)

    def exibir_sucesso_powerbi(self, message):
        self.success.append(message)

    def definir_powerbi_enviado(self, _sent):
        pass


class FakeClient:
    autenticado = True

    def __init__(self):
        self.sent = 0

    def obter_ou_criar_dataset(self):
        return "dataset-1"

    def enviar_linhas(self, _dataset_id, rows, ao_progresso):
        self.sent += len(rows)
        ao_progresso(len(rows), len(rows))
        return len(rows)


def _snapshot(report_id):
    return PowerBiSnapshot(
        report_id,
        (
            {
                "IDRelatorio": report_id,
                "GeradoEm": "2026-08-18T12:00:00",
                "Arquivo": "jornada.xlsx",
                "Matricula": "001",
                "BancoSaldoMinutos": -60,
            },
        ),
        "jornada.xlsx",
    )


def _workflow(tmp_path, confirm):
    report_path = tmp_path / "jornada.xlsx"
    report_path.write_bytes(b"xlsx")
    state = WorkflowState(
        {},
        last_result={"caminho_saida": str(report_path), "tipo_entrada": "CSV"},
    )
    view = FakeView()
    registry = PowerBiSendRegistry(tmp_path / "powerbi_sends.json")
    original = _snapshot("original")
    registry.register(
        fingerprint=calcular_fingerprint_snapshot(original),
        report_id=original.report_id,
        dataset_id="dataset-1",
        source_file=str(report_path),
        row_count=1,
    )
    client = FakeClient()
    workflow = PowerBiWorkflow(
        view,
        state,
        operation_in_progress=lambda: False,
        runner_factory=ImmediateRunner,
        client_factory=lambda: client,
        snapshot_factory=lambda *_args, **_kwargs: _snapshot("novo"),
        desktop_opener=lambda _dataset_id: None,
        registry=registry,
        confirm_duplicate=lambda _title, _text: confirm,
    )
    return workflow, state, view, client


def test_reenvio_identico_e_cancelado_antes_de_chamar_powerbi(tmp_path):
    workflow, state, view, client = _workflow(tmp_path, confirm=False)

    workflow.send_last()

    assert client.sent == 0
    assert "powerbi_dataset_id" not in state.last_result
    assert view.status[-1][1] == "warning"
    assert "duplicados" in view.progress[-1][1]


def test_usuario_pode_confirmar_reenvio_intencional(tmp_path):
    workflow, state, view, client = _workflow(tmp_path, confirm=True)

    workflow.send_last()

    assert client.sent == 1
    assert state.last_result["powerbi_dataset_id"] == "dataset-1"
    assert view.errors == []
