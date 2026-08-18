from app.controllers.connection_diagnostics_workflow import (
    ConnectionDiagnosticsWorkflow,
)
from app.integrations.powerbi_client import PowerBiApiError, PowerBiModelInfo, PowerBiWorkspaceInfo


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
        self.results = []
        self.busy = []
        self.summary = []
        self.reset = 0

    def agendar_na_interface(self, _delay, callback):
        callback()

    def reiniciar_diagnostico(self):
        self.reset += 1

    def definir_diagnostico_ocupado(self, value):
        self.busy.append(value)

    def atualizar_diagnostico(self, key, status, message):
        self.results.append((key, status, message))

    def finalizar_diagnostico(self, message, status):
        self.summary.append((message, status))


class FakeRhidClient:
    autenticado = True

    def listar_empresas(self):
        return (object(), object())


class FakePowerBiClient:
    autenticado = True

    def verificar_workspace(self):
        return PowerBiWorkspaceInfo("workspace-1", "FAS Jornada Analytics")

    def verificar_modelo(self):
        return PowerBiModelInfo("modelo-1", "FAS Jornada Analytics v2", True)


def _workflow(view, rhid=None, powerbi=None, cached=None):
    return ConnectionDiagnosticsWorkflow(
        view,
        operation_in_progress=lambda: False,
        rhid_client_provider=lambda: rhid,
        powerbi_client_provider=lambda: powerbi,
        cache_powerbi_client=lambda client: cached.append(client) if cached is not None else None,
        runner_factory=ImmediateRunner,
    )


def test_diagnostico_valida_quatro_conexoes_sem_gerar_relatorio():
    view = FakeView()
    cached = []
    powerbi = FakePowerBiClient()
    workflow = _workflow(view, FakeRhidClient(), powerbi, cached)

    workflow.run()

    terminal = {key: status for key, status, _message in view.results if status != "running"}
    assert terminal == {
        "rhid": "success",
        "microsoft": "success",
        "workspace": "success",
        "powerbi_model": "success",
    }
    assert view.busy == [True, False]
    assert view.summary[-1][1] == "success"
    assert cached == [powerbi]


def test_falha_rhid_nao_impede_diagnostico_powerbi():
    view = FakeView()
    workflow = _workflow(view, None, FakePowerBiClient())

    workflow.run()

    terminal = {key: status for key, status, _message in view.results if status != "running"}
    assert terminal["rhid"] == "error"
    assert terminal["workspace"] == "success"
    assert terminal["powerbi_model"] == "success"
    assert view.summary[-1][1] == "error"


def test_falha_microsoft_pula_workspace_e_modelo():
    class FailingPowerBiClient:
        autenticado = False

        def login_interativo(self):
            raise PowerBiApiError("Login recusado.")

    view = FakeView()
    workflow = _workflow(view, FakeRhidClient(), FailingPowerBiClient())

    workflow.run()

    terminal = {key: status for key, status, _message in view.results if status != "running"}
    assert terminal["microsoft"] == "error"
    assert terminal["workspace"] == "skipped"
    assert terminal["powerbi_model"] == "skipped"
