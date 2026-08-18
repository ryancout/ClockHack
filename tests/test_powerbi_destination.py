from app.integrations.powerbi_destination import (
    PowerBiDestination,
    PowerBiDestinationKind,
    PushSemanticModelDestination,
)


class FakeClient:
    autenticado = True

    def login_interativo(self):
        pass

    def obter_ou_criar_dataset(self):
        return "dataset-1"

    def enviar_linhas(self, dataset_id, linhas, *, ao_progresso=None):
        rows = list(linhas)
        if ao_progresso:
            ao_progresso(len(rows), len(rows))
        return len(rows)


def test_adaptador_push_cumpre_fronteira_de_destino():
    destination = PushSemanticModelDestination(FakeClient())
    progress = []

    result = destination.publish(
        [{"Funcionario": "Ana"}],
        on_progress=lambda sent, total: progress.append((sent, total)),
    )

    assert isinstance(destination, PowerBiDestination)
    assert result.destination is PowerBiDestinationKind.PUSH_SEMANTIC_MODEL
    assert result.resource_id == "dataset-1"
    assert result.row_count == 1
    assert progress == [(1, 1)]
