from __future__ import annotations

from collections import deque
from threading import Event, Thread, get_ident
from time import sleep

import pytest

from app.services.background_task_runner import BackgroundTaskRunner


class AgendadorManual:
    def __init__(self) -> None:
        self.callbacks = deque()
        self.atrasos = []

    def agendar(self, atraso_ms, callback) -> None:
        self.atrasos.append(atraso_ms)
        self.callbacks.append(callback)

    def executar_ate(self, condicao, limite=1000) -> None:
        for _ in range(limite):
            if condicao():
                return
            if self.callbacks:
                self.callbacks.popleft()()
            sleep(0.001)
        raise AssertionError("O agendador nao atingiu a condicao esperada")


def test_entrega_progresso_e_resultado_pelo_agendador() -> None:
    agendador = AgendadorManual()
    terminou = Event()
    progressos = []
    resultados = []
    erros = []
    thread_trabalho = []

    def trabalho(reportar):
        thread_trabalho.append(get_ident())
        reportar(25)
        reportar(100)
        terminou.set()
        return "arquivo.xlsx"

    runner = BackgroundTaskRunner(agendador.agendar, intervalo_ms=1)

    assert runner.executar(trabalho, progressos.append, resultados.append, erros.append)
    assert runner.ativo
    assert terminou.wait(timeout=2)
    assert progressos == []
    assert resultados == []

    agendador.executar_ate(lambda: not runner.ativo)

    assert progressos == [25, 100]
    assert resultados == ["arquivo.xlsx"]
    assert erros == []
    assert thread_trabalho[0] != get_ident()
    assert agendador.atrasos[0] == 0


def test_entrega_erro_sem_chamar_conclusao() -> None:
    agendador = AgendadorManual()
    terminou = Event()
    erros = []
    resultados = []

    def trabalho(_reportar):
        terminou.set()
        raise ValueError("falha de teste")

    runner = BackgroundTaskRunner(agendador.agendar, intervalo_ms=1)

    assert runner.executar(trabalho, lambda _valor: None, resultados.append, erros.append)
    assert terminou.wait(timeout=2)
    agendador.executar_ate(lambda: not runner.ativo)

    assert resultados == []
    assert len(erros) == 1
    assert isinstance(erros[0], ValueError)
    assert str(erros[0]) == "falha de teste"


def test_recusa_reentrada_enquanto_execucao_esta_ativa() -> None:
    agendador = AgendadorManual()
    liberar = Event()
    iniciou = Event()
    terminou = Event()
    execucoes = []

    def trabalho(_reportar):
        execucoes.append("primeira")
        iniciou.set()
        assert liberar.wait(timeout=2)
        terminou.set()
        return 1

    runner = BackgroundTaskRunner(agendador.agendar, intervalo_ms=1)

    assert runner.executar(trabalho, lambda _valor: None, lambda _valor: None, pytest.fail)
    assert iniciou.wait(timeout=2)
    assert not runner.executar(
        lambda _reportar: execucoes.append("segunda"),
        lambda _valor: None,
        lambda _valor: None,
        pytest.fail,
    )

    liberar.set()
    assert terminou.wait(timeout=2)
    agendador.executar_ate(lambda: not runner.ativo)
    assert execucoes == ["primeira"]


def test_callbacks_rodam_na_thread_que_executa_o_agendador() -> None:
    agendador = AgendadorManual()
    trabalho_terminou = Event()
    threads_callbacks = []
    thread_trabalho = []

    def trabalho(reportar):
        thread_trabalho.append(get_ident())
        reportar("metade")
        trabalho_terminou.set()
        return "fim"

    runner = BackgroundTaskRunner(agendador.agendar, intervalo_ms=1)
    thread_agendador = get_ident()

    assert runner.executar(
        trabalho,
        lambda _valor: threads_callbacks.append(get_ident()),
        lambda _valor: threads_callbacks.append(get_ident()),
        lambda _erro: threads_callbacks.append(get_ident()),
    )
    assert trabalho_terminou.wait(timeout=2)
    agendador.executar_ate(lambda: not runner.ativo)

    assert threads_callbacks == [thread_agendador, thread_agendador]
    assert len(thread_trabalho) == 1
    assert thread_trabalho[0] != thread_agendador


def test_thread_nao_daemon_protege_gravacao_em_andamento() -> None:
    agendador = AgendadorManual()
    argumentos = {}

    def fabrica_thread(**kwargs):
        argumentos.update(kwargs)
        return Thread(**kwargs)

    runner = BackgroundTaskRunner(
        agendador.agendar,
        intervalo_ms=1,
        thread_factory=fabrica_thread,
    )

    assert runner.executar(lambda _reportar: None, lambda _valor: None, lambda _valor: None, pytest.fail)
    agendador.executar_ate(lambda: not runner.ativo)

    assert argumentos["daemon"] is False
    assert argumentos["name"] == "FASJornadaWorker"


def test_falha_no_callback_de_progresso_nao_deixa_runner_travado() -> None:
    agendador = AgendadorManual()
    progresso_publicado = Event()
    liberar_trabalho = Event()
    erros = []
    segunda_execucao = []

    def trabalho(reportar):
        reportar("evento")
        progresso_publicado.set()
        assert liberar_trabalho.wait(timeout=2)
        return "resultado que nao deve ser entregue"

    def falhar_callback(_valor):
        raise RuntimeError("falha na interface")

    runner = BackgroundTaskRunner(agendador.agendar, intervalo_ms=1)

    assert runner.executar(trabalho, falhar_callback, pytest.fail, erros.append)
    assert progresso_publicado.wait(timeout=2)
    agendador.callbacks.popleft()()

    assert runner.ativo is True
    assert not runner.executar(
        lambda _reportar: segunda_execucao.append("iniciada"),
        lambda _valor: None,
        lambda _valor: None,
        erros.append,
    )
    assert erros == []

    liberar_trabalho.set()
    agendador.executar_ate(lambda: not runner.ativo)

    assert len(erros) == 1
    assert isinstance(erros[0], RuntimeError)
    assert str(erros[0]) == "falha na interface"
    assert segunda_execucao == []


def test_falha_ao_agendar_nao_libera_reentrada_com_worker_ativo() -> None:
    iniciou = Event()
    liberar = Event()
    terminou = Event()

    def agendar_com_falha(_atraso, _callback):
        raise RuntimeError("loop da interface indisponível")

    def trabalho(_reportar):
        iniciou.set()
        assert liberar.wait(timeout=2)
        terminou.set()
        return "fim"

    runner = BackgroundTaskRunner(agendar_com_falha)

    with pytest.raises(RuntimeError, match="loop da interface indisponível"):
        runner.executar(
            trabalho,
            lambda _valor: None,
            lambda _valor: None,
            lambda _erro: None,
        )

    assert iniciou.wait(timeout=2)
    assert runner.ativo is True
    assert not runner.executar(
        lambda _reportar: None,
        lambda _valor: None,
        lambda _valor: None,
        lambda _erro: None,
    )

    liberar.set()
    assert terminou.wait(timeout=2)
    for _ in range(100):
        if not runner.ativo:
            break
        sleep(0.001)
    assert runner.ativo is False


def test_falha_no_callback_de_conclusao_e_encaminhada_com_runner_inativo() -> None:
    agendador = AgendadorManual()
    trabalho_terminou = Event()
    erros = []
    estados_ao_erro = []

    def trabalho(_reportar):
        trabalho_terminou.set()
        return "resultado"

    def falhar_conclusao(_resultado):
        raise RuntimeError("falha ao finalizar a interface")

    runner = BackgroundTaskRunner(agendador.agendar, intervalo_ms=1)

    assert runner.executar(
        trabalho,
        lambda _valor: None,
        falhar_conclusao,
        lambda erro: (erros.append(erro), estados_ao_erro.append(runner.ativo)),
    )
    assert trabalho_terminou.wait(timeout=2)
    agendador.executar_ate(lambda: not runner.ativo)

    assert len(erros) == 1
    assert isinstance(erros[0], RuntimeError)
    assert estados_ao_erro == [False]


def test_falha_ao_reagendar_mantem_worker_protegido_ate_terminar() -> None:
    callbacks = deque()
    chamadas_agendador = 0
    progresso_publicado = Event()
    liberar = Event()
    terminou = Event()
    erros = []

    def agendar(_atraso, callback):
        nonlocal chamadas_agendador
        chamadas_agendador += 1
        if chamadas_agendador == 1:
            callbacks.append(callback)
            return
        raise RuntimeError("falha ao reagendar polling")

    def trabalho(reportar):
        reportar("iniciado")
        progresso_publicado.set()
        assert liberar.wait(timeout=2)
        terminou.set()
        return "fim"

    runner = BackgroundTaskRunner(agendar, intervalo_ms=1)

    assert runner.executar(
        trabalho,
        lambda _valor: None,
        lambda _valor: None,
        erros.append,
    )
    assert progresso_publicado.wait(timeout=2)
    callbacks.popleft()()

    assert len(erros) == 1
    assert isinstance(erros[0], RuntimeError)
    assert runner.ativo is True
    assert not runner.executar(
        lambda _reportar: None,
        lambda _valor: None,
        lambda _valor: None,
        erros.append,
    )

    liberar.set()
    assert terminou.wait(timeout=2)
    for _ in range(100):
        if not runner.ativo:
            break
        sleep(0.001)
    assert runner.ativo is False


def test_cancelamento_eh_solicitado_sem_interromper_a_thread() -> None:
    agendador = AgendadorManual()
    cancelamento_observado = Event()
    terminou = Event()
    resultados = []

    runner = None

    def trabalho(_reportar):
        while not runner.cancelamento_solicitado:
            sleep(0.001)
        cancelamento_observado.set()
        terminou.set()
        return "finalizado com seguranca"

    runner = BackgroundTaskRunner(agendador.agendar, intervalo_ms=1)
    assert runner.executar(
        trabalho,
        lambda _valor: None,
        resultados.append,
        pytest.fail,
    )
    assert runner.cancelar() is True
    assert runner.cancelar() is True
    assert cancelamento_observado.wait(timeout=2)
    assert terminou.wait(timeout=2)
    agendador.executar_ate(lambda: not runner.ativo)

    assert resultados == ["finalizado com seguranca"]
    assert runner.cancelar() is False
