"""Execucao de tarefas em segundo plano com callbacks na thread da interface."""

from __future__ import annotations

from collections.abc import Callable
from queue import Empty, Queue
from threading import Event, Lock, Thread
from typing import Any, Generic, Protocol, TypeVar


Resultado = TypeVar("Resultado")
Progresso = TypeVar("Progresso")


class _ThreadIniciavel(Protocol):
    def start(self) -> None: ...


Agendador = Callable[[int, Callable[[], None]], Any]
FabricaThread = Callable[..., _ThreadIniciavel]


class BackgroundTaskRunner(Generic[Resultado, Progresso]):
    """Executa um trabalho fora da interface e entrega seus eventos pelo agendador.

    ``agendar`` deve ter a mesma semantica de ``widget.after``: receber um atraso
    em milissegundos e uma funcao sem argumentos. Assim, o polling e todos os
    callbacks do cliente permanecem na thread responsavel pelo agendador.
    """

    _PROGRESSO = "progresso"
    _CONCLUIDO = "concluido"
    _ERRO = "erro"

    def __init__(
        self,
        agendar: Agendador,
        intervalo_ms: int = 50,
        thread_factory: FabricaThread = Thread,
    ) -> None:
        if intervalo_ms < 0:
            raise ValueError("intervalo_ms nao pode ser negativo")

        self._agendar = agendar
        self._intervalo_ms = intervalo_ms
        self._thread_factory = thread_factory
        self._lock = Lock()
        self._ativo = False
        self._cancelamento = None

    @property
    def ativo(self) -> bool:
        """Indica se ha uma execucao aguardando seu evento terminal."""

        with self._lock:
            return self._ativo

    @property
    def cancelamento_solicitado(self) -> bool:
        with self._lock:
            return bool(self._cancelamento and self._cancelamento.is_set())

    def cancelar(self) -> bool:
        """Solicita parada cooperativa; retorna se havia trabalho ativo."""
        with self._lock:
            if not self._ativo or self._cancelamento is None:
                return False
            self._cancelamento.set()
            return True

    def executar(
        self,
        trabalho: Callable[[Callable[[Progresso], None]], Resultado],
        ao_progresso: Callable[[Progresso], None],
        ao_concluir: Callable[[Resultado], None],
        ao_erro: Callable[[Exception], None],
    ) -> bool:
        """Inicia ``trabalho`` em uma thread dedicada.

        Retorna ``False`` sem iniciar outra thread quando ja existe uma tarefa
        ativa. O trabalho recebe uma funcao para publicar atualizacoes de
        progresso. Resultado, progresso e erro sao consumidos exclusivamente
        pelo polling disparado por ``agendar``.
        """

        with self._lock:
            if self._ativo:
                return False
            self._ativo = True
            self._cancelamento = Event()

        eventos: Queue[tuple[str, object]] = Queue()
        agendamento_falhou = Event()
        worker_finalizado = Event()

        def reportar_progresso(valor: Progresso) -> None:
            eventos.put((self._PROGRESSO, valor))

        def executar_trabalho() -> None:
            try:
                resultado = trabalho(reportar_progresso)
            except Exception as erro:
                eventos.put((self._ERRO, erro))
            else:
                eventos.put((self._CONCLUIDO, resultado))
            finally:
                worker_finalizado.set()
                if agendamento_falhou.is_set():
                    self._marcar_inativo()

        erro_callback_progresso = None

        def polling() -> None:
            nonlocal erro_callback_progresso
            terminou = False

            while not terminou:
                try:
                    tipo, valor = eventos.get_nowait()
                except Empty:
                    break

                if tipo == self._PROGRESSO:
                    try:
                        ao_progresso(valor)  # type: ignore[arg-type]
                    except Exception as erro_callback:
                        if erro_callback_progresso is None:
                            erro_callback_progresso = erro_callback
                    continue

                terminou = True
                self._marcar_inativo()
                if tipo == self._ERRO:
                    ao_erro(valor)  # type: ignore[arg-type]
                elif erro_callback_progresso is not None:
                    ao_erro(erro_callback_progresso)
                else:
                    try:
                        ao_concluir(valor)  # type: ignore[arg-type]
                    except Exception as erro_callback_terminal:
                        ao_erro(erro_callback_terminal)

            if not terminou:
                try:
                    self._agendar(self._intervalo_ms, polling)
                except Exception as erro_agendamento:
                    agendamento_falhou.set()
                    try:
                        ao_erro(erro_agendamento)
                    finally:
                        if worker_finalizado.is_set():
                            self._marcar_inativo()

        thread_iniciada = False
        try:
            thread = self._thread_factory(
                target=executar_trabalho,
                daemon=False,
                name="FASJornadaWorker",
            )
            thread.start()
            thread_iniciada = True
            self._agendar(0, polling)
        except Exception:
            if thread_iniciada:
                agendamento_falhou.set()
                if worker_finalizado.is_set():
                    self._marcar_inativo()
            else:
                self._marcar_inativo()
            raise

        return True

    def _marcar_inativo(self) -> None:
        with self._lock:
            self._ativo = False
            self._cancelamento = None
