import pytest

from app.services import rhid_credentials_service
from app.services.rhid_credentials_service import (
    CredentialStorageError,
    CredentialStorageUnavailable,
    RhidCredentialService,
    RhidCredentials,
    _FallbackCredentialBackend,
    _UnavailableCredentialBackend,
    _WindowsProtectedFileBackend,
)


class MemoryCredentialBackend:
    available = True

    def __init__(self):
        self.values = {}

    def write(self, target, username, secret):
        self.values[target] = (username, secret)

    def read(self, target):
        return self.values.get(target)

    def delete(self, target):
        return self.values.pop(target, None) is not None


class ReversingProtector:
    def protect(self, value):
        return b"encrypted:" + value[::-1]

    def unprotect(self, value):
        assert value.startswith(b"encrypted:")
        return value.removeprefix(b"encrypted:")[::-1]


class FailingCredentialBackend(MemoryCredentialBackend):
    def write(self, target, username, secret):
        raise CredentialStorageError("cofre indisponivel")


def test_salva_e_carrega_todos_os_dados_sem_json():
    backend = MemoryCredentialBackend()
    service = RhidCredentialService(backend=backend)

    service.save(" usuario@empresa.com ", "s3nh@ com acento: ç", " cliente-a ")

    assert service.load() == RhidCredentials(
        email="usuario@empresa.com",
        domain="cliente-a",
        password="s3nh@ com acento: ç",
    )
    _, blob = next(iter(backend.values.values()))
    assert not blob.startswith(b"{")
    assert b'"password"' not in blob


def test_senha_nao_aparece_na_representacao_da_credencial():
    credentials = RhidCredentials(
        email="usuario@empresa.com",
        domain="cliente-a",
        password="segredo",
    )

    assert "segredo" not in repr(credentials)


def test_carregar_sem_credencial_retorna_none():
    service = RhidCredentialService(backend=MemoryCredentialBackend())

    assert service.load() is None


def test_remover_credencial_e_idempotente():
    service = RhidCredentialService(backend=MemoryCredentialBackend())
    service.save("usuario@empresa.com", "segredo")

    assert service.delete() is True
    assert service.delete() is False
    assert service.load() is None


@pytest.mark.parametrize(
    ("email", "password", "message"),
    [
        ("", "segredo", "e-mail"),
        ("usuario@empresa.com", "", "senha"),
    ],
)
def test_rejeita_campos_obrigatorios_vazios(email, password, message):
    service = RhidCredentialService(backend=MemoryCredentialBackend())

    with pytest.raises(ValueError, match=message):
        service.save(email, password)


def test_credencial_corrompida_falha_sem_expor_conteudo():
    backend = MemoryCredentialBackend()
    backend.values["FASJornada:RHiD"] = (
        "usuario@empresa.com",
        b"senha-super-secreta",
    )
    service = RhidCredentialService(backend=backend)

    with pytest.raises(CredentialStorageError) as error:
        service.load()

    assert "senha-super-secreta" not in str(error.value)


def test_backend_indisponivel_falha_de_forma_explicita_e_nao_persiste():
    service = RhidCredentialService(backend=_UnavailableCredentialBackend())

    assert service.available is False
    with pytest.raises(CredentialStorageUnavailable, match="somente no Windows"):
        service.save("usuario@empresa.com", "segredo", "cliente-a")
    with pytest.raises(CredentialStorageUnavailable):
        service.load()
    with pytest.raises(CredentialStorageUnavailable):
        service.delete()


def test_credencial_acima_do_limite_do_windows_e_rejeitada():
    service = RhidCredentialService(backend=MemoryCredentialBackend())

    with pytest.raises(ValueError, match="limite seguro"):
        service.save("usuario@empresa.com", "x" * 3000)


def test_fallback_dpapi_grava_arquivo_binario_sem_texto_sensivel(tmp_path):
    path = tmp_path / "rhid_credentials.dat"
    backend = _WindowsProtectedFileBackend(path, ReversingProtector())
    service = RhidCredentialService(backend=backend)

    service.save("usuario@empresa.com", "senha-super-secreta", "cliente-a")

    content = path.read_bytes()
    assert not content.startswith(b"{")
    assert b"usuario@empresa.com" not in content
    assert b"senha-super-secreta" not in content
    assert service.load() == RhidCredentials(
        email="usuario@empresa.com",
        domain="cliente-a",
        password="senha-super-secreta",
    )


def test_usa_fallback_quando_credential_manager_rejeita_escrita():
    fallback = MemoryCredentialBackend()
    backend = _FallbackCredentialBackend(FailingCredentialBackend(), fallback)
    service = RhidCredentialService(backend=backend)

    service.save("usuario@empresa.com", "segredo", "cliente-a")

    assert service.load() == RhidCredentials(
        email="usuario@empresa.com",
        domain="cliente-a",
        password="segredo",
    )


def test_copia_do_fallback_tem_precedencia_sobre_credencial_antiga():
    primary = MemoryCredentialBackend()
    fallback = MemoryCredentialBackend()
    target = "FASJornada:RHiD"
    primary.values[target] = ("antigo@empresa.com", b"nao-deve-ser-lido")
    current = RhidCredentialService(backend=fallback)
    current.save("atual@empresa.com", "senha-atual", "cliente-atual")
    backend = _FallbackCredentialBackend(primary, fallback)

    credentials = RhidCredentialService(backend=backend).load()

    assert credentials.email == "atual@empresa.com"
    assert credentials.password == "senha-atual"


def test_plataforma_sem_cofre_nativo_nao_cria_fallback_inseguro(monkeypatch):
    monkeypatch.setattr(rhid_credentials_service.sys, "platform", "linux")

    service = RhidCredentialService()

    assert service.available is False
    with pytest.raises(CredentialStorageUnavailable):
        service.save("usuario@empresa.com", "segredo")
