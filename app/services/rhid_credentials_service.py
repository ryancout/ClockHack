"""Armazenamento seguro das credenciais usadas no login do RHiD.

A senha nunca e gravada nas preferencias da aplicacao. No Windows, os dados
sao mantidos pelo Credential Manager. Quando o cofre nao esta acessivel, a
camada usa um arquivo binario protegido pela DPAPI para o usuario conectado.
"""

from __future__ import annotations

import ctypes
import os
import struct
import sys
import tempfile
from ctypes import wintypes
from dataclasses import dataclass, field
from pathlib import Path
from typing import NoReturn, Protocol

from app.core.config import DATA_DIR


DEFAULT_CREDENTIAL_TARGET = "FASJornada:RHiD"

_CREDENTIAL_TYPE_GENERIC = 1
_CREDENTIAL_PERSIST_LOCAL_MACHINE = 2
_ERROR_NOT_FOUND = 1168
_MAX_CREDENTIAL_BLOB_SIZE = 2560
_SECRET_MAGIC = b"FJR1"
_SECRET_HEADER = struct.Struct(">4sII")
_PROTECTED_RECORD_MAGIC = b"FJC1"
_PROTECTED_RECORD_HEADER = struct.Struct(">4sII")
_DPAPI_ENTROPY = b"FAS Jornada:RHiD:v1"
_CRYPTPROTECT_UI_FORBIDDEN = 0x1


class CredentialStorageError(RuntimeError):
    """Falha segura ao acessar o armazenamento de credenciais."""


class CredentialStorageUnavailable(CredentialStorageError):
    """O sistema atual nao possui o armazenamento seguro suportado."""


@dataclass(frozen=True, slots=True)
class RhidCredentials:
    """Credenciais lembradas pelo usuario para o login no RHiD."""

    email: str
    domain: str
    password: str = field(repr=False)


class _CredentialBackend(Protocol):
    @property
    def available(self) -> bool: ...

    def write(self, target: str, username: str, secret: bytes) -> None: ...

    def read(self, target: str) -> tuple[str, bytes] | None: ...

    def delete(self, target: str) -> bool: ...


class _UnavailableCredentialBackend:
    @property
    def available(self) -> bool:
        return False

    @staticmethod
    def _raise() -> NoReturn:
        raise CredentialStorageUnavailable(
            "O armazenamento seguro de credenciais esta disponivel somente no Windows."
        )

    def write(self, target: str, username: str, secret: bytes) -> None:
        self._raise()

    def read(self, target: str) -> tuple[str, bytes] | None:
        self._raise()

    def delete(self, target: str) -> bool:
        self._raise()


class _CREDENTIAL_ATTRIBUTEW(ctypes.Structure):
    _fields_ = [
        ("Keyword", wintypes.LPWSTR),
        ("Flags", wintypes.DWORD),
        ("ValueSize", wintypes.DWORD),
        ("Value", ctypes.POINTER(ctypes.c_ubyte)),
    ]


class _CREDENTIALW(ctypes.Structure):
    _fields_ = [
        ("Flags", wintypes.DWORD),
        ("Type", wintypes.DWORD),
        ("TargetName", wintypes.LPWSTR),
        ("Comment", wintypes.LPWSTR),
        ("LastWritten", wintypes.FILETIME),
        ("CredentialBlobSize", wintypes.DWORD),
        ("CredentialBlob", ctypes.POINTER(ctypes.c_ubyte)),
        ("Persist", wintypes.DWORD),
        ("AttributeCount", wintypes.DWORD),
        ("Attributes", ctypes.POINTER(_CREDENTIAL_ATTRIBUTEW)),
        ("TargetAlias", wintypes.LPWSTR),
        ("UserName", wintypes.LPWSTR),
    ]


class _DATA_BLOB(ctypes.Structure):
    _fields_ = [
        ("cbData", wintypes.DWORD),
        ("pbData", ctypes.POINTER(ctypes.c_ubyte)),
    ]


class _WindowsCredentialBackend:
    """Adaptador minimo para a API WinCred, sem dependencia externa."""

    def __init__(self) -> None:
        try:
            library = ctypes.WinDLL("Advapi32.dll", use_last_error=True)
        except (AttributeError, OSError) as exc:
            raise CredentialStorageUnavailable(
                "O Gerenciador de Credenciais do Windows nao esta disponivel."
            ) from exc

        pointer = ctypes.POINTER(_CREDENTIALW)
        library.CredWriteW.argtypes = [pointer, wintypes.DWORD]
        library.CredWriteW.restype = wintypes.BOOL
        library.CredReadW.argtypes = [
            wintypes.LPCWSTR,
            wintypes.DWORD,
            wintypes.DWORD,
            ctypes.POINTER(pointer),
        ]
        library.CredReadW.restype = wintypes.BOOL
        library.CredDeleteW.argtypes = [
            wintypes.LPCWSTR,
            wintypes.DWORD,
            wintypes.DWORD,
        ]
        library.CredDeleteW.restype = wintypes.BOOL
        library.CredFree.argtypes = [wintypes.LPVOID]
        library.CredFree.restype = None
        self._library = library

    @property
    def available(self) -> bool:
        return True

    def write(self, target: str, username: str, secret: bytes) -> None:
        blob = (ctypes.c_ubyte * len(secret)).from_buffer_copy(secret)
        credential = _CREDENTIALW()
        credential.Type = _CREDENTIAL_TYPE_GENERIC
        credential.TargetName = target
        credential.CredentialBlobSize = len(secret)
        credential.CredentialBlob = ctypes.cast(
            blob, ctypes.POINTER(ctypes.c_ubyte)
        )
        credential.Persist = _CREDENTIAL_PERSIST_LOCAL_MACHINE
        credential.UserName = username

        try:
            if not self._library.CredWriteW(ctypes.byref(credential), 0):
                self._raise_windows_error("salvar")
        finally:
            ctypes.memset(blob, 0, len(secret))

    def read(self, target: str) -> tuple[str, bytes] | None:
        credential_pointer = ctypes.POINTER(_CREDENTIALW)()
        if not self._library.CredReadW(
            target,
            _CREDENTIAL_TYPE_GENERIC,
            0,
            ctypes.byref(credential_pointer),
        ):
            error_code = ctypes.get_last_error()
            if error_code == _ERROR_NOT_FOUND:
                return None
            self._raise_windows_error("carregar", error_code)

        try:
            credential = credential_pointer.contents
            username = credential.UserName or ""
            secret = ctypes.string_at(
                credential.CredentialBlob,
                credential.CredentialBlobSize,
            )
            return username, secret
        finally:
            self._library.CredFree(credential_pointer)

    def delete(self, target: str) -> bool:
        if self._library.CredDeleteW(target, _CREDENTIAL_TYPE_GENERIC, 0):
            return True
        error_code = ctypes.get_last_error()
        if error_code == _ERROR_NOT_FOUND:
            return False
        self._raise_windows_error("remover", error_code)
        return False

    @staticmethod
    def _raise_windows_error(action: str, error_code: int | None = None) -> None:
        code = ctypes.get_last_error() if error_code is None else error_code
        raise CredentialStorageError(
            f"Nao foi possivel {action} as credenciais no Windows (erro {code})."
        )


class _DpapiProtector:
    """Protege bytes com a conta Windows atual, sem exibir prompts."""

    def __init__(self) -> None:
        try:
            crypt32 = ctypes.WinDLL("Crypt32.dll", use_last_error=True)
            kernel32 = ctypes.WinDLL("Kernel32.dll", use_last_error=True)
        except (AttributeError, OSError) as exc:
            raise CredentialStorageUnavailable(
                "A protecao de dados do Windows nao esta disponivel."
            ) from exc

        blob_pointer = ctypes.POINTER(_DATA_BLOB)
        crypt32.CryptProtectData.argtypes = [
            blob_pointer,
            wintypes.LPCWSTR,
            blob_pointer,
            wintypes.LPVOID,
            wintypes.LPVOID,
            wintypes.DWORD,
            blob_pointer,
        ]
        crypt32.CryptProtectData.restype = wintypes.BOOL
        crypt32.CryptUnprotectData.argtypes = [
            blob_pointer,
            ctypes.POINTER(wintypes.LPWSTR),
            blob_pointer,
            wintypes.LPVOID,
            wintypes.LPVOID,
            wintypes.DWORD,
            blob_pointer,
        ]
        crypt32.CryptUnprotectData.restype = wintypes.BOOL
        kernel32.LocalFree.argtypes = [wintypes.HLOCAL]
        kernel32.LocalFree.restype = wintypes.HLOCAL
        self._crypt32 = crypt32
        self._kernel32 = kernel32

    def protect(self, value: bytes) -> bytes:
        return self._transform("protect", value)

    def unprotect(self, value: bytes) -> bytes:
        return self._transform("unprotect", value)

    def _transform(self, operation: str, value: bytes) -> bytes:
        value_blob, value_buffer = self._input_blob(value)
        entropy_blob, entropy_buffer = self._input_blob(_DPAPI_ENTROPY)
        output = _DATA_BLOB()
        description = wintypes.LPWSTR()
        try:
            if operation == "protect":
                succeeded = self._crypt32.CryptProtectData(
                    ctypes.byref(value_blob),
                    "FAS Jornada - RHiD",
                    ctypes.byref(entropy_blob),
                    None,
                    None,
                    _CRYPTPROTECT_UI_FORBIDDEN,
                    ctypes.byref(output),
                )
            else:
                succeeded = self._crypt32.CryptUnprotectData(
                    ctypes.byref(value_blob),
                    ctypes.byref(description),
                    ctypes.byref(entropy_blob),
                    None,
                    None,
                    _CRYPTPROTECT_UI_FORBIDDEN,
                    ctypes.byref(output),
                )
            if not succeeded:
                code = ctypes.get_last_error()
                raise CredentialStorageError(
                    f"A protecao de credenciais do Windows falhou (erro {code})."
                )
            return ctypes.string_at(output.pbData, output.cbData)
        finally:
            ctypes.memset(value_buffer, 0, len(value))
            ctypes.memset(entropy_buffer, 0, len(_DPAPI_ENTROPY))
            if output.pbData:
                ctypes.memset(output.pbData, 0, output.cbData)
                self._kernel32.LocalFree(output.pbData)
            if description:
                self._kernel32.LocalFree(description)

    @staticmethod
    def _input_blob(value: bytes) -> tuple[_DATA_BLOB, ctypes.Array]:
        buffer = (ctypes.c_ubyte * len(value)).from_buffer_copy(value)
        blob = _DATA_BLOB(
            len(value),
            ctypes.cast(buffer, ctypes.POINTER(ctypes.c_ubyte)),
        )
        return blob, buffer


class _WindowsProtectedFileBackend:
    """Fallback binario criptografado pela DPAPI do Windows."""

    def __init__(self, path: Path, protector: _DpapiProtector | None = None) -> None:
        self._path = path
        self._protector = protector or _DpapiProtector()

    @property
    def available(self) -> bool:
        return True

    def write(self, target: str, username: str, secret: bytes) -> None:
        username_bytes = username.encode("utf-8")
        record = _PROTECTED_RECORD_HEADER.pack(
            _PROTECTED_RECORD_MAGIC,
            len(username_bytes),
            len(secret),
        ) + username_bytes + secret
        encrypted = self._protector.protect(record)
        self._atomic_write(encrypted)

    def read(self, target: str) -> tuple[str, bytes] | None:
        try:
            encrypted = self._path.read_bytes()
        except FileNotFoundError:
            return None
        except OSError as exc:
            raise CredentialStorageError(
                "Nao foi possivel ler as credenciais protegidas."
            ) from exc
        try:
            record = self._protector.unprotect(encrypted)
            return self._decode_record(record)
        except CredentialStorageError:
            raise
        except (UnicodeDecodeError, ValueError, struct.error) as exc:
            raise CredentialStorageError(
                "As credenciais protegidas estao invalidas."
            ) from exc

    def delete(self, target: str) -> bool:
        try:
            self._path.unlink()
            return True
        except FileNotFoundError:
            return False
        except OSError as exc:
            raise CredentialStorageError(
                "Nao foi possivel remover as credenciais protegidas."
            ) from exc

    def _atomic_write(self, encrypted: bytes) -> None:
        self._path.parent.mkdir(parents=True, exist_ok=True)
        descriptor, temporary_name = tempfile.mkstemp(
            prefix=f"{self._path.stem}_",
            suffix=".tmp",
            dir=str(self._path.parent),
        )
        try:
            with os.fdopen(descriptor, "wb") as stream:
                stream.write(encrypted)
                stream.flush()
                os.fsync(stream.fileno())
            os.replace(temporary_name, self._path)
        finally:
            if os.path.exists(temporary_name):
                os.remove(temporary_name)

    @staticmethod
    def _decode_record(record: bytes) -> tuple[str, bytes]:
        if len(record) < _PROTECTED_RECORD_HEADER.size:
            raise ValueError("Credencial protegida truncada.")
        magic, username_size, secret_size = _PROTECTED_RECORD_HEADER.unpack_from(record)
        expected_size = _PROTECTED_RECORD_HEADER.size + username_size + secret_size
        if magic != _PROTECTED_RECORD_MAGIC or expected_size != len(record):
            raise ValueError("Credencial protegida invalida.")
        split_at = _PROTECTED_RECORD_HEADER.size + username_size
        username = record[_PROTECTED_RECORD_HEADER.size:split_at].decode("utf-8")
        return username, record[split_at:]


class _FallbackCredentialBackend:
    """Usa WinCred; a DPAPI assume quando o cofre nao aceita a operacao."""

    def __init__(
        self,
        primary: _CredentialBackend,
        fallback: _CredentialBackend,
    ) -> None:
        self._primary = primary
        self._fallback = fallback

    @property
    def available(self) -> bool:
        return self._primary.available or self._fallback.available

    def write(self, target: str, username: str, secret: bytes) -> None:
        try:
            self._primary.write(target, username, secret)
        except CredentialStorageError:
            self._fallback.write(target, username, secret)
            return
        try:
            self._fallback.delete(target)
        except CredentialStorageError:
            # Nao mantemos duas copias que poderiam divergir depois.
            try:
                self._primary.delete(target)
            finally:
                raise

    def read(self, target: str) -> tuple[str, bytes] | None:
        # A existencia do fallback indica que ele e a copia mais recente.
        fallback_value = self._fallback.read(target)
        if fallback_value is not None:
            return fallback_value
        return self._primary.read(target)

    def delete(self, target: str) -> bool:
        deleted_fallback = self._fallback.delete(target)
        deleted_primary = self._primary.delete(target)
        return deleted_fallback or deleted_primary


class RhidCredentialService:
    """API da aplicacao para lembrar uma unica conta do RHiD com seguranca."""

    def __init__(
        self,
        backend: _CredentialBackend | None = None,
        target: str = DEFAULT_CREDENTIAL_TARGET,
        protected_file: Path | None = None,
    ) -> None:
        self._backend = (
            backend
            if backend is not None
            else self._default_backend(protected_file or DATA_DIR / "rhid_credentials.dat")
        )
        self._target = target

    @property
    def available(self) -> bool:
        return self._backend.available

    def save(self, email: str, password: str, domain: str = "") -> None:
        email = (email or "").strip()
        domain = (domain or "").strip()
        if not email:
            raise ValueError("Informe o e-mail do RHiD.")
        if not password:
            raise ValueError("Informe a senha do RHiD.")

        secret = _encode_secret(domain, password)
        self._backend.write(self._target, email, secret)

    def load(self) -> RhidCredentials | None:
        stored = self._backend.read(self._target)
        if stored is None:
            return None
        email, secret = stored
        try:
            domain, password = _decode_secret(secret)
        except (UnicodeDecodeError, ValueError, struct.error) as exc:
            raise CredentialStorageError(
                "As credenciais salvas estao invalidas. Remova-as e salve novamente."
            ) from exc
        if not email or not password:
            raise CredentialStorageError(
                "As credenciais salvas estao incompletas. Remova-as e salve novamente."
            )
        return RhidCredentials(email=email, domain=domain, password=password)

    def delete(self) -> bool:
        return self._backend.delete(self._target)

    @staticmethod
    def _default_backend(protected_file: Path) -> _CredentialBackend:
        if sys.platform != "win32":
            return _UnavailableCredentialBackend()
        protected_backend = _WindowsProtectedFileBackend(protected_file)
        try:
            credential_backend = _WindowsCredentialBackend()
        except CredentialStorageUnavailable:
            return protected_backend
        return _FallbackCredentialBackend(credential_backend, protected_backend)


def _encode_secret(domain: str, password: str) -> bytes:
    domain_bytes = domain.encode("utf-8")
    password_bytes = password.encode("utf-8")
    blob = _SECRET_HEADER.pack(
        _SECRET_MAGIC,
        len(domain_bytes),
        len(password_bytes),
    ) + domain_bytes + password_bytes
    if len(blob) > _MAX_CREDENTIAL_BLOB_SIZE:
        raise ValueError("A credencial informada excede o limite seguro do Windows.")
    return blob


def _decode_secret(blob: bytes) -> tuple[str, str]:
    if len(blob) < _SECRET_HEADER.size:
        raise ValueError("Credencial truncada.")
    magic, domain_size, password_size = _SECRET_HEADER.unpack_from(blob)
    if magic != _SECRET_MAGIC:
        raise ValueError("Versao de credencial desconhecida.")
    expected_size = _SECRET_HEADER.size + domain_size + password_size
    if expected_size != len(blob) or expected_size > _MAX_CREDENTIAL_BLOB_SIZE:
        raise ValueError("Tamanho de credencial invalido.")
    split_at = _SECRET_HEADER.size + domain_size
    domain = blob[_SECRET_HEADER.size:split_at].decode("utf-8")
    password = blob[split_at:].decode("utf-8")
    return domain, password
