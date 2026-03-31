import datetime
import re

def para_minutos(valor):
    if valor is None or valor == "":
        return 0
    if isinstance(valor, datetime.timedelta):
        return int(round(valor.total_seconds() / 60))
    if isinstance(valor, datetime.datetime):
        if valor.year == 1900:
            return (valor.day * 24 + valor.hour) * 60 + valor.minute
        return valor.hour * 60 + valor.minute
    if isinstance(valor, datetime.time):
        return valor.hour * 60 + valor.minute
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return int(round(valor * 1440))

    texto = str(valor).strip()
    negativo = texto.startswith("-")
    if negativo:
        texto = texto[1:].strip()

    m = re.match(r"(\d+)\s+day[s]?,\s*(\d+):(\d{2})(?::(\d{2}))?$", texto)
    if m:
        dias = int(m.group(1))
        horas = int(m.group(2))
        minutos = int(m.group(3))
        total = dias * 1440 + horas * 60 + minutos
        return -total if negativo else total

    m = re.match(r"(\d{4})-(\d{2})-(\d{2})\s+(\d+):(\d{2})(?::(\d{2}))?$", texto)
    if m:
        ano = int(m.group(1))
        dia = int(m.group(3))
        horas = int(m.group(4))
        minutos = int(m.group(5))
        if ano == 1900:
            total = (dia * 24 + horas) * 60 + minutos
        else:
            total = horas * 60 + minutos
        return -total if negativo else total

    partes = texto.split(":")
    if len(partes) >= 2:
        horas = int(partes[0])
        minutos = int(partes[1])
        total = horas * 60 + minutos
        return -total if negativo else total
    return 0

def formatar_horas(total_minutos):
    sinal = "-" if total_minutos < 0 else ""
    total_minutos = abs(int(total_minutos))
    horas = total_minutos // 60
    minutos = total_minutos % 60
    return f"{sinal}{horas}:{minutos:02d}"
