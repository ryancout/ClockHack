class AppError(Exception):
    pass

class ArquivoInvalidoError(AppError):
    pass

class ColunaObrigatoriaError(AppError):
    pass

class ValidacaoNegocioError(AppError):
    pass
