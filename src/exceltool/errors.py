class ExcelToolError(Exception):
    def __init__(self, message, code=5):
        super().__init__(message)
        self.message = message
        self.code = code


class TargetError(ExcelToolError):
    def __init__(self, message):
        super().__init__(message, 3)


class UnsupportedError(ExcelToolError):
    def __init__(self, message):
        super().__init__(message, 4)


class VerificationError(ExcelToolError):
    def __init__(self, message):
        super().__init__(message, 6)


class PatchOperationError(ExcelToolError):
    def __init__(self, message, code, failed_operation):
        super().__init__(message, code)
        self.details = {
            "operation": "patch",
            "failed_operation": failed_operation,
            "published": False,
        }
