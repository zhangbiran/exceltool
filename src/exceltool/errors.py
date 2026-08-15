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


class FormulaVerificationError(VerificationError):
    def __init__(self, message, report, changes):
        super().__init__(message)
        self.details = {
            "formula_verification": report,
            "unexpected_formula_changes": changes,
            "published": False,
        }


class InputHashMismatchError(TargetError):
    def __init__(self, message, expected, actual):
        super().__init__("%s: expected=%s actual=%s" % (message, expected, actual))
        self.details = {
            "expected_sha256": expected,
            "actual_sha256": actual,
            "published": False,
        }


class PatchOperationError(ExcelToolError):
    def __init__(self, message, code, failed_operation):
        super().__init__(message, code)
        self.details = {
            "operation": "patch",
            "failed_operation": failed_operation,
            "published": False,
        }
