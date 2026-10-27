"""Diagnostic presentation policy with backward-compatible string exports."""
from collections.abc import Iterable


class Diagnostic(str):
    """Keep existing text consumers while carrying explicit presentation policy."""

    def __new__(cls, message: str, *, code: str = 'engineering_diagnostic',
                severity: str = 'warning', user_visible: bool = True):
        if severity not in {'info', 'warning', 'error'}:
            raise ValueError('Unsupported diagnostic severity: ' + severity)
        value = super().__new__(cls, message)
        value.code = code
        value.severity = severity
        value.user_visible = user_visible
        return value

    def to_dict(self):
        return {'code': self.code, 'severity': self.severity,
                'user_visible': self.user_visible, 'message': str(self)}


def diagnostic_records(messages: Iterable[str]):
    """Export policy explicitly; legacy diagnostics remain visible warnings."""
    return [as_diagnostic(message).to_dict() for message in messages]


def as_diagnostic(message: str) -> Diagnostic:
    return message if isinstance(message, Diagnostic) else Diagnostic(message)


def visible_diagnostics(messages: Iterable[str]):
    return [message for message in messages if as_diagnostic(message).user_visible]
