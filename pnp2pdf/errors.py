class OperationCancelled(Exception):
    """Raised when the user cancels a running build."""


class InputError(ValueError):
    """Raised when the user-supplied parameters cannot describe a job."""
