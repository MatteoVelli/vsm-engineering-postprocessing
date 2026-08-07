class VSMPostProcessingError(Exception):
    """Base exception for the VSM post-processing package."""


class FileImportError(VSMPostProcessingError):
    """Raised when an input file cannot be read safely."""


class DataValidationError(VSMPostProcessingError):
    """Raised when imported data fails strict validation."""


class ConfigurationError(VSMPostProcessingError):
    """Raised when a user configuration file is invalid."""


class ChannelSelectionError(VSMPostProcessingError):
    """Raised when requested channels cannot be selected safely."""


class MathChannelError(VSMPostProcessingError):
    """Raised when configured math channels cannot be calculated or verified safely."""


class StatisticsError(VSMPostProcessingError):
    """Raised when configured statistics cannot be calculated or verified safely."""


class PlottingError(VSMPostProcessingError):
    """Raised when configured plots cannot be rendered safely."""


class ExcelReportError(VSMPostProcessingError):
    """Raised when an Excel engineering report cannot be generated safely."""


class PowerPointReportError(VSMPostProcessingError):
    """Raised when a PowerPoint engineering report cannot be generated safely."""


class PipelineError(VSMPostProcessingError):
    """Raised when the end-to-end processing pipeline cannot complete safely."""
