"""AKS migration engine.

Deterministic, configuration-driven migration of legacy STADLER AKS workbooks
into the SAP-era AKS V2 template. The engine never guesses: a legacy value with
no defined mapping is flagged for human review, and an unresolvable column stops
the run.

The engine is independent of any user interface. Import what you need from here:

    from migration import analyse_migration, run_migration

    analysis = analyse_migration("P026393_AKS_04.xlsb")
    if analysis.ok and analysis.review_count == 0:
        result = run_migration(analysis)
"""

from .api import (
    AnalysisResult,
    EnvironmentReport,
    MigrationResult,
    OutputVerification,
    SourceInspection,
    analyse_migration,
    check_environment,
    cleanup_staged_uploads,
    describe_configuration,
    inspect_source_file,
    planned_output_name,
    previous_outputs,
    run_migration,
    stage_configuration_file,
    stage_source,
)
from .model import MigrationError
from .paths import (
    APP_ROOT,
    LOGS_DIR,
    OUTPUTS_DIR,
    default_mapping,
    default_template,
    version_label,
)
from .reporting import (
    REVIEW_STATUSES,
    STATUS_ERROR,
    STATUS_OK,
    STATUS_REVIEW,
    STATUS_SKIPPED,
    STATUS_WARNING,
    TABLE_COLUMNS,
)

__version__ = "2.0.0"

__all__ = [
    "APP_ROOT",
    "LOGS_DIR",
    "OUTPUTS_DIR",
    "REVIEW_STATUSES",
    "STATUS_ERROR",
    "STATUS_OK",
    "STATUS_REVIEW",
    "STATUS_SKIPPED",
    "STATUS_WARNING",
    "TABLE_COLUMNS",
    "AnalysisResult",
    "EnvironmentReport",
    "MigrationError",
    "MigrationResult",
    "OutputVerification",
    "SourceInspection",
    "__version__",
    "analyse_migration",
    "check_environment",
    "cleanup_staged_uploads",
    "default_mapping",
    "default_template",
    "describe_configuration",
    "inspect_source_file",
    "planned_output_name",
    "previous_outputs",
    "run_migration",
    "stage_configuration_file",
    "stage_source",
    "version_label",
]
