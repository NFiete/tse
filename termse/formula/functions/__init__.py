"""Importing this package registers every built-in function."""
from . import (  # noqa: F401
    database_fns, date_fns, dist_fns, eng_fns, group_fns, lambda_fns, logical_fns, lookup_fns, math_fns,
    regress_fns, stats_fns, text_fns,
)


def registered():
    """Names of all registered functions."""
    from ..registry import FUNCTIONS

    return FUNCTIONS
