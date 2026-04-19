# -*- coding: utf-8 -*-
"""Helpers va processors cho luong xu ly du lieu trong api_transition."""

__all__ = [
    "PROCESSOR_GROUPS",
    "PROCESSOR_TASKS",
    "ProcessorRunResult",
    "ProcessorTask",
    "list_processors",
    "run_all_processors",
]


def __getattr__(name):
    if name in __all__:
        from api_transition.processors import runner

        return getattr(runner, name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
