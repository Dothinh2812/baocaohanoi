# API transition package

__all__ = ["RuntimeContext", "load_runtime_context", "run_full_pipeline"]


def __getattr__(name):
    if name == "run_full_pipeline":
        from api_transition.full_pipeline import run_full_pipeline

        return run_full_pipeline
    if name == "RuntimeContext":
        from api_transition.runtime_config import RuntimeContext

        return RuntimeContext
    if name == "load_runtime_context":
        from api_transition.runtime_config import load_runtime_context

        return load_runtime_context
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
