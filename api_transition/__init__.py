# API transition package

__all__ = ["run_full_pipeline"]


def __getattr__(name):
    if name == "run_full_pipeline":
        from api_transition.full_pipeline import run_full_pipeline

        return run_full_pipeline
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
