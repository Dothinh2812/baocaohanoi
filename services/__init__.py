from .quangchudong_cache import get_quangchudong_cache, initialize_quangchudong_cache


def initialize_background_services(app, *, warm=False):
    initialize_quangchudong_cache(app, warm=warm)


__all__ = [
    'get_quangchudong_cache',
    'initialize_background_services',
    'initialize_quangchudong_cache',
]
