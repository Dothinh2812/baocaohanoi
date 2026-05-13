import os


def apply_native_thread_limits():
    default_threads = os.getenv('DASHV4_NATIVE_THREADS', '1')
    for name in (
        'OMP_NUM_THREADS',
        'OPENBLAS_NUM_THREADS',
        'MKL_NUM_THREADS',
        'NUMEXPR_NUM_THREADS',
        'VECLIB_MAXIMUM_THREADS',
        'BLIS_NUM_THREADS',
    ):
        os.environ.setdefault(name, default_threads)


apply_native_thread_limits()
