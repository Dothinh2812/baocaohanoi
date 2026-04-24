from .auth_routes import auth_bp
from .growth_routes import growth_bp
from .inventory_routes import inventory_bp
from .operations_routes import operations_bp
from .quality_routes import quality_bp
from .quangchudong_routes import quangchudong_bp
from .retention_routes import retention_bp
from .sa_outage_routes import sa_outage_bp
from .statistics_routes import statistics_bp

__all__ = [
    'auth_bp',
    'growth_bp',
    'inventory_bp',
    'operations_bp',
    'quality_bp',
    'quangchudong_bp',
    'retention_bp',
    'sa_outage_bp',
    'statistics_bp',
]
