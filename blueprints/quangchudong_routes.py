from flask import Blueprint, jsonify, render_template, session

from auth import get_user_by_username, login_required
from services import get_quangchudong_cache


quangchudong_bp = Blueprint('quangchudong', __name__)


@quangchudong_bp.route('/quangchudong')
@login_required
def page_quangchudong():
    user = get_user_by_username(session['username'])
    return render_template('quangchudong.html', current_user=user, active_page='quangchudong')


@quangchudong_bp.route('/api/quangchudong/active')
@login_required
def get_active_outages():
    return jsonify(get_quangchudong_cache().get_active_payload())


@quangchudong_bp.route('/api/quangchudong/dashboard')
@login_required
def get_dashboard_payload():
    return jsonify(get_quangchudong_cache().get_dashboard_payload())


@quangchudong_bp.route('/api/quangchudong/recovered')
@login_required
def get_recovered_alerts():
    return jsonify(get_quangchudong_cache().get_recovered_payload())


@quangchudong_bp.route('/api/quangchudong/wide-area')
@login_required
def get_wide_area_outages():
    return jsonify(get_quangchudong_cache().get_wide_area_payload())


@quangchudong_bp.route('/api/quangchudong/wide-area-groups')
@login_required
def get_wide_area_groups():
    return jsonify(get_quangchudong_cache().get_wide_area_groups_payload())


@quangchudong_bp.route('/api/quangchudong/exclusion-list')
@login_required
def get_exclusion_list():
    return jsonify(get_quangchudong_cache().get_exclusion_list_payload())


@quangchudong_bp.route('/api/quangchudong/pattern-exclusions')
@login_required
def get_pattern_exclusions():
    return jsonify(get_quangchudong_cache().get_pattern_exclusions_payload())


@quangchudong_bp.route('/api/quangchudong/port-down-groups')
@login_required
def get_port_down_groups():
    return jsonify(get_quangchudong_cache().get_port_down_groups_payload())


@quangchudong_bp.route('/api/quangchudong/stats')
@login_required
def get_outage_stats():
    return jsonify(get_quangchudong_cache().get_stats_payload())
