import io
import os
import sqlite3
from datetime import date, timedelta

import pandas as pd
from flask import Blueprint, jsonify, request, send_file

from config import BASE_DATA_PATH
from ticket_tracking import get_statistics_report


statistics_bp = Blueprint('statistics', __name__)


def _serialize_statistics_dataframe(df):
    serialized = df.copy()
    for col in serialized.columns:
        if pd.api.types.is_datetime64_any_dtype(serialized[col]):
            serialized[col] = serialized[col].dt.strftime('%d/%m/%Y')
        elif pd.api.types.is_period_dtype(serialized[col]):
            serialized[col] = serialized[col].astype(str)
    return serialized.fillna(0)


def _sum_column(df, column_name):
    if df.empty or column_name not in df.columns:
        return 0
    return int(df[column_name].sum())


def _mean_column(df, column_name):
    if df.empty or column_name not in df.columns:
        return 0
    return int(df[column_name].mean())


def _summary_period_payload(df):
    summary = {
        'so_phieu_nhan': _sum_column(df, 'so_phieu_nhan'),
        'so_phieu_xu_ly_xong': _sum_column(df, 'so_phieu_xu_ly_xong'),
        'so_phieu_ton': _mean_column(df, 'so_phieu_ton_cuoi'),
    }

    nhan = summary['so_phieu_nhan']
    xuly = summary['so_phieu_xu_ly_xong']
    summary['ty_le_xu_ly'] = round((xuly / nhan) * 100, 1) if nhan > 0 else 0
    return summary


def _database_info():
    db_path = os.path.join(BASE_DATA_PATH, 'database', 'brcd.db')
    if not os.path.exists(db_path):
        return None

    with sqlite3.connect(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute('SELECT COUNT(*) FROM ticket_history WHERE trang_thai = "DANG_XU_LY"')
        tickets_in_progress = cursor.fetchone()[0]

        cursor.execute('SELECT MAX(snapshot_time) FROM ticket_snapshots')
        last_update = cursor.fetchone()[0]

    return {
        'tickets_in_progress': tickets_in_progress,
        'last_update': last_update,
    }


@statistics_bp.route('/api/ticket-statistics')
def get_ticket_statistics():
    try:
        start_date = request.args.get('start_date') or (date.today() - timedelta(days=7)).isoformat()
        end_date = request.args.get('end_date') or date.today().isoformat()
        ticket_type = request.args.get('ticket_type')
        doi_vt = request.args.get('doi_vt')
        nhan_vien = request.args.get('nhan_vien')
        group_by = request.args.get('group_by', 'day')

        df = get_statistics_report(
            start_date=start_date,
            end_date=end_date,
            ticket_type=ticket_type or None,
            doi_vt=doi_vt or None,
            nhan_vien=nhan_vien or None,
            group_by=group_by,
        )
        serialized = _serialize_statistics_dataframe(df)

        return jsonify({
            'filters': {
                'start_date': start_date,
                'end_date': end_date,
                'ticket_type': ticket_type,
                'doi_vt': doi_vt,
                'nhan_vien': nhan_vien,
                'group_by': group_by,
            },
            'columns': serialized.columns.tolist(),
            'data': serialized.to_dict('records'),
            'total_records': len(serialized),
        })
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi lấy thống kê: {exc}'}), 500


@statistics_bp.route('/api/ticket-statistics-summary')
def get_ticket_statistics_summary():
    try:
        today = date.today()
        week_start = today - timedelta(days=today.weekday())
        month_start = today.replace(day=1)

        df_today = get_statistics_report(start_date=today.isoformat(), end_date=today.isoformat())
        df_week = get_statistics_report(start_date=week_start.isoformat(), end_date=today.isoformat())
        df_month = get_statistics_report(start_date=month_start.isoformat(), end_date=today.isoformat())

        summary = {
            'today': _summary_period_payload(df_today),
            'this_week': _summary_period_payload(df_week),
            'this_month': _summary_period_payload(df_month),
        }

        database_info = _database_info()
        if database_info:
            summary['database_info'] = database_info

        return jsonify(summary)
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi lấy tổng quan: {exc}'}), 500


@statistics_bp.route('/api/ticket-history/<ticket_id>')
def get_ticket_history(ticket_id):
    try:
        ticket_type = request.args.get('ticket_type', 'BRCD')
        db_path = os.path.join(BASE_DATA_PATH, 'database', 'brcd.db')
        if not os.path.exists(db_path):
            return jsonify({'error': 'Database không tồn tại'}), 404

        with sqlite3.connect(db_path) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute(
                '''
                SELECT *
                FROM ticket_history
                WHERE ticket_id = ? AND ticket_type = ?
                ORDER BY ngay_nhan DESC
                ''',
                (ticket_id, ticket_type),
            )
            rows = cursor.fetchall()

        if not rows:
            return jsonify({'error': 'Không tìm thấy lịch sử cho phiếu này'}), 404

        history = [dict(row) for row in rows]
        return jsonify({
            'ticket_id': ticket_id,
            'ticket_type': ticket_type,
            'history': history,
            'total_records': len(history),
        })
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi lấy lịch sử: {exc}'}), 500


@statistics_bp.route('/api/ticket-trend')
def get_ticket_trend():
    try:
        days = request.args.get('days', default=30, type=int)
        ticket_type = request.args.get('ticket_type')

        end_date = date.today()
        start_date = end_date - timedelta(days=days)

        df = get_statistics_report(
            start_date=start_date.isoformat(),
            end_date=end_date.isoformat(),
            ticket_type=ticket_type or None,
            group_by='day',
        )

        if df.empty or 'ngay' not in df.columns:
            return jsonify({'labels': [], 'datasets': []})

        df = df.copy()
        df['ngay'] = pd.to_datetime(df['ngay'])
        daily = df.groupby('ngay').agg({
            'so_phieu_nhan': 'sum',
            'so_phieu_xu_ly_xong': 'sum',
            'so_phieu_ton_cuoi': 'mean',
        }).reset_index()
        daily['ngay'] = daily['ngay'].dt.strftime('%d/%m/%Y')

        return jsonify({
            'labels': daily['ngay'].tolist(),
            'datasets': [
                {
                    'label': 'Phiếu nhận',
                    'data': daily['so_phieu_nhan'].tolist(),
                },
                {
                    'label': 'Phiếu xử lý xong',
                    'data': daily['so_phieu_xu_ly_xong'].tolist(),
                },
                {
                    'label': 'Phiếu tồn',
                    'data': daily['so_phieu_ton_cuoi'].tolist(),
                },
            ],
        })
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi lấy xu hướng: {exc}'}), 500


@statistics_bp.route('/download/excel-statistics')
def download_excel_statistics():
    try:
        start_date = request.args.get('start_date', (date.today() - timedelta(days=30)).isoformat())
        end_date = request.args.get('end_date', date.today().isoformat())
        ticket_type = request.args.get('ticket_type')

        df = get_statistics_report(
            start_date=start_date,
            end_date=end_date,
            ticket_type=ticket_type or None,
        )

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='ThongKe', index=False)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name=f'BaoCaoThongKe_{start_date}_{end_date}.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi tạo file Excel: {exc}'}), 500
