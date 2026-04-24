from repositories.sqlite_runtime import read_sql_dataframe, read_sql_rows


def load_ttvt_son_tay_tong_hop_rows(unit_name):
    return read_sql_rows(
        '''
        SELECT
            ngay_du_lieu,
            ttvt,
            nhom_du_lieu,
            nhom_chi_tieu,
            don_vi,
            ten_chi_so,
            gia_tri_so,
            chi_tieu_bsc,
            nguon_view
        FROM v_dashboard_ttvt_son_tay_tong_hop_moi_nhat
        WHERE ttvt = ?
        ORDER BY nhom_du_lieu, nhom_chi_tieu, ten_chi_so
        ''',
        (unit_name,),
    )


def load_cau_hinh_tu_dong_detail_df(unit_name):
    return read_sql_dataframe(
        '''
        SELECT
            serial_number AS "Serial Number",
            ma_thue_bao AS "Mã thuê bao",
            loai_hop_dong AS "Loại hợp đồng",
            loai_cau_hinh AS "Loại cấu hình",
            trang_thai AS "Trang thái",
            trang_thai_chuan_hoa AS "Trang thái chuẩn hóa",
            ma_loi AS "Mã lỗi",
            thoi_gian_cap_nhat AS "Thời gian cập nhật",
            ttvt AS "Trung tâm Viễn thông",
            doi_vien_thong AS "Đội Viễn thông",
            ma_nhan_vien AS "Mã nhân viên",
            nvkt AS "NVKT"
        FROM v_cau_hinh_tu_dong_chi_tiet_moi_nhat
        WHERE ttvt = ?
        ORDER BY "Thời gian cập nhật" DESC, "NVKT", "Mã thuê bao"
        ''',
        (unit_name,),
    )


def load_cau_hinh_tu_dong_tong_hop_df(unit_name):
    return read_sql_dataframe(
        '''
        SELECT
            ma_bao_cao AS "Mã báo cáo",
            ngay_du_lieu AS "Ngày dữ liệu",
            ttvt AS "Trung tâm Viễn thông",
            don_vi AS "Đơn vị",
            loai_dong AS "Loại dòng",
            tong_hop_dong AS "Tổng số",
            khong_thuc_hien_cau_hinh_tu_dong AS "Không thực hiện cấu hình tự động",
            da_day_cau_hinh_tu_dong AS "Đã đẩy cấu hình tự động",
            khong_day_do_loi_he_thong AS "Không đẩy do lỗi hệ thống",
            khong_day_do_tbi_da_co_cau_hinh AS "Không đẩy do TBI đã có cấu hình",
            cau_hinh_thanh_cong AS "Cấu hình thành công",
            ty_le_day_tu_dong AS "Tỷ lệ đẩy tự động",
            ty_le_tbi_da_co_cau_hinh AS "Tỷ lệ TBI đã có cấu hình",
            ty_le_cau_hinh_thanh_cong AS "Tỷ lệ cấu hình thành công"
        FROM v_cau_hinh_tu_dong_tong_hop_moi_nhat
        WHERE ttvt = ?
        ORDER BY ma_bao_cao, loai_dong, don_vi
        ''',
        (unit_name,),
    )


def load_cau_hinh_tu_dong_team_detail_df(unit_name):
    df = read_sql_dataframe(
        '''
        SELECT
            "Trung tâm Viễn thông",
            "Đội Viễn thông",
            "Tổng hợp đồng",
            "Lắp mới",
            "Thay thế",
            "Cấu hình WAN",
            "Cấu hình WiFi",
            "Thành công",
            "Thất bại",
            "Chưa có trạng thái",
            ROUND(COALESCE("Tỷ lệ thành công (%)", 0), 2) AS "Tỷ lệ thành công (%)",
            ROUND(COALESCE("Tỷ lệ thất bại (%)", 0), 2) AS "Tỷ lệ thất bại (%)"
        FROM v_cau_hinh_tu_dong_chi_tiet_th_theo_to
        WHERE "Trung tâm Viễn thông" = ?
        ORDER BY "Tổng hợp đồng" DESC, "Đội Viễn thông"
        ''',
        (unit_name,),
    )
    if df.empty:
        return df

    df = df.reset_index(drop=True)
    df.insert(0, 'STT', range(1, len(df) + 1))
    return df


def load_cau_hinh_tu_dong_team_summary_df():
    df = read_sql_dataframe(
        '''
        SELECT
            "Đơn vị" AS "Đội Viễn thông",
            "Tổng số",
            "Cấu hình thành công" AS "Thành công",
            "Không đẩy do lỗi hệ thống" + "Không đẩy do TBI đã có cấu hình" AS "Thất bại",
            "Không thực hiện cấu hình tự động" AS "Chưa có trạng thái",
            ROUND(COALESCE("Tỷ lệ cấu hình thành công", 0), 2) AS "Tỷ lệ thành công (%)"
        FROM v_cau_hinh_tu_dong_son_tay_tong_hop_theo_to_moi_nhat
        ORDER BY "Tổng số" DESC, "Đội Viễn thông"
        '''
    )
    if df.empty:
        return df

    df = df.reset_index(drop=True)
    df.insert(0, 'STT', range(1, len(df) + 1))
    return df


def load_cau_hinh_tu_dong_nvkt_summary_df(unit_name='TTVT Sơn Tây'):
    df = read_sql_dataframe(
        '''
        SELECT
            "Trung tâm Viễn thông",
            "Đội Viễn thông",
            "NVKT",
            "Tổng hợp đồng" AS "Tổng số",
            "Lắp mới",
            "Thay thế",
            "Cấu hình WAN",
            "Cấu hình WiFi",
            "Thành công",
            "Thất bại",
            "Chưa có trạng thái",
            ROUND(COALESCE("Tỷ lệ thành công (%)", 0), 2) AS "Tỷ lệ thành công (%)",
            ROUND(COALESCE("Tỷ lệ thất bại (%)", 0), 2) AS "Tỷ lệ thất bại (%)"
        FROM v_cau_hinh_tu_dong_chi_tiet_th_theo_nvkt
        WHERE "Trung tâm Viễn thông" = ?
        ORDER BY "Tổng số" DESC, "Đội Viễn thông", "NVKT"
        ''',
        (unit_name,),
    )
    if df.empty:
        return df

    df = df.reset_index(drop=True)
    df.insert(0, 'STT', range(1, len(df) + 1))
    return df


def load_cau_hinh_tu_dong_error_summary_df():
    return read_sql_dataframe(
        '''
        SELECT
            "Mã lỗi",
            "Số lượng"
        FROM v_cau_hinh_tu_dong_son_tay_tong_hop_loi_moi_nhat
        ORDER BY "Số lượng" DESC, "Mã lỗi"
        '''
    )


def load_ket_qua_tiep_thi_nv_df():
    return read_sql_dataframe(
        '''
        SELECT
            "Đơn vị",
            "Mã NV",
            "Tên nhân viên" AS "Tên NV",
            "Dịch vụ BRCD",
            "Dịch vụ MyTV",
            "Tổng"
        FROM v_tiep_thi_ui_chi_tiet_moi_nhat
        ORDER BY "Đơn vị", "Tên NV"
        '''
    )


def load_thu_hoi_tong_hop_df():
    return read_sql_dataframe(
        '''
        SELECT
            "TTVT",
            "Đội VT",
            "NVKT địa bàn giao",
            "Loại vật tư",
            "Trạng thái thu hồi",
            "Số lượng"
        FROM v_thu_hoi_ui_tong_hop_moi_nhat
        ORDER BY "TTVT", "Đội VT", "NVKT địa bàn giao", "Loại vật tư", "Trạng thái thu hồi"
        '''
    )


def load_thu_hoi_chi_tiet_df():
    return read_sql_dataframe(
        '''
        SELECT
            "TTVT",
            "Đội VT",
            "Khu vực",
            "NVKT địa bàn giao",
            "Trạng thái thu hồi",
            "Loại vật tư",
            "Loại phiếu",
            "Mã MEN",
            "Mã thuê bao",
            "Tên thuê bao",
            "Địa chỉ khách hàng",
            "Nhân viên khóa",
            "Nhân viên thu",
            "Nhân viên nhập kho",
            "Ngày khóa",
            "Ngày hoàn công",
            "Ngày hoàn ứng",
            "Tên thiết bị",
            "Serial"
        FROM v_thu_hoi_ui_chi_tiet_moi_nhat
        ORDER BY "TTVT", "Đội VT", "NVKT địa bàn giao", "Ngày hoàn công" DESC, "Mã thuê bao"
        '''
    )


def load_ghtt_don_vi_df():
    return read_sql_dataframe(
        '''
        SELECT
            "Đơn vị",
            "Hoàn thành T",
            "Giao NVKT T" AS "Giao T",
            "Tỷ lệ T",
            "Hoàn thành T+1",
            "Giao NVKT T+1" AS "Giao T+1",
            "Tỷ lệ T+1",
            "SL GHTT >=6T" AS "Số lượng GHTT > 6 tháng",
            "Hoàn thành >=6T T+1" AS "Hoàn thành > 6 tháng T+1",
            "Tỷ lệ >=6T T+1" AS "Tỷ lệ > 6 tháng T+1",
            "Tỷ lệ Tổng" AS "Tỷ lệ tổng"
        FROM v_ghtt_sontay_kq_sontay
        ORDER BY "Đơn vị"
        '''
    )


def load_ghtt_hni_df():
    return read_sql_dataframe(
        '''
        SELECT
            "Đơn vị",
            "Hoàn thành T",
            "Giao NVKT T" AS "Giao T",
            "Tỷ lệ T",
            "Hoàn thành T+1",
            "Giao NVKT T+1" AS "Giao T+1",
            "Tỷ lệ T+1",
            "SL GHTT >=6T" AS "Số lượng GHTT > 6 tháng",
            "Hoàn thành >=6T T+1" AS "Hoàn thành > 6 tháng T+1",
            "Tỷ lệ >=6T T+1" AS "Tỷ lệ > 6 tháng T+1",
            "Tỷ lệ Tổng" AS "Tỷ lệ tổng"
        FROM v_ghtt_hni_kq_hni
        ORDER BY "Đơn vị"
        '''
    )


def load_ghtt_nvkt_df(unit_name):
    return read_sql_dataframe(
        '''
        SELECT
            "Đơn vị",
            "NVKT",
            "Hoàn thành T",
            "Giao NVKT T" AS "Giao T",
            "Tỷ lệ T",
            "Hoàn thành T+1",
            "Giao NVKT T+1" AS "Giao T+1",
            "Tỷ lệ T+1",
            "SL GHTT >=6T" AS "Số lượng GHTT > 6 tháng",
            "Hoàn thành >=6T T+1" AS "Hoàn thành > 6 tháng T+1",
            "Tỷ lệ >=6T T+1" AS "Tỷ lệ > 6 tháng T+1",
            "Tỷ lệ Tổng" AS "Tỷ lệ tổng"
        FROM ghtt_ghtt_nvktdb_report_kq_nvktdb
        WHERE "TTVT" = ?
        ORDER BY "Đơn vị", "NVKT"
        ''',
        (unit_name,),
    )


def load_nvkt_tong_hop_da_nguon_df():
    return read_sql_dataframe(
        '''
        SELECT
            nvkt_hoac_ten_nv,
            to_doi_hoac_don_vi,
            ttvt,
            ma_nv,
            c11_tong_phieu,
            c11_so_phieu_dat,
            c11_ty_le,
            c12_sm1_so_phieu_hll,
            c12_sm1_so_phieu_bao_hong,
            c12_sm1_ty_le_hll,
            c14_tong_phieu_ks_thanh_cong,
            c14_tong_phieu_khl,
            c14_ty_le_hl,
            i15_k1_so_tb_shc,
            i15_k1_so_tb_quan_ly,
            i15_k1_ty_le_shc,
            i15_k2_so_tb_shc,
            i15_k2_so_tb_quan_ly,
            i15_k2_ty_le_shc,
            ghtt_hoan_thanh_t,
            ghtt_giao_nvkt_t,
            ghtt_ty_le_t,
            ghtt_hoan_thanh_t1,
            ghtt_giao_nvkt_t1,
            ghtt_ty_le_t1,
            ghtt_sl_6t,
            ghtt_hoan_thanh_6t_t1,
            ghtt_ty_le_6t_t1,
            ghtt_ty_le_tong,
            kpi_c11_sm1,
            kpi_c11_sm2,
            kpi_c11_ty_le_dat_yeu_cau,
            kpi_c11_sm3,
            kpi_c11_sm4,
            kpi_c11_ty_le_dung_hen,
            kpi_c11_chi_tieu_bsc,
            kpi_c12_sm1,
            kpi_c12_sm2,
            kpi_c12_ty_le_lap_lai,
            kpi_c12_sm3,
            kpi_c12_sm4,
            kpi_c12_ty_le_su_co,
            kpi_c12_chi_tieu_bsc,
            kqtt_brcd,
            kqtt_mytv,
            kqtt_tong
        FROM v_nvkt_tong_hop_da_nguon
        ORDER BY COALESCE(to_doi_hoac_don_vi, ''), COALESCE(nvkt_hoac_ten_nv, '')
        '''
    )


def load_nvkt_tong_hop_da_nguon_by_name_df(nvkt_name):
    return read_sql_dataframe(
        '''
        SELECT *
        FROM v_nvkt_tong_hop_da_nguon
        WHERE TRIM(COALESCE(nvkt_hoac_ten_nv, '')) = TRIM(?)
        ''',
        (nvkt_name,),
    )


def load_don_vi_tong_hop_da_nguon_df():
    return read_sql_dataframe(
        '''
        SELECT *
        FROM v_don_vi_tong_hop_da_nguon
        '''
    )


def load_bsc_kpi_cac_to_df():
    return read_sql_dataframe(
        '''
        SELECT *
        FROM v_chi_tieu_bsc_kpi_cac_to
        '''
    )


def load_chat_luong_dashboard_df():
    raise NotImplementedError('Use dedicated C1 summary loaders instead of a UNION with different column counts.')


def load_v_chi_tieu_c_c1_1_report_th_c1_1_df():
    return read_sql_dataframe('SELECT * FROM v_chi_tieu_c_c1_1_report_th_c1_1 ORDER BY "Đơn vị"')


def load_v_chi_tieu_c_c1_1_chitiet_report_chi_tiet_df():
    return read_sql_dataframe('SELECT * FROM v_chi_tieu_c_c1_1_chitiet_report_chi_tiet ORDER BY "TEN_DOI", "NVKT"')


def load_v_chi_tieu_c_c1_1_chitiet_report_chi_tieu_ko_hen_18h_df():
    return read_sql_dataframe(
        'SELECT * FROM v_chi_tieu_c_c1_1_chitiet_report_chi_tieu_ko_hen_18h ORDER BY "TEN_DOI", "NVKT"'
    )


def load_v_chi_tieu_c_c1_2_report_th_c1_2_df():
    return read_sql_dataframe('SELECT * FROM v_chi_tieu_c_c1_2_report_th_c1_2 ORDER BY "Đơn vị"')


def load_v_chi_tieu_c_c1_2_chitiet_sm1_report_th_sm1c12_hll_thang_df():
    return read_sql_dataframe(
        'SELECT * FROM v_chi_tieu_c_c1_2_chitiet_sm1_report_th_sm1c12_hll_thang ORDER BY "TEN_DOI", "NVKT"'
    )


def load_v_chi_tieu_c_c1_3_report_th_c1_3_df():
    return read_sql_dataframe('SELECT * FROM v_chi_tieu_c_c1_3_report_th_c1_3 ORDER BY "Đơn vị"')


def load_v_chi_tieu_c_c1_4_report_th_c1_4_df():
    return read_sql_dataframe('SELECT * FROM v_chi_tieu_c_c1_4_report_th_c1_4 ORDER BY "Đơn vị"')


def load_v_chi_tieu_c_c1_4_chitiet_report_th_hl_nvkt_df():
    return read_sql_dataframe(
        'SELECT * FROM v_chi_tieu_c_c1_4_chitiet_report_th_hl_nvkt ORDER BY "DOIVT", "NVKT"'
    )


def load_v_chi_tieu_c_c1_5_report_th_c1_5_df():
    return read_sql_dataframe('SELECT * FROM v_chi_tieu_c_c1_5_report_th_c1_5 ORDER BY "Đơn vị"')


def load_c1_1_summary_df():
    return read_sql_dataframe('SELECT * FROM v_ui_c1_1_tong_hop_moi_nhat ORDER BY "Đơn vị"')


def load_c1_2_summary_df():
    return read_sql_dataframe('SELECT * FROM v_ui_c1_2_tong_hop_moi_nhat ORDER BY "Đơn vị"')


def load_c1_3_summary_df():
    return read_sql_dataframe('SELECT * FROM v_ui_c1_3_tong_hop_moi_nhat ORDER BY "Đơn vị"')


def load_c1_4_summary_df():
    return read_sql_dataframe('SELECT * FROM v_ui_c1_4_tong_hop_moi_nhat ORDER BY "Đơn vị"')


def load_c11_nvkt_df(moc_gio):
    view_name_map = {
        'tong': 'v_ui_c11_nvkt_tong_moi_nhat',
        '15h': 'v_ui_c11_nvkt_15h_moi_nhat',
        '16h': 'v_ui_c11_nvkt_16h_moi_nhat',
        '17h': 'v_ui_c11_nvkt_17h_moi_nhat',
        '18h': 'v_ui_c11_nvkt_18h_moi_nhat',
    }
    view_name = view_name_map[moc_gio]
    return read_sql_dataframe(
        f'''
        SELECT
            ngay_du_lieu,
            TEN_DOI,
            NVKT,
            "Tổng phiếu",
            "Số phiếu đạt",
            "Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn"
        FROM {view_name}
        ORDER BY TEN_DOI, NVKT
        '''
    )


def load_c12_nvkt_df():
    return read_sql_dataframe(
        '''
        SELECT
            ngay_du_lieu,
            TEN_DOI,
            NVKT,
            "Số phiếu HLL",
            "Số phiếu báo hỏng",
            "Tỉ lệ HLL tháng (2.5%)"
        FROM v_ui_c12_repeat_failure_nvkt_moi_nhat
        ORDER BY TEN_DOI, NVKT
        '''
    )


def load_c14_tong_hop_df():
    return read_sql_dataframe(
        '''
        SELECT
            don_vi AS "Đơn vị",
            tong_phieu AS "Tổng phiếu",
            so_luong_da_khao_sat AS "Số lượng đã khảo sát",
            so_luong_khao_sat_thanh_cong AS "Số lượng khảo sát thành công",
            so_luong_khach_hang_hai_long AS "Số lượng khách hàng hài lòng",
            khong_hai_long_ky_thuat_phuc_vu AS "Không hài lòng kỹ thuật phục vụ",
            ty_le_hai_long_ky_thuat_phuc_vu AS "Tỷ lệ hài lòng kỹ thuật phục vụ",
            khong_hai_long_ky_thuat_dich_vu AS "Không hài lòng kỹ thuật dịch vụ",
            ty_le_hai_long_ky_thuat_dich_vu AS "Tỷ lệ hài lòng kỹ thuật dịch vụ",
            tong_phieu_hai_long_ky_thuat AS "Tổng phiếu hài lòng kỹ thuật",
            ty_le_khach_hang_hai_long AS "Tỷ lệ khách hàng hài lòng",
            diem_bsc AS "Điểm BSC"
        FROM v_c14_tong_hop_moi_nhat
        ORDER BY "Đơn vị"
        '''
    )


def load_c14_nvkt_df():
    return read_sql_dataframe(
        '''
        SELECT
            ngay_du_lieu,
            DOIVT,
            NVKT,
            "Tổng phiếu KS thành công",
            "Tổng phiếu KHL",
            "Tỉ lệ HL NVKT (%)"
        FROM v_ui_c14_nvkt_moi_nhat
        ORDER BY DOIVT, NVKT
        '''
    )


def load_i15_dashboard_df():
    return read_sql_dataframe(
        '''
        SELECT
            ngay_du_lieu AS "Ngày dữ liệu",
            k_suffix AS "Kỳ",
            doi_one AS "Đơn vị",
            nvkt_db_normalized AS "NVKT",
            tong_so_hien_tai AS "Tổng số hiện tại",
            so_tang_moi AS "Số tăng mới",
            so_giam_het AS "Số giảm/hết",
            so_van_con AS "Số vẫn còn",
            so_tb_quan_ly AS "Số TB quản lý",
            ty_le_shc AS "Tỷ lệ SHC (%)",
            cap_du_lieu AS "Cấp dữ liệu",
            nhom_du_lieu AS "Nhóm dữ liệu",
            nhom_chi_tieu AS "Nhóm chỉ tiêu",
            nguon_view AS "Nguồn view"
        FROM v_dashboard_i15_moi_nhat
        ORDER BY "Kỳ", "Đơn vị", "NVKT"
        '''
    )


def load_i15_tong_hop_df():
    return read_sql_dataframe(
        '''
        SELECT
            ngay_du_lieu AS "Ngày dữ liệu",
            ttvt AS "TTVT",
            nhom_du_lieu AS "Nhóm dữ liệu",
            nhom_chi_tieu AS "Nhóm chỉ tiêu",
            don_vi AS "Kỳ",
            ten_chi_so AS "Tên chỉ số",
            gia_tri_so AS "Giá trị số",
            chi_tieu_bsc AS "Chỉ tiêu BSC",
            nguon_view AS "Nguồn view"
        FROM v_dashboard_i15_tong_hop_moi_nhat
        ORDER BY "Kỳ", "Tên chỉ số"
        '''
    )


def load_i15_tracking_df():
    return read_sql_dataframe(
        '''
        SELECT
            k_suffix AS "Kỳ",
            account_cts AS "Account CTS",
            ngay_xuat_hien_dau_tien AS "Ngày xuất hiện đầu tiên",
            ngay_thay_cuoi_cung AS "Ngày thay cuối cùng",
            so_ngay_lien_tuc AS "Số ngày liên tục",
            doi_one AS "Đơn vị",
            nvkt_db_normalized AS "NVKT",
            sa AS "SA",
            trang_thai AS "Trạng thái",
            nhom_du_lieu AS "Nhóm dữ liệu",
            nhom_chi_tieu AS "Nhóm chỉ tiêu",
            nguon_view AS "Nguồn view"
        FROM v_dashboard_i15_tracking_hien_tai
        ORDER BY "Kỳ", "Đơn vị", "Số ngày liên tục" DESC, "Account CTS"
        '''
    )


def load_dashboard_kpi_nvkt_df():
    return read_sql_dataframe(
        '''
        SELECT
            ngay_du_lieu AS "Ngày dữ liệu",
            nhom_chi_tieu AS "Nhóm chỉ tiêu",
            don_vi AS "Đơn vị",
            nvkt AS "NVKT",
            sm1 AS "SM1",
            sm2 AS "SM2",
            sm3 AS "SM3",
            sm4 AS "SM4",
            sm5 AS "SM5",
            sm6 AS "SM6",
            chi_so_1 AS "Chỉ số 1",
            ten_chi_so_1 AS "Tên chỉ số 1",
            chi_so_2 AS "Chỉ số 2",
            ten_chi_so_2 AS "Tên chỉ số 2",
            chi_so_3 AS "Chỉ số 3",
            ten_chi_so_3 AS "Tên chỉ số 3",
            chi_tieu_bsc AS "Điểm BSC"
        FROM v_dashboard_kpi_nvkt_moi_nhat
        ORDER BY "Đơn vị", "NVKT", "Nhóm chỉ tiêu"
        '''
    )


def load_kpi_nvkt_tong_hop_df():
    return read_sql_dataframe(
        '''
        SELECT
            nhom_chi_tieu,
            ngay_du_lieu,
            don_vi,
            nvkt,
            sm1,
            sm2,
            sm3,
            sm4,
            sm5,
            sm6,
            chi_so_1,
            ten_chi_so_1,
            chi_so_2,
            ten_chi_so_2,
            chi_so_3,
            ten_chi_so_3,
            chi_tieu_bsc
        FROM v_kpi_nvkt_tong_hop_moi_nhat
        ORDER BY don_vi, nvkt, nhom_chi_tieu
        '''
    )


def load_dashboard_dich_vu_theo_to_df(unit_name):
    return read_sql_dataframe(
        '''
        SELECT
            ngay_du_lieu AS "Ngày dữ liệu",
            loai_dich_vu AS "Loại dịch vụ",
            hanh_dong AS "Hành động",
            doi_vien_thong AS "Đội viễn thông",
            nvkt AS "NVKT",
            so_luong AS "Số lượng"
        FROM v_dashboard_dich_vu_theo_to_moi_nhat
        WHERE ttvt = ?
        ORDER BY "Loại dịch vụ", "Đội viễn thông", "NVKT", "Hành động"
        ''',
        (unit_name,),
    )


def load_dashboard_thuc_tang_df(unit_name):
    return read_sql_dataframe(
        '''
        SELECT
            ngay_du_lieu AS "Ngày dữ liệu",
            loai_dich_vu AS "Loại dịch vụ",
            cap_tong_hop AS "Cấp tổng hợp",
            doi_vien_thong AS "Đội viễn thông",
            nvkt AS "NVKT",
            hoan_cong AS "Hoàn công",
            ngung_phat_sinh_cuoc AS "Ngưng phát sinh cước",
            thuc_tang AS "Thực tăng",
            ty_le_ngung_psc AS "Tỷ lệ ngưng PSC"
        FROM v_dashboard_thuc_tang_moi_nhat
        WHERE ttvt = ?
        ORDER BY "Loại dịch vụ", "Cấp tổng hợp", "Đội viễn thông", "NVKT"
        ''',
        (unit_name,),
    )


def load_ngung_psc_fiber_df(unit_name):
    return read_sql_dataframe(
        '''
        SELECT *
        FROM v_ngung_psc_fiber_thang_t_1_cap_ttvt
        ORDER BY "Đơn vị/Nhân viên KT"
        ''',
    )


def load_ngung_psc_mytv_df(unit_name):
    return read_sql_dataframe(
        '''
        SELECT *
        FROM v_ngung_psc_mytv_thang_t_1_cap_ttvt
        ORDER BY "Đơn vị/Nhân viên KT"
        ''',
    )


def load_xac_minh_chi_tiet_df(unit_name):
    return read_sql_dataframe(
        '''
        SELECT
            ngay_du_lieu,
            "Dịch vụ",
            "Mã thuê bao",
            "Tên thuê bao",
            "Kiểu lệnh",
            "Ngày lập hợp đồng",
            "Ngày hoàn thành",
            "Loại phiếu",
            "Khu vực",
            "Đội VT",
            "TTVT",
            "NVKT"
        FROM v_xac_minh_ui_chi_tiet_moi_nhat
        WHERE "TTVT" = ?
        ORDER BY "Đội VT", "NVKT", "Ngày hoàn thành" DESC
        ''',
        (unit_name,),
    )


def load_tam_dung_khoi_phuc_tong_hop_theo_to_df():
    return read_sql_dataframe(
        '''
        SELECT *
        FROM v_tam_dung_khoi_phuc_dich_vu_chi_tiet_combined_tong_hop_theo_to
        ORDER BY CASE WHEN "TTVT" = 'TỔNG CỘNG' THEN 1 ELSE 0 END, "TTVT", "DOIVT"
        '''
    )


def load_tam_dung_khoi_phuc_tong_hop_theo_nvkt_df():
    return read_sql_dataframe(
        '''
        SELECT *
        FROM v_tam_dung_khoi_phuc_dich_vu_chi_tiet_combined_tong_hop_theo_nvkt
        ORDER BY "TTVT", "DOIVT", "NVKT"
        '''
    )


def load_khoi_phuc_fiber_df(unit_name):
    return read_sql_dataframe(
        '''
        SELECT
            ma_thue_bao AS "Mã thuê bao",
            ten_thue_bao AS "Tên thuê bao",
            ngay_lap_hop_dong AS "Ngày lập hợp đồng",
            ngay_thuc_hien AS "Ngày thực hiện",
            ten_kieu_lenh AS "Tên kiểu lệnh",
            ly_do_huy AS "Lý do hủy",
            trang_thai_thue_bao AS "Trạng thái thuê bao",
            ten_loai_hop_dong AS "Tên loại hợp đồng",
            trang_thai_hop_dong AS "Trạng thái hợp đồng",
            ma_giao_dich AS "Mã giao dịch",
            doi_vien_thong AS "Đội viễn thông",
            nvkt AS "NVKT"
        FROM khoi_phuc_fiber
        WHERE ttvt = ?
        ORDER BY "Đội viễn thông", "Ngày thực hiện" DESC
        ''',
        (unit_name,),
    )
