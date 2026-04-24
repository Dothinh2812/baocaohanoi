# Consumer Guide Cho App Doc SQLite

Tai lieu nay danh cho ung dung khac muon doc du lieu tu SQLite history.

Path can uu tien:

- che do `standalone`: [report_history.db](/home/vtst/baocaohanoi/api_transition/report_history.db)
- che do `multi-instance`: `api_transition/runtime/<unit>/sqlite_history/report_history.db`

Neu app hien thi du lieu cho tung don vi rieng, moi don vi phai doc DB trong workspace runtime cua chinh don vi do. Khong dung chung 1 DB cho nhieu instance.

Muc tieu:

- Giup app khac doc dung va on dinh.
- Giam phu thuoc truc tiep vao bang noi bo.
- Chot data contract o muc `view`, khong o muc `table`.

## 1. Nguyen tac bat buoc

App khac nen tuan thu 4 nguyen tac sau:

1. Chi doc `view`, khong doc truc tiep cac `table` nghiep vu hay `dong_bao_cao_goc`.
2. Uu tien cac `view` co hau to `moi_nhat` neu dashboard hien thi snapshot gan nhat.
3. Xem `ngay_du_lieu` la ngay nghiep vu cua snapshot, khong phai thoi diem app query.
4. Chap nhan rang trong cung 1 ngay, du lieu co the bi ghi de khi batch import chay lai.
5. Trong mo hinh multi-instance, chon dung DB theo `unit` truoc khi query `view`.

## 2. App duoc phep doc gi

Danh sach `view` uu tien cho app:

- `v_tien_do_nap_bao_cao`
- `v_dashboard_chat_luong_don_vi_moi_nhat`
- `v_dashboard_kpi_nvkt_moi_nhat`
- `v_dashboard_dich_vu_theo_to_moi_nhat`
- `v_dashboard_thuc_tang_moi_nhat`
- `v_dashboard_xac_minh_moi_nhat`
- `v_dashboard_xac_minh_tam_dung_moi_nhat`
- `v_dashboard_cau_hinh_tu_dong_moi_nhat`
- `v_dashboard_vat_tu_thu_hoi_moi_nhat`
- `v_dashboard_chi_so_nvkt_moi_nhat`
- `v_dashboard_ttvt_son_tay_chi_so_don_vi_moi_nhat`
- `v_dashboard_ttvt_son_tay_tong_hop_moi_nhat`

Neu can chi tiet hon nua:

- `v_c11_tong_hop_moi_nhat`
- `v_c15_tong_hop_moi_nhat`
- `v_c15_chi_tiet_nvkt_moi_nhat`
- `v_c15_tong_hop_doi_moi_nhat`
- `v_dashboard_i15_moi_nhat`
- `v_dashboard_i15_tong_hop_moi_nhat`
- `v_dashboard_i15_tracking_hien_tai`
- `v_i15_daily_summary_moi_nhat`
- `v_i15_snapshots_moi_nhat`
- `v_i15_tracking_hien_tai`
- `v_ghtt_don_vi_moi_nhat`
- `v_hoan_cong_fiber_moi_nhat`
- `v_ngung_psc_fiber_moi_nhat`
- `v_khoi_phuc_fiber_moi_nhat`
- `v_cau_hinh_tu_dong_tong_hop_moi_nhat`

App khac khong nen doc:

- `dong_bao_cao_goc`
- `sheet_bao_cao`
- `bao_cao_ngay`
- `c11_tong_hop`, `ghtt_don_vi`, `hoan_cong_fiber`, ... truc tiep

Ly do:

- schema bang noi bo co the doi khi importer duoc nang cap
- `view` la lop contract on dinh hon

## 3. Quy tac refresh du lieu

He thong hien tai dung quy tac:

- cung `ma_bao_cao + ngay_du_lieu` => ghi de snapshot cu
- khac `ngay_du_lieu` => tao snapshot moi

Tac dong toi app:

- neu app dang doc `view ... moi_nhat`, du lieu co the thay doi trong ngay khi batch import chay lai
- app khong nen cache vo thoi han
- neu can doi soat, app nen luu kem `ngay_du_lieu`

## 4. Tu dien cot dung chung

Nhieu `view` dung chung cac cot sau:

- `ngay_du_lieu`
  - ngay cua snapshot du lieu, dang `YYYY-MM-DD`
- `nhom_du_lieu`
  - nhom lon de app chia module, vi du `chat_luong`, `dich_vu`, `thuc_tang`
- `nhom_chi_tieu`
  - nhom nghiep vu nho hon, vi du `c11`, `c12`, `Fiber`, `ptm`, `thay_the`
- `don_vi`
  - ten don vi hien thi, co the la `To Ky thuat...`, `Tong`, hoac `TTVT Son Tay`
- `ttvt`
  - ten trung tam vien thong
- `nvkt`
  - ten nhan vien ky thuat
- `ten_chi_so`
  - ten metric da chuan hoa
- `gia_tri_so`
  - gia tri metric dang so
- `chi_tieu_bsc`
  - diem/chi tieu BSC di kem, neu view co
- `nguon_view`
  - view goc tao ra dong du lieu nay, huu ich khi debug

## 5. Quy uoc hien thi

App dashboard nen hien thi theo quy uoc sau:

- cac cot bat dau bang `ty_le_`: hien thi dang phan tram, nhung gia tri hien tai dang o thang `0-100`, khong phai `0-1`
- cac cot nhu `hoan_cong`, `ngung_psc`, `so_phieu_xac_minh`, `tong_hop_dong`: hien thi dang so nguyen neu co the
- `chi_tieu_bsc`: hien thi dang diem, giu 2 chu so thap phan khi can
- neu `gia_tri_so` la `NULL`: app hien thi `-` thay vi `0`
- neu `ten_chi_so_3` hoac metric thu 3 khong ton tai: khong tao card rong

## 6. Data contract theo tung view

### 6.1. `v_tien_do_nap_bao_cao`

Muc dich:

- man hinh quan tri import
- theo doi report nao da nap, report nao loi

Do hat:

- 1 dong = 1 `ma_bao_cao` moi nhat

Cot quan trong:

- `ma_bao_cao`
- `ten_bao_cao`
- `nhom_bao_cao`
- `ngay_du_lieu`
- `trang_thai_nap`
- `so_dong_goc`
- `so_dong_tong_hop`
- `so_dong_chi_tiet`
- `trang_thai_nhat_ky`
- `thong_diep`

Goi y UI:

- bang admin
- badge `thanh_cong/that_bai`
- nut retry/import lai o app van hanh neu can

### 6.2. `v_dashboard_chat_luong_don_vi_moi_nhat`

Muc dich:

- hien thi cac chi so `c11/c12/c13/c14` theo don vi

Do hat:

- 1 dong = 1 `nhom_chi_tieu + don_vi`

Cot quan trong:

- `nhom_chi_tieu`
- `ngay_du_lieu`
- `don_vi`
- `chi_tieu_bsc`
- `ten_chi_so_1`, `chi_so_1`
- `ten_chi_so_2`, `chi_so_2`
- `ten_chi_so_3`, `chi_so_3`

Goi y UI:

- bang tong hop chat luong theo to
- card 3 metric + BSC
- radar/bar chart theo `don_vi`

### 6.3. `v_dashboard_kpi_nvkt_moi_nhat`

Muc dich:

- hien thi KPI NVKT cho `c11/c12/c13`

Do hat:

- 1 dong = 1 `nhom_chi_tieu + don_vi + nvkt`

Cot quan trong:

- `nhom_chi_tieu`
- `don_vi`
- `nvkt`
- `sm1..sm6`
- `chi_so_1/2/3`
- `chi_tieu_bsc`

Goi y UI:

- ranking NVKT
- top/bottom theo `chi_tieu_bsc`
- filter theo `nhom_chi_tieu` va `don_vi`

### 6.4. `v_dashboard_dich_vu_theo_to_moi_nhat`

Muc dich:

- tong hop so luong dich vu theo to/NVKT

Do hat:

- 1 dong = 1 `loai_dich_vu + hanh_dong + ttvt + doi_vien_thong + nvkt`

Cot quan trong:

- `loai_dich_vu`
- `hanh_dong`
- `ttvt`
- `doi_vien_thong`
- `nvkt`
- `so_luong`

Goi y UI:

- bar chart theo to
- filter `Fiber/MyTV`
- card tong `hoan_cong / ngung_psc / khoi_phuc`

### 6.5. `v_dashboard_thuc_tang_moi_nhat`

Muc dich:

- hien thi thuc tang theo `Fiber/MyTV`

Do hat:

- 1 dong = 1 `loai_dich_vu + cap_tong_hop + ttvt + doi_vien_thong + nvkt`

Cot quan trong:

- `loai_dich_vu`
- `cap_tong_hop`
- `doi_vien_thong`
- `nvkt`
- `hoan_cong`
- `ngung_phat_sinh_cuoc`
- `thuc_tang`
- `ty_le_ngung_psc`

Goi y UI:

- bang thuc tang
- chart so sanh `hoan_cong` va `ngung_phat_sinh_cuoc`

### 6.6. `v_dashboard_xac_minh_moi_nhat`

Muc dich:

- tong hop xac minh theo NVKT/to

Do hat:

- 1 dong = 1 `ttvt + doi_vien_thong + nvkt`

Cot quan trong:

- `ttvt`
- `doi_vien_thong`
- `nvkt`
- `so_phieu_xac_minh`

Goi y UI:

- bang theo doi xac minh
- bar chart theo to hoac NVKT

### 6.7. `v_dashboard_xac_minh_tam_dung_moi_nhat`

Muc dich:

- tong hop xac minh tam dung theo NVKT/to

Do hat:

- 1 dong = 1 `ttvt + doi_vien_thong + nvkt`

Cot quan trong:

- `ttvt`
- `doi_vien_thong`
- `nvkt`
- `so_phieu_xac_minh`

Goi y UI:

- bang tong hop xac minh tam dung
- bar chart theo to hoac NVKT

### 6.8. `v_dashboard_cau_hinh_tu_dong_moi_nhat`

Muc dich:

- tong hop cau hinh tu dong

Do hat:

- 1 dong = 1 `ma_bao_cao + ttvt + don_vi + loai_dong`

Cot quan trong:

- `ma_bao_cao`
- `ttvt`
- `don_vi`
- `loai_dong`
- `tong_hop_dong`
- `da_day_cau_hinh_tu_dong`
- `cau_hinh_thanh_cong`
- `ty_le_day_tu_dong`
- `ty_le_cau_hinh_thanh_cong`

Goi y UI:

- card tong `PTM` va `Thay the`
- bang theo TTVT / To

### 6.8. `v_dashboard_vat_tu_thu_hoi_moi_nhat`

Muc dich:

- tong hop vat tu thu hoi theo NVKT dia ban

Do hat:

- 1 dong = 1 `nvkt_dia_ban_giao + loai_vat_tu + trang_thai_thu_hoi`

Cot quan trong:

- `nvkt_dia_ban_giao`
- `loai_vat_tu`
- `trang_thai_thu_hoi`
- `so_luong`

### 6.10. `v_dashboard_ttvt_son_tay_chi_so_don_vi_moi_nhat`

Muc dich:

- data feed tong hop rieng cho Son Tay theo dang long-metrics
- co ca cap `ttvt` va cap `don_vi`

Do hat:

- 1 dong = 1 chi so cua 1 don vi/ttvt

Cot quan trong:

- `ttvt`
- `nhom_du_lieu`
- `nhom_chi_tieu`
- `cap_du_lieu`
- `don_vi`
- `ten_chi_so`
- `gia_tri_so`
- `chi_tieu_bsc`

Goi y UI:

- dashboard linh hoat
- chart dong, card, bang drill-down

### 6.11. `v_dashboard_chi_so_nvkt_moi_nhat`

Muc dich:

- hien thi tat ca chi so co the quy ve 1 NVKT cu the
- dung cho man hinh profile NVKT, drill-down va doi soat theo nguoi

Do hat:

- 1 dong = 1 chi so cua 1 NVKT

Cot quan trong:

- `ngay_du_lieu`
- `ttvt`
- `don_vi`
- `nvkt`
- `nhom_du_lieu`
- `nhom_chi_tieu`
- `ten_chi_so`
- `gia_tri_so`
- `chi_tieu_bsc`
- `loai_dich_vu`
- `hanh_dong`
- `nguon_view`

Nguon du lieu dang gom:

- `kpi_nvkt`
- `hai_long_nvkt`
- `ghtt`
- `ket_qua_tiep_thi`
- `dich_vu`
- `thuc_tang`
- `xac_minh`
- `cau_hinh_tu_dong`
- `vat_tu_thu_hoi`

Goi y UI:

- header profile NVKT
- card tong hop theo `nhom_du_lieu`
- bang long-metrics co bo loc `nhom_chi_tieu`

### 6.12. `v_dashboard_ttvt_son_tay_tong_hop_moi_nhat`

Muc dich:

- chi giu chi so cap tong hop `TTVT Son Tay`
- khong co dong theo tung to

Do hat:

- 1 dong = 1 chi so tong hop cua `TTVT Son Tay`

Cot quan trong:

- `ttvt`
- `nhom_du_lieu`
- `nhom_chi_tieu`
- `don_vi`
- `ten_chi_so`
- `gia_tri_so`
- `chi_tieu_bsc`

Goi y UI:

- man hinh executive summary
- card tong hop theo nhom `chat_luong`, `dich_vu`, `thuc_tang`, `cau_hinh_tu_dong`

## 7. Cac query mau

Lay toan bo chi so tong hop Son Tay:

```sql
SELECT *
FROM v_dashboard_ttvt_son_tay_tong_hop_moi_nhat
ORDER BY nhom_du_lieu, nhom_chi_tieu, ten_chi_so;
```

Lay toan bo chi so cua 1 NVKT:

```sql
SELECT *
FROM v_dashboard_chi_so_nvkt_moi_nhat
WHERE nvkt = 'Bùi Văn Cường'
ORDER BY nhom_du_lieu, nhom_chi_tieu, ten_chi_so;
```

Lay chi so chat luong theo to:

```sql
SELECT *
FROM v_dashboard_chat_luong_don_vi_moi_nhat
ORDER BY nhom_chi_tieu, don_vi;
```

Top 10 NVKT theo BSC:

```sql
SELECT *
FROM v_dashboard_kpi_nvkt_moi_nhat
WHERE chi_tieu_bsc IS NOT NULL
ORDER BY chi_tieu_bsc DESC
LIMIT 10;
```

Tong hoan cong/ngung PSC/khoi phuc theo dich vu:

```sql
SELECT loai_dich_vu, hanh_dong, SUM(so_luong) AS tong_so_luong
FROM v_dashboard_dich_vu_theo_to_moi_nhat
GROUP BY loai_dich_vu, hanh_dong
ORDER BY loai_dich_vu, hanh_dong;
```

Tong hop cau hinh tu dong Son Tay:

```sql
SELECT nhom_chi_tieu, ten_chi_so, gia_tri_so
FROM v_dashboard_ttvt_son_tay_tong_hop_moi_nhat
WHERE nhom_du_lieu = 'cau_hinh_tu_dong'
ORDER BY nhom_chi_tieu, ten_chi_so;
```

## 8. Quy uoc xu ly NULL va du lieu bat thuong

App can xu ly:

- `NULL` => hien thi `-`
- gia tri am => khong tu dong coi la loi, vi co the do report goc tinh ra
- `chi_tieu_bsc` co the duoc dung lam diem tong hop, khong phai luc nao cung trung voi `gia_tri_so`
- mot so view long-metrics co nhieu dong cho cung `ten_chi_so` nhung khac `nhom_chi_tieu`

## 9. Nen cache the nao

Khuyen nghi:

- dashboard thong thuong: cache 1-5 phut
- man hinh quan tri import: khong cache hoac cache rat ngan
- neu can doi soat lich su: luon gui kem `ngay_du_lieu`

## 10. Kiem tra nhanh truoc khi app doc

Kiem tra DB co du lieu moi nhat:

```bash
sqlite3 api_transition/report_history.db \
  "SELECT COUNT(*) FROM v_tien_do_nap_bao_cao WHERE trang_thai_nap = 'thanh_cong';"
```

Kiem tra view tong hop Son Tay:

```bash
sqlite3 api_transition/report_history.db \
  "SELECT COUNT(*) FROM v_dashboard_ttvt_son_tay_tong_hop_moi_nhat;"
```

## 11. Scope on dinh cua data contract

Tai lieu nay xem nhu contract cho app doc du lieu.

Cam ket on dinh tuong doi:

- ten `view` trong muc 2
- y nghia logic cua tung `view`
- y nghia cot dung chung trong muc 4

Co the thay doi trong tuong lai:

- them `view` moi
- them cot moi
- map them nhieu report hon vao cung `view`

Neu can thay doi vo contract, nen cap nhat tai lieu nay truoc.
