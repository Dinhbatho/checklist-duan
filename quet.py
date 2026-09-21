#!/usr/bin/env python3
"""CheckList Dự án đầu tư — quét thư mục dự án và xuất báo cáo Excel.

Bản Python thay cho phần VBA của CheckList_Dự_án_đầu_tư_v4.xlsm.
Kết quả vẫn là file Excel với đúng bố cục Info / ListFolder / ListFile.

Cách dùng:
    python quet.py                          # đọc config.ini
    python quet.py --root "L:\\05-Du an dau tu\\215DT-Yen Binh"
    python quet.py --root "..." --out "bao_cao.xlsx"
"""

from __future__ import annotations

import argparse
import os
import sys
from datetime import datetime

from checklist.config import CauHinh, TrangThai
from checklist.excel import BaoCaoExcel
from checklist.scanner import Scanner, thu_muc_ton_tai

THU_MUC_GOC = os.path.dirname(os.path.abspath(__file__))


def ten_file_ra(root: str, thu_muc_ra: str) -> str:
    ten_da = os.path.basename(root.rstrip("\\/")) or "ketqua"
    an_toan = "".join(c if c.isalnum() or c in " -_." else "_" for c in ten_da).strip()
    dau_tg = datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(thu_muc_ra, f"CheckList_{an_toan}_{dau_tg}.xlsx")


CO_TTY = sys.stdout.isatty()


def tien_trinh(so_folder: int, so_file: int, duong_dan: str) -> None:
    if not CO_TTY:                      # tránh rác khi ghi log ra file
        return
    dong = f"  Đang quét: {so_folder:,} folder | {so_file:,} file | {duong_dan[-70:]}"
    sys.stdout.write("\r" + dong.ljust(120)[:120])
    sys.stdout.flush()


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Quét thư mục dự án, xuất báo cáo Excel.")
    ap.add_argument("--root", help="Thư mục gốc cần quét (ghi đè config.ini)")
    ap.add_argument("--out", help="Đường dẫn file .xlsx kết quả")
    ap.add_argument("--config", default=os.path.join(THU_MUC_GOC, "config.ini"))
    ap.add_argument("--duoi", help="Danh sách đuôi file, ngăn bởi dấu phẩy. VD: pdf,docx,xlsx")
    ap.add_argument("--do-sau", type=int, help="Độ sâu tối đa (mặc định 120)")
    ap.add_argument("--khong-to-mau", action="store_true", help="Không tô màu từng cấp đường dẫn (nhanh hơn)")
    a = ap.parse_args(argv)

    cfg = CauHinh.doc(a.config)
    root = a.root or cfg.root
    if a.duoi:
        cfg.duoi_file = [d.strip().lstrip(".").lower() for d in a.duoi.split(",") if d.strip()]
    if a.do_sau:
        cfg.do_sau_toi_da = a.do_sau
    if a.khong_to_mau:
        cfg.to_mau_duong_dan = False

    if not root:
        print("LỖI: chưa khai báo thư mục gốc. Sửa 'folder_goc' trong config.ini "
              "hoặc chạy với --root \"đường dẫn\".", file=sys.stderr)
        return 2
    if not thu_muc_ton_tai(root):
        print(f"LỖI: không tìm thấy hoặc không truy cập được thư mục gốc:\n  {root}", file=sys.stderr)
        return 2

    thu_muc_ra = cfg.thu_muc_ra if os.path.isabs(cfg.thu_muc_ra) else os.path.join(THU_MUC_GOC, cfg.thu_muc_ra)
    os.makedirs(thu_muc_ra, exist_ok=True)
    duong_dan_ra = a.out or ten_file_ra(root, thu_muc_ra)

    tt = TrangThai(os.path.join(THU_MUC_GOC, ".trang_thai_quet.json"))
    lan_truoc = tt.lan_truoc(root)
    bat_dau = datetime.now()

    print(f"Thư mục gốc : {root}")
    print(f"Đuôi file   : {', '.join(cfg.duoi_file) or '(tất cả)'}")
    print(f"Lần quét trước: {lan_truoc.strftime('%d/%m/%Y %H:%M:%S') if lan_truoc else '(chưa có)'}")
    print("Nhấn Ctrl+C để dừng giữa chừng — dữ liệu đã quét vẫn được ghi ra file.\n")

    sc = Scanner(
        root=root,
        duoi_file=cfg.duoi_file,
        loai_tru=cfg.loai_tru,
        do_sau_toi_da=cfg.do_sau_toi_da,
        bo_qua_an=cfg.bo_qua_an,
        theo_doi_tien_trinh=tien_trinh,
    )
    kq = sc.quet()
    if CO_TTY:
        sys.stdout.write("\r" + " " * 120 + "\r")

    bc = BaoCaoExcel(
        kq=kq, root=root, duoi_file=cfg.duoi_file,
        lan_quet_truoc=lan_truoc, chay_lan_dau=lan_truoc is None,
        to_mau_duong_dan=cfg.to_mau_duong_dan,
    )
    bc.xuat(duong_dan_ra)
    if not kq.bi_huy:
        tt.ghi(root, bat_dau)

    print(f"Folder : {len(kq.folders):,}")
    print(f"File   : {kq.so_file:,}")
    print(f"Lỗi    : {len(kq.loi):,}" + ("  (xem sheet Info, cột N:R)" if kq.loi else ""))
    print(f"Thời gian: {kq.giay:.1f} giây")
    if kq.bi_huy:
        print("Trạng thái: ĐÃ DỪNG GIỮA CHỪNG — kết quả chưa đầy đủ, không cập nhật mốc thời gian.")
    if kq.vuot_gioi_han:
        print("Trạng thái: VƯỢT GIỚI HẠN DÒNG EXCEL — đã cắt bớt.")
    print(f"\nĐã ghi: {duong_dan_ra}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\nĐã hủy.", file=sys.stderr)
        raise SystemExit(130)
