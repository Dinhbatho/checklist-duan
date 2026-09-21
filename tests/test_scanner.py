"""Kiểm thử engine quét — chạy: python -m pytest -q  (hoặc python tests/test_scanner.py)"""

import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from checklist.scanner import Scanner, ghep  # noqa: E402

DUOI = ["pdf", "docx", "xlsx"]


def _tao_cay(goc: str) -> None:
    os.makedirs(ghep(goc, "01-Phap ly"), exist_ok=True)
    os.makedirs(ghep(goc, "02-Rong"), exist_ok=True)
    for ten in ("hop dong.pdf", "bao gia.xlsx", "ghi chu.txt", "~$tam.docx"):
        open(os.path.join(goc, "01-Phap ly", ten), "w").write("x")


def test_loc_duoi_file_va_bo_file_tam():
    with tempfile.TemporaryDirectory() as d:
        _tao_cay(d)
        kq = Scanner(d, DUOI).quet()
        ten = [f.ten for node in kq.folders for f in node.files]
        assert sorted(ten) == ["bao gia.xlsx", "hop dong.pdf"], ten
        assert kq.so_file == 2


def test_folder_rong_van_duoc_liet_ke():
    with tempfile.TemporaryDirectory() as d:
        _tao_cay(d)
        kq = Scanner(d, DUOI).quet()
        assert len(kq.folders) == 3          # gốc + 01-Phap ly + 02-Rong


def test_ma_phan_cap():
    with tempfile.TemporaryDirectory() as d:
        _tao_cay(d)
        kq = Scanner(d, DUOI).quet()
        ma = [n.ma for n in kq.folders]
        assert ma[0] == "I"
        assert set(ma[1:]) == {"I.1", "I.2"}


def test_duong_dan_dai_va_do_sau():
    """Cây sâu 40 cấp, đường dẫn > 400 ký tự — không được tràn stack, không lỗi."""
    with tempfile.TemporaryDirectory() as d:
        p = d
        for i in range(40):
            p = os.path.join(p, f"thu_muc_ten_kha_dai_de_test_duong_dan_{i:02d}")
        os.makedirs(p, exist_ok=True)
        open(os.path.join(p, "cuoi cung.pdf"), "w").write("x")
        kq = Scanner(d, DUOI).quet()
        assert kq.so_file == 1
        assert len(kq.folders) == 41
        assert kq.loi == []


def test_gioi_han_do_sau_ghi_log_chu_khong_vo():
    with tempfile.TemporaryDirectory() as d:
        p = d
        for i in range(6):
            p = os.path.join(p, f"c{i}")
        os.makedirs(p, exist_ok=True)
        kq = Scanner(d, DUOI, do_sau_toi_da=3).quet()
        assert len(kq.folders) == 4          # gốc + 3 cấp
        assert any("độ sâu" in l.mo_ta for l in kq.loi)


def test_symlink_vong_khong_lap_vo_tan():
    if os.name == "nt":
        return
    with tempfile.TemporaryDirectory() as d:
        con = os.path.join(d, "con")
        os.makedirs(con, exist_ok=True)
        os.symlink(d, os.path.join(con, "quay_lai"))
        kq = Scanner(d, DUOI).quet()
        assert len(kq.folders) == 2          # gốc + con, lối tắt bị bỏ qua


def test_folder_khong_quyen_chi_ghi_1_dong_log():
    if os.name == "nt" or os.geteuid() == 0:
        return                                # root đọc được mọi thứ
    with tempfile.TemporaryDirectory() as d:
        cam = os.path.join(d, "cam")
        os.makedirs(os.path.join(cam, "sub"), exist_ok=True)
        os.chmod(cam, 0o000)
        try:
            kq = Scanner(d, DUOI).quet()
            assert len(kq.loi) == 1
            assert "Permission" in kq.loi[0].mo_ta
        finally:
            os.chmod(cam, 0o755)


if __name__ == "__main__":
    loi = 0
    for ten, ham in sorted(globals().items()):
        if ten.startswith("test_") and callable(ham):
            try:
                ham()
                print(f"  OK   {ten}")
            except AssertionError as e:
                loi += 1
                print(f"  FAIL {ten}: {e}")
    print("\nTất cả đạt." if not loi else f"\n{loi} test không đạt.")
    sys.exit(1 if loi else 0)
