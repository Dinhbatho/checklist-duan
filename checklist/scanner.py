"""Engine quét cây thư mục — thay cho phần VBA LayTenFolder / DuyetThuMuc.

Khắc phục toàn bộ lỗi của bản VBA:
  * Lỗi 91/92 (Object variable not set / For loop not initialized)
      -> nguyên nhân cũ: `For Each ... In folder.SubFolders` của FSO ném lỗi khi
         gặp path > 259 ký tự hoặc folder không quyền đọc, rồi `Resume Next`
         nhảy lại vào chính vòng lặp đã hỏng.
      -> nay: mỗi folder được quét trong một khối try/except riêng, lỗi được
         ghi log và bỏ qua đúng một folder đó.
  * Lỗi 70 (Permission denied) -> bắt PermissionError, ghi log, đi tiếp.
  * Lỗi 28 (Out of stack space) -> duyệt bằng STACK tường minh, không đệ quy.
  * Path dài > 260 ký tự -> tiền tố \\\\?\\ (Windows) như API FindFirstFileW.
  * Vòng lặp junction/symlink vô tận -> bỏ qua reparse point + nhớ (dev, inode).
"""

from __future__ import annotations

import os
import sys
from dataclasses import dataclass, field
from datetime import datetime
from typing import Callable, Iterable

# Thuộc tính file của Windows
FILE_ATTRIBUTE_HIDDEN = 0x2
FILE_ATTRIBUTE_SYSTEM = 0x4
FILE_ATTRIBUTE_REPARSE_POINT = 0x400

# Giới hạn số dòng dữ liệu của Excel (1.048.576 trừ dòng tiêu đề)
MAX_DONG_EXCEL = 1_048_575

IS_WINDOWS = os.name == "nt"


# --------------------------------------------------------------------------- #
# Kiểu dữ liệu
# --------------------------------------------------------------------------- #
@dataclass
class FileRow:
    ten: str
    duong_dan: str
    ngay_tao: datetime | None
    ngay_sua: datetime | None


@dataclass
class FolderNode:
    duong_dan: str
    ma: str                      # mã phân cấp: I, I.1, I.1.3 ...
    files: list[FileRow] = field(default_factory=list)


@dataclass
class LoiRow:
    ma_loi: str
    mo_ta: str
    ham: str
    duong_dan: str
    thoi_diem: datetime


@dataclass
class KetQuaQuet:
    folders: list[FolderNode] = field(default_factory=list)
    loi: list[LoiRow] = field(default_factory=list)
    so_file: int = 0
    bi_huy: bool = False
    vuot_gioi_han: bool = False
    giay: float = 0.0


# --------------------------------------------------------------------------- #
# Tiện ích đường dẫn
# --------------------------------------------------------------------------- #
def them_tien_to_dai(p: str) -> str:
    """Thêm tiền tố \\\\?\\ để Windows chấp nhận đường dẫn > 260 ký tự."""
    if not IS_WINDOWS:
        return p
    if p.startswith("\\\\?\\"):
        return p
    if p.startswith("\\\\"):                       # UNC: \\server\share
        return "\\\\?\\UNC\\" + p[2:]
    if len(p) >= 2 and p[1] == ":":                # ổ đĩa cục bộ
        return "\\\\?\\" + p
    return p


def bo_tien_to_dai(p: str) -> str:
    if p.startswith("\\\\?\\UNC\\"):
        return "\\\\" + p[8:]
    if p.startswith("\\\\?\\"):
        return p[4:]
    return p


def ghep(a: str, b: str) -> str:
    sep = "\\" if IS_WINDOWS else "/"
    return a + b if a.endswith(sep) else a + sep + b


def _thoi_diem_tao(st) -> float | None:
    """Thời điểm TẠO file.

    Windows: st_ctime chính là thời điểm tạo. Linux/macOS: st_ctime là thời điểm
    đổi metadata (không phải ngày tạo) nên chỉ dùng st_birthtime nếu có, ngược
    lại trả None để không đánh dấu nhầm 'file mới'.
    """
    bt = getattr(st, "st_birthtime", None)
    if bt:
        return bt
    return st.st_ctime if IS_WINDOWS else None


def _ngay(ts: float | None) -> datetime | None:
    if not ts:
        return None
    try:
        return datetime.fromtimestamp(ts)
    except (OSError, OverflowError, ValueError):
        return None


# --------------------------------------------------------------------------- #
# Scanner
# --------------------------------------------------------------------------- #
class Scanner:
    def __init__(
        self,
        root: str,
        duoi_file: Iterable[str],
        loai_tru: Iterable[str] = (),
        do_sau_toi_da: int = 120,
        bo_qua_an: bool = True,
        theo_doi_tien_trinh: Callable[[int, int, str], None] | None = None,
    ) -> None:
        self.root = root.rstrip("\\/") if len(root.rstrip("\\/")) > 2 else root
        self.duoi_file = {d.strip().lstrip(".").lower() for d in duoi_file if d.strip()}
        self.loai_tru = {n.strip().lower() for n in loai_tru if n.strip()}
        self.do_sau_toi_da = do_sau_toi_da
        self.bo_qua_an = bo_qua_an
        self.theo_doi = theo_doi_tien_trinh
        self.kq = KetQuaQuet()
        self._da_tham: set[tuple[int, int]] = set()
        self._so_dong = 0            # số dòng sẽ ghi ra sheet ListFile

    # ---------------- log lỗi ---------------- #
    def _ghi_loi(self, ham: str, err: BaseException, duong_dan: str) -> None:
        ma = getattr(err, "winerror", None) or getattr(err, "errno", None) or ""
        self.kq.loi.append(
            LoiRow(str(ma), f"{type(err).__name__}: {err}".strip(), ham,
                   bo_tien_to_dai(duong_dan), datetime.now())
        )

    def _ghi_loi_text(self, ham: str, mo_ta: str, duong_dan: str) -> None:
        self.kq.loi.append(LoiRow("", mo_ta, ham, bo_tien_to_dai(duong_dan), datetime.now()))

    # ---------------- lọc ---------------- #
    def _duoi_hop_le(self, ten: str) -> bool:
        if not self.duoi_file:
            return True
        i = ten.rfind(".")
        return i > 0 and ten[i + 1:].lower() in self.duoi_file

    @staticmethod
    def _la_reparse(entry: os.DirEntry) -> bool:
        try:
            if entry.is_symlink():
                return True
            attrs = getattr(entry.stat(follow_symlinks=False), "st_file_attributes", 0)
            return bool(attrs & FILE_ATTRIBUTE_REPARSE_POINT)
        except OSError:
            return False

    def _bi_an(self, entry: os.DirEntry) -> bool:
        if not self.bo_qua_an:
            return False
        try:
            attrs = getattr(entry.stat(follow_symlinks=False), "st_file_attributes", 0)
            if attrs & (FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_SYSTEM):
                return True
        except OSError:
            pass
        return entry.name.startswith(".") and not IS_WINDOWS

    # ---------------- quét một folder ---------------- #
    def _quet_mot_lan(self, duong_dan: str) -> tuple[list[str], list[FileRow]]:
        """Quét đúng MỘT folder. Mọi lỗi được nuốt tại đây, không lan ra ngoài."""
        thu_muc_con: list[str] = []
        files: list[FileRow] = []
        try:
            with os.scandir(them_tien_to_dai(duong_dan)) as it:
                while True:
                    # next() được bọc riêng: một entry hỏng không làm chết cả folder
                    try:
                        entry = next(it)
                    except StopIteration:
                        break
                    except OSError as e:
                        self._ghi_loi("scandir.next", e, duong_dan)
                        break
                    try:
                        self._xu_ly_entry(entry, thu_muc_con, files)
                    except OSError as e:
                        self._ghi_loi("xu_ly_entry", e, ghep(duong_dan, entry.name))
        except (PermissionError, FileNotFoundError, NotADirectoryError, OSError) as e:
            self._ghi_loi("quet_folder", e, duong_dan)

        thu_muc_con.sort(key=str.lower)
        files.sort(key=lambda f: f.ten.lower())
        return thu_muc_con, files

    def _xu_ly_entry(self, entry: os.DirEntry, thu_muc_con: list[str], files: list[FileRow]) -> None:
        ten = entry.name
        if ten in (".", ".."):
            return
        try:
            la_thu_muc = entry.is_dir(follow_symlinks=False)
        except OSError as e:
            self._ghi_loi("is_dir", e, entry.path)
            return

        if la_thu_muc:
            if ten.lower() in self.loai_tru or self._bi_an(entry) or self._la_reparse(entry):
                return
            thu_muc_con.append(ten)
            return

        if ten.startswith("~$"):            # file khóa tạm của Office
            return
        if not self._duoi_hop_le(ten):
            return
        try:
            st = entry.stat(follow_symlinks=False)
            ngay_tao = _ngay(_thoi_diem_tao(st))
            ngay_sua = _ngay(st.st_mtime)
        except OSError as e:
            self._ghi_loi("stat", e, entry.path)
            ngay_tao = ngay_sua = None
        # duong_dan để trống, vòng ngoài sẽ ghép với path gốc (không kèm \\?\)
        files.append(FileRow(ten, "", ngay_tao, ngay_sua))

    # ---------------- duyệt toàn cây (KHÔNG đệ quy) ---------------- #
    def quet(self) -> KetQuaQuet:
        import time

        t0 = time.time()
        # stack: (đường dẫn, mã phân cấp, độ sâu)
        stack: list[tuple[str, str, int]] = [(self.root, "I", 0)]

        try:
            while stack:
                duong_dan, ma, do_sau = stack.pop()

                if self.kq.vuot_gioi_han:
                    break

                # chống vòng lặp: mỗi thư mục vật lý chỉ vào một lần
                try:
                    st = os.stat(them_tien_to_dai(duong_dan))
                    khoa = (st.st_dev, st.st_ino)
                    if khoa in self._da_tham and st.st_ino != 0:
                        self._ghi_loi_text("quet", "Bỏ qua vì đã duyệt (junction/lối tắt vòng)", duong_dan)
                        continue
                    if st.st_ino != 0:
                        self._da_tham.add(khoa)
                except OSError as e:
                    self._ghi_loi("stat_folder", e, duong_dan)
                    continue

                thu_muc_con, files = self._quet_mot_lan(duong_dan)
                for f in files:
                    f.duong_dan = ghep(duong_dan, f.ten)

                # giới hạn dòng của Excel
                if self._so_dong + 1 + len(files) > MAX_DONG_EXCEL:
                    self.kq.vuot_gioi_han = True
                    self._ghi_loi_text("quet", "Vượt giới hạn dòng Excel — dừng tại đây", duong_dan)
                    break

                self.kq.folders.append(FolderNode(duong_dan, ma, files))
                self._so_dong += 1 + len(files)
                self.kq.so_file += len(files)

                if self.theo_doi and len(self.kq.folders) % 50 == 0:
                    self.theo_doi(len(self.kq.folders), self.kq.so_file, duong_dan)

                if do_sau >= self.do_sau_toi_da:
                    if thu_muc_con:
                        self._ghi_loi_text(
                            "quet", f"Vượt độ sâu {self.do_sau_toi_da} cấp — không quét tiếp", duong_dan
                        )
                    continue

                # đẩy ngược để thứ tự duyệt là A→Z (pre-order như bản VBA)
                for k in range(len(thu_muc_con), 0, -1):
                    stack.append((ghep(duong_dan, thu_muc_con[k - 1]), f"{ma}.{k}", do_sau + 1))

        except KeyboardInterrupt:
            self.kq.bi_huy = True
            print("\n[!] Đã dừng giữa chừng (Ctrl+C). Dữ liệu quét được vẫn sẽ ghi ra file.",
                  file=sys.stderr)

        self.kq.giay = time.time() - t0
        return self.kq


def thu_muc_ton_tai(p: str) -> bool:
    try:
        return os.path.isdir(them_tien_to_dai(p))
    except OSError:
        return False
