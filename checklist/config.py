"""Đọc cấu hình từ config.ini và lưu trạng thái lần quét trước."""

from __future__ import annotations

import configparser
import json
import os
from dataclasses import dataclass, field
from datetime import datetime

DUOI_MAC_DINH = ["pdf", "xls", "xlsx", "xlsm", "doc", "docx", "jpg", "png", "dwg", "ppt", "pptx"]
LOAI_TRU_MAC_DINH = ["$RECYCLE.BIN", "System Volume Information", "$SysReset", "Recovery"]


@dataclass
class CauHinh:
    root: str = ""
    thu_muc_ra: str = "ketqua"
    duoi_file: list[str] = field(default_factory=lambda: list(DUOI_MAC_DINH))
    loai_tru: list[str] = field(default_factory=lambda: list(LOAI_TRU_MAC_DINH))
    do_sau_toi_da: int = 120
    bo_qua_an: bool = True
    to_mau_duong_dan: bool = True

    @classmethod
    def doc(cls, duong_dan: str) -> "CauHinh":
        c = cls()
        if not os.path.isfile(duong_dan):
            return c
        p = configparser.ConfigParser(interpolation=None)
        p.read(duong_dan, encoding="utf-8")
        g = p["quet"] if p.has_section("quet") else {}

        def lay(khoa, mac_dinh):
            v = g.get(khoa, "") if hasattr(g, "get") else ""
            return v.strip() if v and v.strip() else mac_dinh

        c.root = lay("folder_goc", "")
        c.thu_muc_ra = lay("thu_muc_ra", c.thu_muc_ra)
        duoi = lay("duoi_file", "")
        if duoi:
            c.duoi_file = [d.strip().lstrip(".").lower() for d in duoi.split(",") if d.strip()]
        loai_tru = lay("loai_tru", "")
        if loai_tru:
            c.loai_tru = [x.strip() for x in loai_tru.split(",") if x.strip()]
        c.do_sau_toi_da = int(lay("do_sau_toi_da", str(c.do_sau_toi_da)))
        c.bo_qua_an = lay("bo_qua_an", "1") not in ("0", "false", "False")
        c.to_mau_duong_dan = lay("to_mau_duong_dan", "1") not in ("0", "false", "False")
        return c


# --------------------------------------------------------------------------- #
class TrangThai:
    """Nhớ thời điểm quét trước để đánh dấu file mới tạo / đã sửa."""

    def __init__(self, duong_dan: str) -> None:
        self.duong_dan = duong_dan
        self.du_lieu: dict = {}
        if os.path.isfile(duong_dan):
            try:
                with open(duong_dan, encoding="utf-8") as f:
                    self.du_lieu = json.load(f)
            except (OSError, ValueError):
                self.du_lieu = {}

    def lan_truoc(self, root: str) -> datetime | None:
        muc = self.du_lieu.get(root.rstrip("\\/").lower())
        if not muc:
            return None
        try:
            return datetime.fromisoformat(muc["thoi_diem"])
        except (KeyError, ValueError):
            return None

    def ghi(self, root: str, thoi_diem: datetime) -> None:
        self.du_lieu[root.rstrip("\\/").lower()] = {"thoi_diem": thoi_diem.isoformat()}
        try:
            with open(self.duong_dan, "w", encoding="utf-8") as f:
                json.dump(self.du_lieu, f, ensure_ascii=False, indent=2)
        except OSError:
            pass
