# -*- coding: utf-8 -*-
"""
Kiểm thử số học tiền & "đáp án vàng" tính tay.

Mọi con số kỳ vọng dưới đây được tính tay (phân số, độc lập với code cần kiểm), KHÔNG lấy từ chính
kết quả của chương trình. Các phép tính đều có trường hợp "đúng một nửa" (x.5) để phân biệt
ROUND của Excel (ra xa số 0) với round() của Python (về số chẵn).
"""

import os
import sys
import tempfile
import unittest
from decimal import Decimal

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from tools.money import mul, round_half_up, round_vnd, total
from tools.payment import compute_payment, load_progress
from tools.qs_loader import QSEstimate, QSItem, apply_rate_overrides, load_qs


def make_estimate(items, rates_pct):
    est = QSEstimate(source="golden")
    for k, (q, p) in enumerate(items, 1):
        est.items.append(QSItem(row=k, stt=str(k), code=f"X{k}", description=f"Công tác {k}", unit="m3",
                                quantity=q, unit_price=p, amount=q * p))
    apply_rate_overrides(est, rates_pct)
    return est


class MoneyRoundingTest(unittest.TestCase):

    def test_matches_excel_round_half_away_from_zero(self):
        # Excel: ROUND(0.5,0)=1  ROUND(1.5,0)=2  ROUND(2.5,0)=3  ROUND(-2.5,0)=-3  ROUND(2.4999,0)=2
        self.assertEqual([round_vnd(x) for x in (0.5, 1.5, 2.5, -2.5, -0.5, 2.4999)], [1, 2, 3, -3, -1, 2])
        # Python round() về số chẵn → khác Excel: đây là lý do không dùng nó cho tiền
        self.assertEqual([round(0.5), round(2.5)], [0, 2])

    def test_decimal_places_like_excel(self):
        self.assertEqual(round_half_up(2.675, 2), Decimal("2.68"))     # float nhị phân sẽ ra 2.67
        self.assertEqual(round_half_up(1.005, 2), Decimal("1.01"))

    def test_exact_products_no_binary_error(self):
        self.assertEqual(mul(592.4, 980000), Decimal("580552000.0"))
        self.assertEqual(total([0.1, 0.2]), Decimal("0.3"))            # float: 0.30000000000000004


class QSGoldenTest(unittest.TestCase):
    RATES = {"chung": 5, "nha_tam": 1, "kxd": 2, "tl": 5.5, "vat": 10}

    def test_ties_default_rounds_sum_once(self):
        # 0.5×1001 = 500.5 ; 2.5×1001 = 2502.5 ; tổng 3003.0
        # GT = 150 + 30 + 60 = 240 ; TL = round(3243×5.5%) = round(178.365) = 178
        # G = 3421 ; VAT = round(342.1) = 342 ; G_XD = 3763
        est = make_estimate([(0.5, 1001), (2.5, 1001)], self.RATES)
        est.compute()
        self.assertEqual((est.T, est.GT, est.TL, est.G, est.VAT, est.G_XD), (3003, 240, 178, 3421, 342, 3763))

    def test_ties_line_rounding(self):
        # từng dòng: round(500.5)=501 ; round(2502.5)=2503 ; T = 3004 (khác 3003 khi làm tròn một lần)
        # GT = 150 + 30 + 60 = 240 ; TL = round(3244×5.5%) = round(178.42) = 178 ; G = 3422 ; VAT = 342
        est = make_estimate([(0.5, 1001), (2.5, 1001)], self.RATES)
        est.compute(line_rounding=True)
        self.assertEqual((est.T, est.GT, est.TL, est.G, est.VAT, est.G_XD), (3004, 240, 178, 3422, 342, 3764))

    def test_realistic_estimate(self):
        # 100.5×1,234,567 = 124,073,983.5 ; 2000.25×98,765 = 197,554,691.25 ; 10.125×21,500,000 = 217,687,500
        # T = round(539,316,174.75) = 539,316,175
        # chung 6.5% = 35,055,551 ; nhà tạm 1% = 5,393,162 ; KXĐ 2% = 10,786,324 → GT = 51,235,037
        # TL = round(590,551,212 × 5.5%) = 32,480,317 ; G = 623,031,529 ; VAT = 62,303,153
        est = make_estimate([(100.5, 1234567), (2000.25, 98765), (10.125, 21500000)],
                            {"chung": 6.5, "nha_tam": 1, "kxd": 2, "tl": 5.5, "vat": 10})
        est.compute()
        self.assertEqual(est.T, 539_316_175)
        self.assertEqual(est.GT_components, {"chung": 35_055_551, "nha_tam": 5_393_162, "kxd": 10_786_324})
        self.assertEqual((est.GT, est.TL, est.G, est.VAT, est.G_XD),
                         (51_235_037, 32_480_317, 623_031_529, 62_303_153, 685_334_682))


class PaymentGoldenTest(unittest.TestCase):

    def test_contract_basis_with_roundings(self):
        # Hợp đồng: A 100×1,000,000 ; B 3×333,333. Kỳ này: A 40, B 1.
        # Giá trị = 40,000,000 + 333,333 = 40,333,333 ; VAT = round(4,033,333.3) = 4,033,333
        # Cộng = 44,366,666 ; thu hồi tạm ứng 20% = round(8,873,333.2) = 8,873,333
        # Giữ lại 5% = round(2,218,333.3) = 2,218,333 ; đề nghị thanh toán = 33,275,000
        with tempfile.TemporaryDirectory() as tmp:
            qs = os.path.join(tmp, "qs.csv")
            with open(qs, "w", encoding="utf-8") as f:
                f.write("STT;Mã hiệu;Nội dung công tác;ĐVT;Khối lượng;Đơn giá\n"
                        "1;A;Công tác A;m3;100;1000000\n2;B;Công tác B;m3;3;333333\n")
            pg = os.path.join(tmp, "kl.csv")
            with open(pg, "w", encoding="utf-8") as f:
                f.write("Mã hiệu;Nội dung;KL lũy kế kỳ trước;KL thực hiện kỳ này\nA;;0;40\nB;;0;1\n")
            est = load_qs(qs)
            apply_rate_overrides(est, {"chung": 5, "nha_tam": 1, "kxd": 1, "tl": 5, "vat": 10})
            est.compute()
            res = compute_payment(est, load_progress(pg), "contract", 20, 5)
        self.assertEqual(res.this_value, 40_333_333)
        self.assertEqual(res.this_vat, 4_033_333)
        self.assertEqual(res.this_total, 44_366_666)
        self.assertEqual((res.advance_recovery, res.retention), (8_873_333, 2_218_333))
        self.assertEqual(res.payable, 33_275_000)


class GoldenFilesTest(unittest.TestCase):
    """Đối chiếu với bảng dự toán THẬT đã được duyệt (xem tests/golden/README.md). Không có file → bỏ qua."""

    def test_approved_estimates(self):
        import glob
        import json
        specs = sorted(glob.glob(os.path.join(ROOT, "tests", "golden", "*.json")))
        if not specs:
            self.skipTest("Chưa có bảng dự toán duyệt trong tests/golden/ — xem tests/golden/README.md")
        for spec_path in specs:
            with open(spec_path, encoding="utf-8") as f:
                spec = json.load(f)
            with self.subTest(spec=os.path.basename(spec_path)):
                est = load_qs(os.path.join(ROOT, spec["qs_file"]), spec.get("sheet"))
                if spec.get("rates_pct"):
                    apply_rate_overrides(est, spec["rates_pct"])
                est.compute(line_rounding=bool(spec.get("line_rounding", False)))
                for key, expected in spec["expected"].items():   # T, GT, TL, G, VAT, G_XD — sai số cho phép 0 đồng
                    self.assertEqual(getattr(est, key), expected, f"{key} khác bảng duyệt")


if __name__ == "__main__":
    unittest.main()
