# -*- coding: utf-8 -*-
"""
PAYMENT AGENT — Sub-Agent Thanh toán khối lượng hoàn thành (Mẫu 03a, NĐ 254/2025/NĐ-CP)

Input : bảng QS hợp đồng (--qs) + file khối lượng thực hiện (--progress)
        + cơ sở đơn giá (--price-basis) + tỷ lệ thu hồi tạm ứng / giữ lại (bắt buộc khai báo)
Output: qs_data.payment_summary, qs_data.payment_period_03a_vnd; file Mẫu 03a (--payment-out)
Thiếu dữ liệu → dừng; --demo dùng bảng QS mẫu và khối lượng giả định 50% (đánh dấu mẫu).
"""

from __future__ import annotations
import os
import sys
from typing import Dict, Optional

_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from core.supervisor.base_agent import BaseAgent, DataInputError, MissingDataError
from core.state.state_bus import StateBus

DEMO_HINT = "Muốn chạy thử với dữ liệu mẫu Cầu Km19+529.080 thì thêm cờ --demo."
MAX_LISTED = 10


class PaymentAgent(BaseAgent):

    def __init__(
        self,
        qs_path: Optional[str] = None,
        qs_sheet: Optional[str] = None,
        rate_overrides: Optional[Dict[str, Optional[float]]] = None,
        progress_path: Optional[str] = None,
        progress_sheet: Optional[str] = None,
        price_basis: Optional[str] = None,
        advance_recovery_pct: Optional[float] = None,
        retention_pct: Optional[float] = None,
        advance_outstanding: Optional[float] = None,
        period: str = "",
        payment_out: Optional[str] = None,
        sample_path: Optional[str] = None,
    ):
        super().__init__(agent_id="payment_agent",
                         description="Thanh toán khối lượng hoàn thành — Mẫu 03a NĐ 254/2025")
        self.qs_path = qs_path
        self.qs_sheet = qs_sheet
        self.rate_overrides = rate_overrides or {}
        self.progress_path = progress_path
        self.progress_sheet = progress_sheet
        self.price_basis = price_basis
        self.advance_recovery_pct = advance_recovery_pct
        self.retention_pct = retention_pct
        self.advance_outstanding = advance_outstanding
        self.period = period
        self.payment_out = payment_out
        self.sample_path = sample_path or os.path.join(
            _ROOT, "templates", "Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx")

    def run(self, bus: StateBus) -> bool:
        print("  [PaymentAgent] Lập bảng xác định giá trị khối lượng hoàn thành (Mẫu 03a)...")
        from tools.payment import PaymentError, ProgressRow, compute_payment, load_progress, write_payment_workbook
        from tools.qs_loader import QSLoadError, apply_rate_overrides, load_qs

        demo = bus.is_demo_mode()
        qs_path, progress_path = self.qs_path, self.progress_path
        basis, advance, retention = self.price_basis, self.advance_recovery_pct, self.retention_pct

        missing = []
        if not qs_path:
            missing.append("bảng QS hợp đồng (--qs)")
        if not progress_path:
            missing.append("khối lượng thực hiện kỳ này (--progress)")
        if not basis:
            missing.append("cơ sở đơn giá (--price-basis direct: đơn giá QS là chi phí trực tiếp; "
                           "contract: đã là đơn giá hợp đồng trước thuế)")
        if advance is None:
            missing.append("tỷ lệ thu hồi tạm ứng (--advance-recovery-pct, ghi 0 nếu không có)")
        if retention is None:
            missing.append("tỷ lệ giữ lại (--retention-pct, ghi 0 nếu không có)")
        if missing and not demo:
            raise MissingDataError("Thiếu dữ liệu lập Mẫu 03a: " + "; ".join(missing) + ". " + DEMO_HINT)
        if missing:
            bus.mark_sample_data(self.agent_id, "03a mẫu: " + "; ".join(missing)
                                 + " — dùng bảng QS mẫu, khối lượng kỳ này = 50% HĐ, tạm ứng 20%, giữ lại 5%")
            qs_path = qs_path or self.sample_path
            basis = basis or "direct"
            advance = 20.0 if advance is None else advance
            retention = 5.0 if retention is None else retention

        try:
            est = load_qs(qs_path, sheet=self.qs_sheet)
            if est.errors:
                raise DataInputError(f"Bảng QS {est.source} có {len(est.errors)} dòng sai dữ liệu: {est.errors[:3]}")
            apply_rate_overrides(est, self.rate_overrides)
            if basis == "direct" or self.rate_overrides.get("vat") is None:
                if est.missing_rates:
                    raise MissingDataError("Thiếu tỷ lệ GT / TL / VAT để quy đổi đơn giá và tính thuế: "
                                           f"{est.missing_rates} — khai báo --rate-* / --vat")
                est.compute()
            if progress_path:
                progress = load_progress(progress_path, sheet=self.progress_sheet)
            else:
                progress = [ProgressRow(i.row, i.code, i.stt, i.description, 0.0, round(i.quantity * 0.5, 3))
                            for i in est.items]
            result = compute_payment(
                est, progress, basis, advance, retention,
                vat_rate=None if self.rate_overrides.get("vat") is None else self.rate_overrides["vat"] / 100,
                advance_outstanding=self.advance_outstanding, period=self.period,
            )
        except (QSLoadError, PaymentError) as e:
            raise DataInputError(str(e)) from e
        if result.errors:
            raise DataInputError("Mẫu 03a có lỗi:\n    " + "\n    ".join(result.errors[:MAX_LISTED]))

        paid_items = sum(1 for l in result.lines if l.this_qty > 0)
        print(f"  [PaymentAgent] Hợp đồng: {len(result.lines)} công tác, giá trị trước thuế {result.contract_value:,}")
        if basis == "direct":
            print(f"  [PaymentAgent] Đơn giá HĐ = đơn giá QS × {result.price_factor:.6f} (hệ số G/T)")
        print(f"  [PaymentAgent] Kỳ {self.period or '...'}: {paid_items} công tác có khối lượng, "
              f"giá trị {result.this_value:,} + VAT {result.this_vat:,} = {result.this_total:,}")
        print(f"  [PaymentAgent] Thu hồi tạm ứng {result.advance_recovery:,} | giữ lại {result.retention:,} "
              f"| ĐỀ NGHỊ THANH TOÁN {result.payable:,} VNĐ")
        print(f"  [PaymentAgent] Lũy kế đến hết kỳ: {result.cumulative_value:,} "
              f"({result.cumulative_value / result.contract_value:.1%} giá trị HĐ)" if result.contract_value else "")
        for w in result.warnings[:MAX_LISTED]:
            print(f"    ⚠ {w}")

        bus.set_qs_data({
            "payment_period_03a_vnd": result.payable,
            "payment_summary": {
                "period": self.period, "price_basis": basis, "price_factor": result.price_factor,
                "contract_value": result.contract_value, "cumulative_value": result.cumulative_value,
                "this_value": result.this_value, "vat_rate": result.vat_rate, "this_vat": result.this_vat,
                "this_total": result.this_total, "advance_recovery": result.advance_recovery,
                "retention": result.retention, "payable": result.payable,
                "overrun_value": result.overrun_value, "unmatched": len(result.unmatched),
                "line_values": [l.this_value for l in result.lines],
                "warnings": list(result.warnings),
            },
        })
        if self.payment_out:
            write_payment_workbook(self.payment_out, result, bus._state.project_name)
            print(f"  [PaymentAgent] Đã xuất Mẫu 03a: {self.payment_out}")
        return True
