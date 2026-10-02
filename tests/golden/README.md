# Bảng dự toán / thanh toán chuẩn (golden files)

Thư mục này chứa các bảng QS **thật, đã được duyệt**, cùng con số kỳ vọng để chứng minh chương trình tính
đúng so với thực tế (không chỉ nhất quán với chính nó). Test `GoldenFilesTest` trong
`tests/test_money_and_golden.py` tự nhận mọi file `*.json` ở đây; chưa có file thì test bị bỏ qua.

Định dạng một file `ten_du_an.json`:

```json
{
  "qs_file": "tests/golden/ten_du_an_qs.xlsx",
  "sheet": null,
  "rates_pct": {"chung": 6.5, "nha_tam": 1, "kxd": 2, "tl": 5.5, "vat": 10},
  "line_rounding": false,
  "expected": {"T": 0, "GT": 0, "TL": 0, "G": 0, "VAT": 0, "G_XD": 0}
}
```

- `expected` lấy từ bảng đã duyệt (số do kỹ sư QS/Chủ đầu tư xác nhận), sai số cho phép **0 đồng**.
- `line_rounding`: `true` nếu bảng duyệt làm tròn từng dòng thành tiền trước khi cộng, `false` nếu cộng rồi làm tròn một lần.
- Không đưa dữ liệu nhạy cảm (đơn giá thầu bí mật) vào repo công khai; có thể ẩn danh hóa hạng mục nhưng giữ nguyên số.
