# Hướng dẫn đóng góp — 23HG-AEC-MultiAgent-System

Cảm ơn bạn muốn đóng góp. Dự án phục vụ hồ sơ kỹ thuật, dự toán và thanh toán thật, nên
nguyên tắc số một là: **con số sai nguy hiểm hơn con số thiếu**. Thiếu dữ liệu thì dừng và
báo rõ; không đoán, không lấp bằng số mẫu.

## 1. Chuẩn bị môi trường

```bash
git clone https://github.com/baotuhg/23HG-multiagent-system.git
cd 23HG-multiagent-system
python -m venv venv
source venv/bin/activate          # Windows: .\venv\Scripts\activate
pip install -r requirements.txt
```

Python 3.10 trở lên. Chức năng AutoCAD COM chỉ chạy trên Windows có AutoCAD; mọi phần còn lại
(kể cả đọc DXF) chạy được trên Linux / macOS.

## 2. Chạy kiểm thử

```bash
python -m unittest discover -s tests -t . -v     # toàn bộ test
python run_state_graph.py --solver-test          # tự kiểm tra OR-Tools + CPM
python run_state_graph.py --demo                 # chạy thử 8 pha bằng dữ liệu mẫu
```

CI (GitHub Actions) chạy hai lệnh đầu trên Ubuntu và Windows cho mọi Pull Request. PR chỉ được
gộp khi CI xanh.

## 3. Quy tắc bắt buộc khi viết code

1. **Không để LLM làm toán.** Mọi phép tính khối lượng, đơn giá, cắt thép, tiến độ nằm trong
   Python thuần (`tools/`), có test. LLM chỉ gọi công cụ và trình bày kết quả.
2. **Dữ liệu thật trước, dữ liệu mẫu chỉ khi `--demo`.** Agent thiếu dữ liệu thật phải
   `raise MissingDataError(...)` và nói rõ cần cờ / file nào. Khi dùng số mẫu ở chế độ demo,
   gọi `bus.mark_sample_data(agent_id, ghi_chú)` để log, Quality Gate và Human Gate đều cảnh báo.
3. **Lỗi dữ liệu đầu vào không retry.** Dùng `DataInputError` (sửa file mới hết lỗi); chỉ trả
   `False` từ `run()` khi chạy lại có thể cho kết quả khác.
4. **Không trả kết quả "khớp" / "đạt" khi chưa kiểm tra.** Chưa có dữ liệu thì ghi
   `NOT_CHECKED` / `NOT_RUN`, không ghi `MATCHED` / `PASS`.
5. **Không ghi đường dẫn máy cá nhân** (`C:\Users\...`, `D:\Code\...`) vào code trong `core/`,
   `tools/`, `agents/`. Nhận đường dẫn qua tham số dòng lệnh.
6. **Không commit file sinh ra khi chạy**: `.aec_state/`, file Excel / CSV đầu ra, file khóa
   Office `~$*.xlsx`. Xem `.gitignore`.
7. Căn cứ pháp lý / tiêu chuẩn (TCVN, TT, NĐ) ghi rõ số hiệu trong docstring của hàm áp dụng.

## 4. Thêm một Sub-Agent mới

1. Viết công cụ tính toán trong `tools/<ten>.py` + test trong `tests/test_<ten>.py`.
2. Viết agent kế thừa `core.supervisor.base_agent.BaseAgent`, chỉ đọc / ghi qua `StateBus`.
3. Nếu cần trường state mới, thêm vào dataclass tương ứng trong `core/state/shared_state.py`
   (có giá trị mặc định để state cũ vẫn khôi phục được).
4. Đăng ký agent và thêm cờ dòng lệnh trong `run_state_graph.py`; nếu agent đọc file, bổ sung
   vào `check_inputs()` để `--check-inputs` kiểm tra được file đó.
5. Cập nhật README (mục lệnh thực thi) và quy trình trong `workflows/` nếu có.

## 5. Quy trình Pull Request

1. Tạo nhánh từ `main`: `feat/...`, `fix/...`, `docs/...`.
2. Commit nhỏ, thông điệp nói rõ **vì sao** (tiếng Việt hoặc tiếng Anh đều được).
3. Mở PR, mô tả: thay đổi gì, căn cứ tiêu chuẩn nào, đã kiểm thử bằng dữ liệu nào.
   Nếu thay đổi làm đổi con số đầu ra (dự toán, số cây thép, ngày hoàn thành...), ghi rõ số cũ → số mới.
4. Không đưa dữ liệu dự án chưa được phép công bố (giá thầu, hồ sơ chủ đầu tư) vào repo.

## 6. Báo lỗi

Mở Issue kèm: lệnh đã chạy, log lỗi, và nếu được thì file đầu vào đã lược bỏ thông tin nhạy cảm.
