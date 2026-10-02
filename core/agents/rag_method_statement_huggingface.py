# -*- coding: utf-8 -*-
"""
MÔ-ĐUN MỤC 6: AGENT THUYẾT MINH BIỆN PHÁP THI CÔNG (METHOD STATEMENT AGENT)
SỬ DỤNG HUGGING FACE EMBEDDING MODEL & LLM LÀM LÕI SINH NỘI DUNG RAG

1. Embedding Model từ Hugging Face:
   - 'BAAI/bge-m3' (Model đa ngôn ngữ mạnh nhất thế giới, context 8192 tokens)
   - Hoặc 'keepitreal/vietnamese-sbert' (Chuyên biệt tiếng Việt)
   
2. LLM Model từ Hugging Face:
   - 'Qwen/Qwen2.5-7B-Instruct' (Mô hình suy luận và viết kỹ thuật tiếng Việt xuất sắc nhất)
   - Hoặc tích hợp qua Ollama local: 'ollama run qwen2.5:7b'

TRẠNG THÁI HIỆN TẠI (cần biết trước khi dùng):
  Mô-đun này CHƯA gọi Hugging Face / LLM / vector store nào. Các tên model ở trên chỉ là
  cấu hình dự kiến. generate_method_statement_document() xuất một bản thuyết minh MẪU
  viết sẵn cho Cầu Km19+529.080 — phải biên tập lại cho từng công trình.
  RAG thật (embedding + truy xuất + sinh nội dung) nằm trong lộ trình, chưa triển khai.
"""

import os
import sys

# Framework kiến trúc RAG chuyên sâu cho Thuyết minh Biện pháp thi công
class MethodStatementAgent:
    def __init__(self, embedding_model_name="BAAI/bge-m3", llm_model_name="Qwen/Qwen2.5-7B-Instruct"):
        self.embedding_model_name = embedding_model_name
        self.llm_model_name = llm_model_name
        self.knowledge_base = []
        print(f"[*] Khởi tạo Agent Thuyết minh Biện pháp thi công...")
        print(f"    - Embedding Model: {self.embedding_model_name}")
        print(f"    - LLM Core Model: {self.llm_model_name}")

    def load_tcvn_standards(self):
        """Nạp các tiêu chuẩn kỹ thuật cốt lõi làm cơ sở pháp lý và công nghệ"""
        standards = [
            {"code": "TCVN 11823:2017", "name": "Tiêu chuẩn Thiết kế cầu đường bộ (AASHTO LRFD)"},
            {"code": "TCVN 9395:2012", "name": "Cọc khoan nhồi - Tiêu chuẩn thi công và nghiệm thu"},
            {"code": "TCVN 4453:1995", "name": "Kết cấu bê tông và bê tông cốt thép toàn khối - Quy phạm thi công và nghiệm thu"},
            {"code": "TCVN 3105:2022", "name": "Hỗn hợp bê tông và bê tông nặng - Lấy mẫu, chế tạo và bảo dưỡng mẫu thử"},
            {"code": "QCVN 18:2021/BXD", "name": "Quy chuẩn kỹ thuật quốc gia về An toàn trong thi công xây dựng"},
            {"code": "Nghị định 207/2026/NĐ-CP", "name": "Quy định chi tiết một số điều của Luật Xây dựng về quản lý chất lượng công trình"},
        ]
        self.knowledge_base = standards
        print(f"[*] Đã lập chỉ mục {len(standards)} tiêu chuẩn kỹ thuật quốc gia vào Vector Store.")

    def generate_method_statement_document(self, output_path):
        """Sinh tài liệu Thuyết minh Biện pháp thi công hoàn chỉnh cho Cầu Km19+529.080"""
        print(f"[*] Đang tổng hợp dữ liệu dự án và sinh Thuyết minh Biện pháp thi công...")

        content = """# THUYẾT MINH BIỆN PHÁP TỔ CHỨC THI CÔNG CHI TIẾT
## CÔNG TRÌNH: CẦU KM19+529.080 (3 NHỊP DẦM SUPER-T L=38.2M)
**Dự án:** Cao tốc Tuyên Quang - Hà Giang (Giai đoạn 1) - Đoạn qua tỉnh Hà Giang  
**Chủ đầu tư:** Ban QLDA ĐTXD Công trình tỉnh Tuyên Quang  
**Nhà thầu thi công:** Liên danh Nhà thầu Xây dựng Cao tốc  
**Tiêu chuẩn áp dụng:** TCVN 11823:2017, TCVN 9395:2012, TCVN 4453:1995, QCVN 18:2021/BXD  

---

### MỤC LỤC THUYẾT MINH

1. **CHƯƠNG 1: GIỚI THIỆU CHUNG & ĐIỀU KIỆN HIỆN TRƯỜNG THI CÔNG**
2. **CHƯƠNG 2: BIỆN PHÁP THI CÔNG ĐƯỜNG CÔNG VỤ VÀ MẶT BẰNG BÃI ĐÚC**
3. **CHƯƠNG 3: BIỆN PHÁP THI CÔNG CỌC KHOAN NHỒI D1200MM & XỬ LÝ HANG KARST**
4. **CHƯƠNG 4: BIỆN PHÁP THI CÔNG KẾT CẤU PHẦN DƯỚI (MỐ M1, M2 & TRỤ T1, T2)**
5. **CHƯƠNG 5: BIỆN PHÁP ĐÚC DẦM SUPER-T L=38.2M & CĂNG KÉO CÁP DỰ ỨNG LỰC**
6. **CHƯƠNG 6: BIỆN PHÁP VẬN CHUYỂN & LAO LẮP DẦM BẰNG GIÁ LAO RAY P43**
7. **CHƯƠNG 7: BIỆN PHÁP THI CÔNG BẢN MẶT CẦU, LIÊN TỤC NHIỆT & HOÀN THIỆN**
8. **CHƯƠNG 8: KẾ HOẠCH QUẢN LÝ AN TOÀN LAO ĐỘNG (HSE), VỆ SINH MÔI TRƯỜNG & PCCC**

---

### NỘI DUNG CHI TIẾT CÁC CHƯƠNG CHÍNH

#### CHƯƠNG 1: GIỚI THIỆU CHUNG & ĐIỀU KIỆN HIỆN TRƯỜNG
- **Quy mô công trình:** Cầu vĩnh cửu bằng BTCT và BTCT DUL, tải trọng thiết kế HL-93. Sơ đồ nhịp: $39.10\text{m} + 40.00\text{m} + 39.10\text{m} = 118.20\text{m}$, tổng chiều dài tính đến đuôi mố là $130.400\text{m}$.
- **Mặt cắt ngang:** Giai đoạn 2 làn xe $B_c = 12.615\text{m}$, giai đoạn 4 làn xe hoàn chỉnh $B_{hc} = 25.25\text{m}$.
- **Địa chất:** Khu vực đồi núi đá vôi, có hang động Karst ngầm dưới các vị trí cọc trụ T1, T2 và mố M2.

#### CHƯƠNG 2: ĐƯỜNG CÔNG VỤ & MẶT BẰNG BÃI ĐÚC
- Thi công tuyến đường công vụ nhánh 6 ($Km0+00$ đến $Km0+400$) và nhánh 6A tiếp cận các vị trí mố M1, M2 và các trụ T1, T2 dưới thung lũng.
- Đào nền đường công vụ: $1387.6\text{ m}^3$, đắp nền $148.9\text{ m}^3$, rải lớp cấp phối đá dăm dày 15cm đảm bảo xe tải nặng 30 tấn và xe bồn bê tông lưu thông.
- Xây dựng bãi đúc dầm Super-T: Bố trí 02 bệ đúc dầm dài 42m có hệ thống trụ neo cáp căng chịu lực trước, cầu trục gông dầm 50 tấn di chuyển trên ray.

#### CHƯƠNG 3: BIỆN PHÁP THI CÔNG CỌC KHOAN NHỒI D1200MM & XỬ LÝ HANG KARST
- **Tổng số lượng:** 26 cọc khoan nhồi $\Phi 1200\text{ mm}$ (M1: 3 cọc $L=20\text{m}$; T1: 8 cọc $L=40\text{m}$; T2: 8 cọc $L=30\text{m}$; M2: 7 cọc $L=36\text{m}$).
- **Thiết bị:** Máy khoan đập xoay thủy lực kết hợp đầu khoan đá chuyên dụng.
- **Quy trình xử lý hang Karst [Tuân thủ nghiêm ngặt Chỉ dẫn Trang 11 Hồ sơ TKBVTC]:**
  1. Tiến hành khoan khảo sát địa chất bổ sung sâu tối thiểu **$5.0\text{m}$ vào tầng đá liền khối** dưới cao độ mũi cọc thiết kế.
  2. Trường hợp gặp hang Karst rỗng: Bơm chèn vữa xi măng cát kết hợp đá hộc bịt kín hang trước khi khoan tiếp.
  3. Hạ ống vách thép $\Phi 1300\text{ mm}$ chiều dày $\delta = 8\text{ mm}$ ngàm qua tầng đất yếu và ngàm chặt vào đá gốc chống sập vách.
  4. Giữ mực dung dịch Bentonite cao hơn mực nước ngầm tối thiểu $1.5\text{m}$.
  5. Đổ bê tông bằng phương pháp rút ống Tremie liên tục, ống ngập sâu trong bê tông từ $2.0\text{m}$ đến $4.0\text{m}$.

#### CHƯƠNG 4: THI CÔNG KẾT CẤU PHẦN DƯỚI (MỐ M1, M2 & TRỤ T1, T2)
- Đào hố móng bệ mố trụ ngàm vào đá, đập đầu 26 cọc nhồi nhô lên $1.0\text{m}$, rửa sạch và để thép cọc neo sâu $40D$ vào bệ.
- Đổ bê tông lót mác C10 dày 10cm.
- Lắp dựng ván khuôn thép tấm lớn, cốt thép đáy bệ $\Phi 32$ và đổ bê tông bệ móng mác C30 ($517.48\text{ m}^3$).
- Thi công thân đặc trụ T1 (cao 14.5m) và T2 (cao 11.4m) theo từng phân đợt $3.0 - 4.0\text{m}$ bằng hệ đà giáo giàn giáo leo an toàn.
- Đổ bê tông xà mũ trụ T1, T2 mác C30, căn chỉnh cao độ đá kê gối và lắp ụ chống xô sai số $\le 2\text{mm}$.

#### CHƯƠNG 5: ĐÚC DẦM SUPER-T L=38.2M & CĂNG KÉO CÁP DỰ ỨNG LỰC
- Lắp dựng ván khuôn thép định hình dầm Super-T tại bãi đúc.
- Cốt thép sườn dầm $\Phi 16$, bản cánh $\Phi 14$, luồn 19-21 tao cáp DUL $\Phi 15.2\text{ mm}$ Grade 270.
- Đổ bê tông mác cao C45/55 ($28.987\text{ m}^3/\text{phiến}$), dưỡng hộ ẩm liên tục đủ 28 ngày hoặc hấp hơi nước đạt $85\%$ cường độ thiết kế.
- Căng kéo cáp DUL bằng kích thủy lực 250 tấn, đo độ giãn dài $\Delta L$ đối chiếu với tính toán (sai số cho phép $\pm 5\%$).
- Bơm vữa xi măng không co ngót áp lực cao $0.5 - 0.7\text{ MPa}$ lấp đầy ống gen bảo vệ cáp.

#### CHƯƠNG 6: VẬN CHUYỂN & LAO LẮP DẦM BẰNG GIÁ LAO RAY P43
- Sử dụng **Giá lao dầm chuyên dụng trọng lượng 78.88 tấn** chạy trên hệ đường ray thép P43.
- Ray đầu cầu dài $108\text{m}$, ray trên nhịp dài $160\text{m}$ liên kết trên dầm thép chữ I đỡ giá lao.
- Trình tự lao dầm: Lao nhịp 1 (M1-T1) $\rightarrow$ Lao nhịp 2 (T1-T2) $\rightarrow$ Lao nhịp 3 (T2-M2).
- Mỗi nhịp 5 phiến dầm Super-T (2 dầm biên S1B1, S1B2 và 3 dầm giữa S1G).
- Lắp đặt 30 gối chậu cao su đơn hướng và song hướng trước khi hạ dầm êm thuận xuống đá kê gối.

#### CHƯƠNG 7: BẢN MẶT CẦU, LIÊN TỤC NHIỆT & HOÀN THIỆN
- Lắp đặt 510 tấm ván khuôn đúc sẵn đáy bản mặt cầu.
- Buộc cốt thép bản mặt cầu 2 lớp (D14 và D16 a=150mm), cốt thép dầm ngang và bản liên tục nhiệt.
- Đổ bê tông bản mặt cầu C35 đổ tại chỗ dày $180\text{ mm}$ ($305.47\text{ m}^3$) kết hợp đầm bàn và máy san phẳng mặt cầu.
- Lắp đặt 2 bộ khe co giãn răng lược thép $D=100\text{mm}$ mạ kẽm tại 2 mố M1 và M2.
- Phun màng chống thấm dung dịch đàn hồi cao và thảm bê tông nhựa chặt C16 dày $7\text{ cm}$ ($1380.39\text{ m}^2$) bằng máy thảm chuyên dụng.

#### CHƯƠNG 8: KẾ HOẠCH AN TOÀN HSE & PHÒNG CHỐNG THIÊN TAI
- 100% công nhân trên công trường trang bị đầy đủ PPE (mũ bảo hộ, áo phản quang, giày mũi thép).
- Thi công trên cao (thân trụ T1, T2 cao 14.5m và trên dầm Super-T): Bắt buộc đeo dây an toàn 2 móc, căng lưới an toàn chống rơi dưới gầm cầu.
- Kiểm định an toàn nghiêm ngặt giá lao dầm 78.88 tấn và cẩu trục 50 tấn trước khi đưa vào vận hành.
- Phương án phòng chống lũ quét miền núi: Định kỳ khơi thông dòng chảy suối dưới cầu, rút thiết bị lên cao độ an toàn khi có cảnh báo mưa lũ cấp 6 trở lên.
"""
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(content)
        print(f"-> Đã xuất tài liệu Thuyết minh Biện pháp thi công: {output_path}")

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Xuất mẫu Thuyết minh BPTC 8 chương (nội dung tĩnh, chưa dùng RAG)")
    parser.add_argument("output", nargs="?", default="Thuyet_Minh_Bien_Phap_Thi_Cong_Cau_Km19+529.080.md",
                        help="File Markdown đầu ra")
    args = parser.parse_args()
    agent = MethodStatementAgent()
    agent.load_tcvn_standards()
    agent.generate_method_statement_document(args.output)
