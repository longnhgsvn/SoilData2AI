# 20 câu hỏi phản biện và trả lời gợi ý

**Dùng cho phiên thảo luận tại Hội nghị ERSD 2026.**

Bài báo: *Sử dụng Python xử lý dữ liệu chỉ tiêu cơ lý đất để chuẩn bị dữ liệu huấn luyện AI*  
Tác giả: Ngô Thị Phương Thảo, Ngô Hùng Long, Nguyễn Thị Hải Yến.

---

## Cách dùng tài liệu này

- Mỗi câu trả lời có 2 phần: **trả lời ngắn 30–45 giây** và **nếu bị hỏi sâu**.
- Nguyên tắc 3 nhịp: **ghi nhận câu hỏi → trả lời trực tiếp → nêu giới hạn/hướng tiếp**.
- Không nên cố trả lời mọi thứ. Với nội dung ngoài phạm vi bài, nói rõ: *“Đây là hướng chúng tôi chưa triển khai trong bài; nếu được tiếp tục, chúng tôi sẽ…”*.
- Luôn phân biệt: **đã làm và đã kiểm chứng** / **mới là thiết kế** / **hướng phát triển**.

### Bảng số liệu tra nhanh

| Nội dung | Số liệu |
|---|---:|
| Tệp đầu vào | 166 (158 `.xlsx`, 8 `.xls`) |
| Tệp xử lý thành công (lần thử nghiệm) | 166/166 |
| Sheet có dữ liệu → CSV (bài báo) | 963 |
| Mẫu trong AllData (bài báo) | 20.532 |
| Chạy lại trên snapshot hiện hành | 962 CSV / 20.352 dòng |
| Số cột chỉ tiêu trong AllData | 400 |
| Thời gian chạy toàn bộ | khoảng 1,5–2,5 phút |
| Test tự động | 12/12 pass |

---

# Nhóm A — Mục tiêu, đóng góp, lựa chọn công nghệ

## Q1. Đóng góp mới của bài là gì, trong khi đã có nhiều công cụ chuyển Excel sang CSV?

**Trả lời ngắn:**

Đóng góp của bài không phải là “phát minh” định dạng CSV, mà là một **quy trình tích hợp và tái lập** cho một loại dữ liệu rất cụ thể: hồ sơ thí nghiệm đất lịch sử, nhiều sheet, nhiều định dạng, văn bản TCVN3. Quy trình gồm: quét batch, tách sheet, nhận diện TCVN3 bằng font + mẫu ký tự, chuyển Unicode bằng thuật toán hai pha, xuất CSV UTF-8 có truy vết, và kiểm chứng trên bộ dữ liệu thực.

**Nếu bị hỏi sâu:**

- Các script Excel → CSV thông thường giả định dữ liệu đã là Unicode và cấu trúc đồng nhất. Dữ liệu của chúng tôi thì không.
- Điểm khó nhất không phải đọc Excel, mà là **nhận diện đúng khi nào cần chuyển mã** và **chuyển đúng thứ tự** để không phá các cụm ký tự.
- Bài cung cấp thêm CLI, GUI, test và log để người khác chạy lại, không chỉ một lần chạy minh họa.
- Giá trị khoa học nằm ở **quy trình xử lý dữ liệu di sản tiếng Việt** và **bằng chứng thực nghiệm**, không phải ở việc viết một hàm đọc Excel.

**Câu chốt:** *“Chúng tôi không nói CSV là mới; chúng tôi nói quy trình tạo ra CSV đúng từ dữ liệu TCVN3 là phần đóng góp.”*

---

## Q2. Vì sao chọn Python? Có nhất thiết phải dùng Python không?

**Trả lời ngắn:**

Python phù hợp vì hệ sinh thái xử lý dữ liệu và học máy rất mạnh (`pandas`, `scikit-learn`, `PyTorch`), đồng thời có thư viện đọc cả `.xlsx` (`openpyxl`) và `.xls` cũ (`xlrd 1.2.0`, kể cả font metadata). Python cũng cho phép đóng gói CLI, GUI Qt6 và test tự động trong cùng một mã nguồn.

**Nếu bị hỏi sâu:**

- Không có nghĩa Python là lựa chọn duy nhất. Power Query, R, KNIME hay C# đều có thể làm được.
- Điểm quyết định là **cần đọc font của từng ô trong `.xls`** để nhận diện TCVN3; `openpyxl`/`xlrd` cho phép điều này.
- Python giúp bước tiếp theo — đưa CSV vào `pandas`/`scikit-learn`/`PyTorch` — liền mạch, không phải chuyển đổi công cụ.
- Với nhóm đã dùng Python cho AI, việc giữ một ngôn ngữ giảm chi phí bảo trì.

**Câu chốt:** *“Python không phải lựa chọn duy nhất, nhưng là lựa chọn ít ma sát nhất cho cả khâu đọc Excel lẫn khâu huấn luyện AI.”*

---

## Q3. Vì sao tự xây bảng ánh xạ TCVN3 thay vì dùng thư viện `vietnamese-conversion`?

**Trả lời ngắn:**

Dự án dùng bảng ánh xạ **tương thích VietKey** và nhúng trực tiếp vào mã nguồn để bảo đảm tính tái lập, không phụ thuộc phiên bản thư viện bên ngoài. Thư viện `vietnamese-conversion` là nguồn tham chiếu, nhưng với dữ liệu có cả cụm hai ký tự, chúng tôi cần kiểm soát thứ tự thay thế và chạy test hồi quy trên dữ liệu thực.

**Nếu bị hỏi sâu:**

- Nếu dùng thư viện ngoài, hành vi có thể thay đổi khi nâng phiên bản, ảnh hưởng tính tái lập của kết quả công bố.
- Bảng nhúng giúp chạy offline, không cần mạng, phù hợp môi trường hội nghị/đơn vị khảo sát.
- Chúng tôi không nói bảng tự xây là hoàn hảo; nó được kiểm chứng trên bộ dữ liệu và có test cho các cụm điển hình.
- Hướng tốt hơn là **đối chiếu chéo** với `vietnamese-conversion` trên một tập test chuẩn và công bố sai lệch.

**Câu chốt:** *“Chúng tôi không phát minh bảng mã; chúng tôi đóng băng một bảng đã kiểm chứng để kết quả có thể tái lập.”*

---

## Q4. Làm sao đảm bảo chuyển mã TCVN3 → Unicode là đúng, nhất là khi ánh xạ một chiều?

**Trả lời ngắn:**

Chúng tôi dùng ba lớp bảo vệ: (1) bảng ánh xạ VietKey có cả cụm hai ký tự; (2) sắp xếp giảm dần theo độ dài và thay thế hai pha; (3) kiểm chứng định tính bằng đối chiếu ngẫu nhiên với bản gốc, cùng test hồi quy cho các cụm tiếng Việt điển hình.

**Nếu bị hỏi sâu:**

- Chuyển mã một chiều nghĩa là không thể “undo” hoàn toàn; vì vậy phải kiểm tra **trước khi ghi**, không sửa CSV sau đó.
- Các cụm như `¢` + `Ê` → `Ấ` là điểm dễ sai nếu thay ký tự đơn trước.
- Chúng tôi thừa nhận chưa có ground truth do chuyên gia gán cho toàn bộ 20 nghìn dòng; đây là hạn chế của nghiên cứu.
- Đề xuất: xây bộ test chuẩn có chuyên gia xác nhận, đo precision/recall của nhận diện TCVN3, và lưu log các ô nghi ngờ.

**Câu chốt:** *“Chúng tôi không tuyên bố chuyển mã đúng tuyệt đối; chúng tôi tuyên bố quy trình có thể kiểm tra, lặp lại và chỉ ra chỗ cần review.”*

---

## Q5. Thuật toán “hai pha” là gì? Có thực sự cần thiết hay chỉ làm phức tạp?

**Trả lời ngắn:**

Cần, vì bảng ánh xạ có cả ký tự đơn và cụm hai ký tự. Nếu thay `¢` trước, cụm `¢Ê` sẽ bị phá và không còn cơ hội chuyển thành `Ấ`. Hai pha thay mỗi chuỗi nguồn bằng placeholder `::i::`, sau đó thay placeholder bằng Unicode đích. Cách này cô lập các lần thay thế và tránh chồng lấn.

**Nếu bị hỏi sâu:**

- Về bản chất, đây là cách xử lý **longest-match-first** một cách tường minh và dễ debug.
- Có thể dùng regex alternation hoặc `str.translate`; nhưng `str.translate` chỉ ánh xạ 1-1, không xử lý cụm.
- Placeholder `::i::` có thể va chạm nếu dữ liệu gốc chứa đúng chuỗi đó; với dữ liệu đất, rủi ro rất thấp. Có thể đổi sang ký tự private-use (`\uE000…`) nếu cần.
- Test hồi quy kiểm tra các cụm `¢`, `¢Ê`, `Aµ`, `A¸`…

**Câu chốt:** *“Hai pha không phải để làm phức tạp, mà để bảo toàn cụm dài trước khi thay ký tự đơn.”*

---

# Nhóm B — Kỹ thuật nhận diện và xử lý dữ liệu

## Q6. Nhận diện TCVN3 dựa trên font và mẫu ký tự. Nếu file không còn font metadata thì sao? Có false positive không?

**Trả lời ngắn:**

Khi thiếu font, chương trình dùng mẫu ký tự: marker mạnh (`¸`, `©`, `§`, `¢`…) hoặc từ hai marker yếu trở lên. Nếu không đủ chắc, chương trình **giữ nguyên** thay vì đoán. Cách này giảm false positive với các chữ Unicode hợp lệ như `núi`, `báo cáo`, `µm`, nhưng vẫn có thể bỏ sót (false negative).

**Nếu bị hỏi sâu:**

- Đây là heuristic, không phải bộ phân loại đã huấn luyện.
- Marker mạnh gần như không xuất hiện trong văn bản Unicode tiếng Việt; marker yếu (`Ì`, `Ð`, `Ý`, `Ó`, `Ê`…) có thể trùng, nên cần ≥ 2 lần.
- Trường hợp khó: một ô chỉ có `µ` (có thể là micro-mét) — chúng tôi không chuyển nếu không có font/marker khác.
- Đề xuất: log các ô “nghi ngờ”, cho phép người dùng review, và bổ sung bộ test với câu tiếng Việt đã biết.

**Câu chốt:** *“Nguyên tắc của chúng tôi là thà giữ nguyên còn hơn chuyển sai; dữ liệu sai mã còn nguy hiểm hơn dữ liệu chưa chuyển.”*

---

## Q7. Nếu một ô trộn Unicode và TCVN3, hoặc dùng VNI/VIQR thì sao?

**Trả lời ngắn:**

Với ô trộn, chương trình chuyển khi có tín hiệu TCVN3; các ký tự không nằm trong bảng ánh xạ được giữ nguyên. Tuy nhiên đây là trường hợp có rủi ro và cần kiểm tra thủ công. VNI/VIQR **chưa được hỗ trợ** trong phạm vi bài; chúng tôi xem đây là hướng mở rộng.

**Nếu bị hỏi sâu:**

- TCVN3, VNI, VIQR dùng các quy ước khác nhau; không thể dùng chung một bảng ánh xạ.
- Nếu nhận diện sai bảng mã, hậu quả nặng hơn là không chuyển.
- Đề xuất: thêm tham số `--encoding` hoặc tự nhận diện theo font + mẫu, kèm bảng ánh xạ riêng.
- Với ô trộn, có thể tách thành các đoạn Unicode/TCVN3 trước khi chuyển; hiện tại ưu tiên bảo toàn và ghi log.

**Câu chốt:** *“Bài chỉ giải quyết TCVN3 — bài toán đã đủ lớn; VNI/VIQR là hướng mở rộng, không phải điều chúng tôi tuyên bố đã làm.”*

---

## Q8. CSV vẫn giữ cấu trúc lưới. Vậy đã là “dữ liệu huấn luyện AI” chưa?

**Trả lời ngắn:**

Chưa. CSV trong bài là **dữ liệu trung gian đã chuẩn hóa mã ký tự và định dạng**, không phải bảng huấn luyện hoàn chỉnh. Muốn huấn luyện AI cần thêm bước: nhận diện vùng bảng, ánh xạ schema cột, xử lý đơn vị/missing/outlier, gắn nhãn và chia train/validation. Bài báo nói rõ đây là nền tảng cho các bước tiếp theo.

**Nếu bị hỏi sâu:**

- Giữ cấu trúc lưới là lựa chọn có chủ đích: bảo toàn thông tin gốc và khả năng truy vết.
- Nếu ép ngay về schema chuẩn, có thể mất dữ liệu ở các bảng có cấu trúc khác nhau.
- “Sẵn sàng cho AI” trong bài nên hiểu là **sẵn sàng cho bước làm sạch/schema**, không phải “đưa thẳng vào model”.
- Đây là lý do chúng tôi tách `normalize.py` — bước gộp dọc — nhưng vẫn chưa phải schema cuối.

**Câu chốt:** *“CSV là đầu vào đã sạch mã ký tự, chưa phải đầu vào đã sạch ngữ nghĩa.”*

---

## Q9. Xử lý merged cells, công thức, ngày tháng, nhiều dòng tiêu đề, sheet ẩn thế nào?

**Trả lời ngắn:**

- `.xlsx`: đọc `data_only=True`, lấy giá trị đã tính; công thức chưa cache có thể trả `None`.
- `.xls`: đọc `cell_value`; ngày tháng có thể ra serial/float, chưa chuẩn hóa.
- Merged cells: chỉ ô top-left giữ giá trị, các ô còn lại thường là `None`; CSV vẫn giữ lưới.
- Nhiều dòng tiêu đề: `normalize.py` ghép tối đa 5 dòng tiêu đề, nhưng không suy luận ngữ nghĩa.
- Sheet ẩn: hiện vẫn được duyệt nếu có dữ liệu; có thể thêm tùy chọn bỏ qua `ws.sheet_state != "visible"`.

**Nếu bị hỏi sâu:**

- Ảnh, biểu đồ, định dạng ô không được đưa vào CSV; mục tiêu là dữ liệu số/chữ.
- Với công thức, nếu file chưa được Excel tính, cần mở bằng Excel/LibreOffice để cache giá trị hoặc tính lại.
- Ngày tháng nên chuẩn hóa theo kiểu dữ liệu và đơn vị; hiện là hạn chế.
- Đề xuất: thêm `--skip-hidden`, `--date-format`, `--header-depth` cho các bộ dữ liệu khác.

**Câu chốt:** *“Chúng tôi xử lý dữ liệu ô, không xử lý toàn bộ ngữ nghĩa của workbook; điều này cần nói rõ để tránh kỳ vọng sai.”*

---

# Nhóm C — Số liệu, kiểm chứng, hiệu năng

## Q10. 166 file, 963 sheet, 20.532 mẫu được đếm thế nào? “Mẫu hợp lệ” là gì?

**Trả lời ngắn:**

- **166** là số workbook đầu vào.
- **963** là số sheet có dữ liệu, mỗi sheet tạo một CSV.
- **20.532** là số dòng mẫu trong `AllData.csv` sau bước gộp dọc, không phải số ô.
- Một dòng được coi là “mẫu” khi bảng có anchor dạng báo cáo thí nghiệm (Thành phần hạt, Độ ẩm, Chiều sâu…), có STT, số hiệu mẫu/hố khoan và ít nhất một chỉ tiêu phía sau.

**Nếu bị hỏi sâu:**

- Đây là **heuristic theo cấu trúc**, không phải định nghĩa chuẩn của một mẫu đất.
- Các bảng dọc, bảng tóm tắt, trang bìa, metadata bị loại.
- Một sheet có thể có nhiều mẫu; một workbook có nhiều sheet.
- Con số 20.532 không nên so sánh trực tiếp với “số mẫu đất ngoài thực tế” nếu chưa kiểm tra trùng lặp.

**Câu chốt:** *“Chúng tôi đếm theo đơn vị dòng mẫu trong bảng đã chuẩn hóa, không đếm ô hay file.”*

---

## Q11. Số liệu 963 sheet / 20.532 mẫu có tái lập được không? Vì sao chạy lại có thể khác?

**Trả lời ngắn:**

Số liệu trong bài là của **lần chạy trước**. Trên snapshot dữ liệu hiện hành, chúng tôi chạy lại được **962 CSV / 20.352 dòng**. Kết quả phụ thuộc vào: phiên bản dữ liệu đầu vào, sheet rỗng/có dữ liệu, phiên bản heuristic normalize và cách đếm. Để tái lập, cần ghi rõ ngày chạy, git commit, hash dữ liệu và log.

**Nếu bị hỏi sâu:**

- Đây không phải mâu thuẫn khoa học; đây là vấn đề **version hóa dữ liệu**.
- Nếu bổ sung một hồ sơ mới, số sheet và số dòng sẽ thay đổi.
- Nếu một sheet trước có dữ liệu, nay bị làm rỗng, số CSV giảm.
- Nếu file `AllData.csv` bị tính vào số CSV, con số sheet có thể lệch 1.
- Đề xuất: lưu `run_manifest.json` gồm danh sách file, hash, phiên bản thư viện, số dòng, thời gian chạy.

**Câu chốt:** *“Con số trong bài gắn với một snapshot; khi công bố chúng tôi sẽ ghi rõ snapshot và cập nhật con số cuối cùng.”*

---

## Q12. 166/166 thành công — tỷ lệ 100% có đáng tin? File hỏng hoặc có mật khẩu thì sao?

**Trả lời ngắn:**

100% chỉ đúng với **bộ dữ liệu thử nghiệm đã được lọc**. Chương trình có xử lý lỗi: file ZIP hỏng, file `.xls` bị mã hóa hoặc không đọc được sẽ được ghi `FAIL` và bỏ qua, không làm dừng cả batch. Trong bộ này không có file lỗi, nên tỷ lệ là 166/166.

**Nếu bị hỏi sâu:**

- Đây là tỷ lệ **trên tập đã chọn**, không phải cam kết cho mọi kho dữ liệu.
- File bị mật khẩu cần mở khóa thủ công hoặc có mật khẩu; không nên cố vượt qua.
- File `.xlsx` hỏng cấu trúc ZIP không thể phục hồi bằng openpyxl.
- Nên báo cáo cả **denominator** (tổng số file) và danh sách lỗi, thay vì chỉ nói “100%”.

**Câu chốt:** *“100% là kết quả trên bộ thử nghiệm; công cụ được thiết kế để thất bại có kiểm soát, không phải để đọc được mọi file hỏng.”*

---

## Q13. Hiệu năng thế nào? Nếu có 10.000 file thì sao?

**Trả lời ngắn:**

Trên máy thử, 166 file chạy khoảng 1,5–2,5 phút, tùy kích thước workbook. Kiến trúc hiện tại xử lý tuần tự từng workbook, bộ nhớ theo từng file. Với 10.000 file, cần song song hóa theo file, thêm hàng đợi, retry và theo dõi tiến trình; đây là hướng mở rộng, chưa triển khai trong bài.

**Nếu bị hỏi sâu:**

- Nút thắt là openpyxl/xlrd khi parse workbook lớn, không phải ghi CSV.
- Có thể dùng `multiprocessing`/`ProcessPoolExecutor` theo file; mỗi process một workbook.
- Cần giới hạn số process theo RAM/CPU, vì workbook lớn có thể chiếm nhiều bộ nhớ.
- GUI hiện chạy một worker tuần tự; có thể thêm nhiều worker nhưng phải quản lý log/tiến trình.
- Với dữ liệu cực lớn, có thể chuyển sang streaming hoặc tiền xử lý bằng LibreOffice headless.

**Câu chốt:** *“Hiện tại ưu tiên đúng và tái lập hơn tốc độ; song song hóa là bước kỹ thuật tiếp theo.”*

---

## Q14. Kiểm soát chất lượng số — đơn vị, missing, outlier, trùng lặp — đã làm chưa?

**Trả lời ngắn:**

Chưa. Trong phạm vi bài, chúng tôi giữ nguyên giá trị số gốc, chỉ chuẩn hóa cột góc DDMM → độ thập phân. Kiểm tra đơn vị, missing, outlier và trùng lặp là bước tiếp theo, chưa triển khai.

**Nếu bị hỏi sâu:**

- Đề xuất rule-based range check theo tiêu chuẩn thí nghiệm: độ ẩm 0–100%, khối lượng thể tích ~1,2–2,5 g/cm³, góc ma sát 0–45°, lực dính 0–200 kPa…
- Missing: báo cáo tỷ lệ theo cột, không tự điền.
- Outlier: IQR hoặc robust z-score, giữ lại giá trị gốc và gắn cờ.
- Trùng lặp: khóa `(workbook, sheet, stt, do_sau_tu, do_sau_den, so_hieu_mau)`.
- Đơn vị: từ điển đơn vị + kiểm tra header; không tự đổi đơn vị khi chưa chắc.

**Câu chốt:** *“Hiện tại chúng tôi mới sạch mã ký tự; sạch số là bài toán khác và là hướng tiếp theo.”*

---

# Nhóm D — Ứng dụng AI, chất lượng dữ liệu, đạo đức

## Q15. Bài nói “20% dữ liệu gán nhãn, 20% kiểm chứng” — hiểu thế nào? Có data leakage không?

**Trả lời ngắn:**

Đó là **ví dụ minh họa cách chia** cho mô hình, không phải kết quả huấn luyện trong bài. Bài chưa huấn luyện mô hình AI. Khi làm thật, cần chia theo **công trình/hố khoan** thay vì chia ngẫu nhiên theo dòng, để tránh leakage giữa train và test.

**Nếu bị hỏi sâu:**

- Các dòng trong cùng một hố khoan/công trình có tương quan mạnh; chia ngẫu nhiên theo dòng sẽ làm test “quá dễ”.
- Nên dùng `GroupKFold` theo `workbook_key` hoặc `ma_ho_khoan`.
- 20% labeled/20% validation chỉ là tỷ lệ minh họa; tỷ lệ thực phụ thuộc bài toán.
- Cần kiểm tra trùng mẫu trước khi chia.

**Câu chốt:** *“Bài này dừng ở chuẩn bị dữ liệu; chia train/test là thiết kế cho nghiên cứu tiếp, không phải kết quả đã chạy.”*

---

## Q16. 20.532 mẫu đã đủ để huấn luyện AI chưa?

**Trả lời ngắn:**

20 nghìn mẫu là khởi đầu tốt cho mô hình dạng bảng (tabular), nhưng “đủ” còn phụ thuộc số đặc trưng, độ phân giải nhãn, mức nhiễu và mô hình. Với học sâu, có thể cần nhiều hơn hoặc dữ liệu chất lượng hơn. Bài chỉ nói đây là nền tảng bước đầu, không khẳng định đủ cho mọi mô hình.

**Nếu bị hỏi sâu:**

- Với bài toán hồi quy vài chỉ tiêu cơ lý, 20 nghìn dòng có thể đủ cho baseline.
- Vấn đề lớn hơn là **chất lượng nhãn** và **tính độc lập** giữa các mẫu.
- Nếu nhãn đến từ cùng một quy trình thí nghiệm, mô hình có thể học đặc trưng của phòng thí nghiệm thay vì vật lý đất.
- Cần baseline, cross-validation theo nhóm, và kiểm tra ngoài miền (external validation).

**Câu chốt:** *“20 nghìn mẫu là đủ để bắt đầu, chưa đủ để khẳng định độ tin cậy nếu chưa kiểm chứng ngoài miền.”*

---

## Q17. So với dùng AI/LLM để trích xuất bảng, cách rule-based này có lỗi thời?

**Trả lời ngắn:**

Không. Với dữ liệu cần kiểm toán và tái lập, rule-based có lợi thế: deterministic, chi phí thấp, không hallucinate, chạy offline. AI/LLM có thể hỗ trợ nhận diện vùng bảng hoặc suy luận schema ở bước sau, nhưng cần ground truth và kiểm chứng. Hướng phù hợp là **kết hợp rule-based + AI**, không thay thế hoàn toàn.

**Nếu bị hỏi sâu:**

- LLM có thể đọc hiểu tiêu đề bảng tốt hơn, nhưng khó đảm bảo 100% nhất quán trên 20 nghìn dòng.
- Rule-based dễ debug, dễ giải thích với hội đồng và đơn vị khảo sát.
- AI phù hợp cho bước **schema mapping** và **phát hiện bất thường**, nơi cần suy luận ngữ nghĩa.
- Có thể dùng LLM để tạo candidate mapping, sau đó con người/rule xác nhận.

**Câu chốt:** *“Chúng tôi dùng rule-based cho phần cần chắc chắn, và xem AI là lớp hỗ trợ ngữ nghĩa ở bước sau.”*

---

## Q18. Dữ liệu địa kỹ thuật nhạy cảm. Tọa độ, tên công trình, mã hố khoan được xử lý thế nào?

**Trả lời ngắn:**

Bài nêu rõ dữ liệu hiện **không bao gồm tọa độ công trình**. Tuy nhiên tên file/sheet có thể chứa địa danh, tên công trình, mã hố khoan. Công cụ chạy cục bộ, không gửi dữ liệu lên cloud. Khi chia sẻ, cần ẩn danh tên file, mã hố khoan và kiểm soát quyền truy cập.

**Nếu bị hỏi sâu:**

- Log có thể chứa tên file; cần kiểm tra trước khi công bố.
- Nếu dữ liệu thuộc đơn vị khác, cần thỏa thuận chia sẻ và tuân thủ quy định dữ liệu.
- Không nên đưa `dataXLSX/`, `dataCSV_UTF8/`, `AllData.csv` lên GitHub.
- Repo chỉ chứa dữ liệu mẫu tổng hợp trong `examples/`.

**Câu chốt:** *“Công cụ xử lý cục bộ; trách nhiệm ẩn danh và chia sẻ dữ liệu thuộc về quy trình quản trị dữ liệu của đơn vị.”*

---

# Nhóm E — Tái lập, công bố, hướng phát triển

## Q19. Tính tái lập và minh bạch được đảm bảo thế nào?

**Trả lời ngắn:**

Dự án pin phiên bản thư viện (`xlrd==1.2.0`, `openpyxl`), có test tự động, log OK/FAIL, quy ước tên file và mã nguồn mở. Để tái lập đầy đủ, cần thêm **manifest** gồm git commit, phiên bản Python/thư viện, hash dữ liệu, ngày chạy và danh sách file lỗi. Đây là khuyến nghị đang được bổ sung.

**Nếu bị hỏi sâu:**

- Kết quả có thể thay đổi nếu dữ liệu đầu vào thay đổi; cần version hóa dữ liệu.
- Không có random trong pipeline, nên kết quả deterministic với cùng input/phiên bản.
- CSV dùng UTF-8 BOM để tránh lỗi hiển thị khi mở bằng Excel.
- Repo có `CITATION.cff`, `LICENSE`, README và hướng dẫn chạy.

**Câu chốt:** *“Tái lập không chỉ là cùng code, mà là cùng code + cùng dữ liệu + cùng phiên bản thư viện; chúng tôi đang chuẩn hóa cả ba.”*

---

## Q20. Hướng phát triển nào quan trọng nhất, và đo lường thành công ra sao?

**Trả lời ngắn:**

Ưu tiên 1 là **schema chuẩn + kiểm soát chất lượng số**; ưu tiên 2 là pipeline ML với chia nhóm theo công trình; ưu tiên 3 là báo cáo/version hóa tự động. Thành công được đo bằng: tỷ lệ sheet map được sang schema, tỷ lệ missing/outlier phát hiện, chất lượng mô hình baseline (RMSE/R²) và khả năng tái lập.

**Nếu bị hỏi sâu:**

- Schema: định nghĩa cột chỉ tiêu cố định, ánh xạ từ tiêu đề sheet.
- QC: range check, unit dictionary, missing report, outlier flags, dedup.
- ML: `GroupKFold` theo công trình/hố khoan, baseline + external validation.
- Vận hành: `run_manifest.json`, log có cấu trúc, báo cáo HTML/PDF tự động.
- Đo lường: không chỉ số dòng, mà là **tỷ lệ dữ liệu dùng được** và **độ tin cậy mô hình**.

**Câu chốt:** *“Bước tiếp theo không phải chạy thêm model, mà là làm cho dữ liệu có schema và kiểm chứng được — nếu không, 20 nghìn dòng vẫn chỉ là CSV thô.”*

---

## Ba câu chốt nên nhớ

1. **Về phạm vi:** “Bài tập trung vào chuẩn hóa mã ký tự và định dạng; chưa huấn luyện mô hình AI.”
2. **Về giới hạn:** “Chúng tôi đã nêu rõ chưa có schema thống nhất, chưa kiểm tra đơn vị/outlier; đó là hướng tiếp theo.”
3. **Về giá trị:** “Điểm mạnh là xử lý dữ liệu TCVN3 lịch sử, chạy batch, có log và có thể tái lập — đúng phần khó nhất của dữ liệu địa kỹ thuật Việt Nam.”

---

## Phụ lục: câu hỏi có thể bị hỏi ngoài luồng

- **“Sao không dùng `pandas.read_excel` cho nhanh?”**  
  `pandas` không giữ font metadata của từng ô trong `.xls`; cần `xlrd` để nhận diện TCVN3. Ngoài ra `pandas` mặc định không tách sheet theo quy ước của bài.

- **“Sao không xuất Parquet?”**  
  CSV UTF-8 dễ mở, dễ kiểm tra bằng Excel và phù hợp đối tượng người dùng địa kỹ thuật. Parquet là hướng tối ưu cho pipeline lớn.

- **“Sao không chuẩn hóa Unicode NFC?”**  
  Đây là bước nên thêm; hiện tại bảng ánh xạ trả về ký tự dựng sẵn (NFC) cho các cụm phổ biến. Cần kiểm tra toàn bộ output bằng `unicodedata.normalize`.

- **“Có hỗ trợ file `.xlsm` không?”**  
  Có thể đọc như `.xlsx` nếu đổi phần mở rộng/lọc; hiện `discover_workbooks` chỉ lấy `.xls`/`.xlsx`. Cần bổ sung `.xlsm` nếu dữ liệu có macro.

- **“Có hỗ trợ file bị khóa bằng mật khẩu không?”**  
  Không. Cần mở khóa thủ công; đây là giới hạn có chủ đích vì lý do bảo mật.
