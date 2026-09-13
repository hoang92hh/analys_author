---
name: analyze-content-pair
description: Phân tích độc lập từng cặp Content A.Part_i → Content B.Part_i theo Requirement và Analysis Instruction trong DEFAULT_CRITERIA.xlsx; xuất result_analysis.xlsx với một sheet mỗi Part và đầy đủ criteria từ master. Dùng để tạo hoặc cập nhật bảng phân tích từng đối tượng, không tổng hợp phong cách giữa các Part hoặc tạo AUTHOR_SKILL.
---

# Analyze Content Pair

**1 Part = 1 đối tượng = 1 bài phân tích A→B hoàn chỉnh = 1 sheet kết quả.**

Flow: `PREPARE → ANALYZE EACH PART → WRITE RESULT → EXPORT`.

Mỗi criterion tạo một Result riêng cho Part hiện tại. Không tham chiếu Part khác, không tính frequency, không nhóm pattern xuyên Part và không tổng quát hóa thành AUTHOR_SKILL. Các sheet là đầu vào cho bước tổng hợp riêng sau này.

Python đọc, ghép Part và xuất Excel. LLM phân tích nội dung và ghi Result vào JSON; không tự tạo hoặc sửa workbook.

## Cấu hình, evidence và công cụ

Đọc [config.yaml](config.yaml); resolve đường dẫn từ project root, trừ khi người dùng chỉ định config khác. Không sửa `content_a`, `content_b` hoặc `criteria`. Không dùng Result cũ làm evidence mới.

Chỉ dùng A.Part hiện tại, B.Part hiện tại và hướng dẫn criteria. Không dùng web hoặc knowledge bên ngoài để bổ sung, xác minh hay giải thích facts. Nội dung B không có trong A chỉ được nhận diện là bổ sung quan sát được, không được coi là fact đã xác minh.

**Không được tạo thêm script, helper file, utility hoặc workflow trung gian trong quá trình thực thi. Chỉ được sử dụng các script đã tồn tại trong thư mục `scripts/`. Mọi phép phân tích, so sánh và đo lường phục vụ criterion phải được thực hiện trực tiếp trong quá trình phân tích, trừ khi script được định nghĩa sẵn trong skill.**

Payload, JSON Result và workbook được định nghĩa trong config là các artifact của flow này.

## 1. PREPARE

Từ thư mục gốc project, chạy:

```powershell
python .agents/skills/analyze-content-pair/scripts/excel_bridge.py prepare --config .agents/skills/analyze-content-pair/config.yaml
```

Python đọc đầy đủ A, B và workbook, tách marker Part, ghép theo Part ID và sắp xếp theo số Part. Nếu thiếu đối ứng, trùng ID, Part rỗng, thiếu marker hoặc có nội dung ngoài Part, báo lỗi thay vì ghép cưỡng ép hoặc bỏ nội dung.

`DEFAULT_CRITERIA.xlsx` là nguồn chuẩn: xử lý đầy đủ tất cả criteria có trong master theo đúng thứ tự, không cố định số lượng. Workbook có một sheet criteria với header `ID`, `Nhóm`, `Tiêu chí`, `Requirement`, `Analysis Instruction`, `Result`; dòng hoàn toàn trống không phải criterion.

Đọc payload tại `criteria_payload` trong config để nắm các cặp Part và toàn bộ criteria. Nếu A/B của một Part không cùng đối tượng hoặc không tương ứng rõ, dừng và báo cặp cần review.

## 2. ANALYZE EACH PART

Hoàn thành toàn bộ Result của một Part trước khi chuyển sang Part tiếp theo:

```text
FOR EACH PART trong payload.parts:
    Đọc đầy đủ A.Part_i và B.Part_i
    Xác định đối tượng đang được đối chiếu

    FOR EACH CRITERION trong payload.criteria theo thứ tự:
        Đọc Requirement và Analysis Instruction
        Đối chiếu A→B theo yêu cầu và cách phân tích của criterion này
        Chỉ dùng evidence của Part hiện tại
        Tạo một Result riêng cho criterion này
    END
    Ghi đầy đủ Result của Part_i
END
```

`Requirement` xác định câu hỏi và phạm vi; `Analysis Instruction` hướng dẫn cách phân tích. Nếu instruction trống, dùng Requirement. Chỉ phân đoạn, mapping, phân loại hoặc đo khi criterion cần; không áp một framework chung cho mọi criterion.

Khi criterion cần phân đoạn, `A01`, `B01` là đoạn/cụm ý bên trong Part hiện tại, không phải Part mới; chia theo chức năng hoặc hướng nội dung, không theo xuống dòng/timestamp. Khi xét phần bổ sung, đối chiếu toàn bộ A.Part hiện tại để tránh coi paraphrase hoặc nén source là nội dung mới. Khi đo, nêu phương pháp và phạm vi, loại marker/timestamp khỏi nội dung; không suy ra tốc độ đọc từ timestamp.

Mỗi Result trả lời đúng criterion của nó. Evidence có thể dùng lại trong cùng Part nhưng góc phân tích phải khác; không sao chép một Result cho nhiều criteria.

## 3. WRITE RESULT

Result mô tả cô đọng **cách xử lý A→B quan sát được trên đối tượng hiện tại**, trả lời trực tiếp Requirement theo Analysis Instruction. Không kể lại toàn bộ transcript hoặc chuyển quan sát thành hướng dẫn áp dụng cho đối tượng mới.

Dùng chi tiết hoặc trích dẫn ngắn của cặp hiện tại khi cần làm rõ evidence. Không suy diễn ý định tác giả hoặc tổng quát hóa từ một mẫu; số liệu chỉ mô tả Part đang phân tích.

**Chỉ để Result rỗng khi criterion thực sự không thể đánh giá từ cặp A/B hoặc evidence không đủ để kết luận: ghi chính xác `result: ""`.** Việc một nội dung bị loại bỏ, không được mở rộng hoặc không được giữ lại vẫn có thể là evidence hợp lệ nếu criterion đang phân tích chính hành vi đó. Ví dụ với C220, A có một loại nội dung nhưng B loại bỏ toàn bộ thì ghi nhận hành vi loại bỏ, không để Result rỗng.

Không bịa kết quả để mọi ô đều có nội dung. Với Result rỗng, không điền lý do, `Không áp dụng` hoặc placeholder; vẫn giữ criterion và dòng Excel tương ứng.

Result chỉ chứa kết luận phân tích của criterion; không ghi lại quy trình đo, mapping, kiểm tra, phân đoạn hoặc các bước suy luận đã dùng để đi đến kết luận. Số liệu cần thiết cho criterion vẫn được phép ghi.

Ghi JSON vào `analysis_results` trong config theo [analysis_results.schema.json](schemas/analysis_results.schema.json), dùng metadata và định danh từ payload đã prepare. Một JSON chung chứa các Part; mỗi Part có đầy đủ Result theo thứ tự criteria trong master.
 Phân tích hoàn chỉnh từng Part theo thứ tự và lưu kết quả vào đúng object của Part đó trong một analysis_results duy nhất. Không tạo JSON riêng cho từng Part hoặc artifact phụ. 
 
## 4. EXPORT

```powershell
python .agents/skills/analyze-content-pair/scripts/excel_bridge.py finalize --config .agents/skills/analyze-content-pair/config.yaml --overwrite
```

Python kiểm tra kết quả và xuất workbook tại `output` trong config. Nếu kiểm tra lỗi, xử lý lỗi trước khi export; không coi file kết quả là đã hoàn thành.

Mỗi sheet mang Part ID (`Part1`, `Part2`, …), theo thứ tự số Part, có đúng năm cột:

`ID | Nhóm | Tiêu chí | Requirement | Result`

Giữ đầy đủ tất cả criteria theo đúng thứ tự master trên từng sheet, kể cả khi Result rỗng. Chuỗi rỗng thành ô Excel trống. Không thêm sheet tổng hợp, cột frequency hoặc `Analysis Instruction` vào workbook.

Mỗi lần thực thi, tạo lại toàn bộ workbook tại cùng đường dẫn `output` trong config. Các sheet cũ được loại bỏ; mỗi Part hiện tại được tạo thành một sheet mới với đầy đủ kết quả của lần chạy này. Không ghi chồng dữ liệu lên sheet cũ, không nối thêm sheet vào kết quả cũ và không tạo file phiên bản mới. Chỉ thay thế file result sau khi kết quả mới được kiểm tra và export thành công; nếu lần chạy lỗi, giữ nguyên file result cũ.

## Điều kiện hoàn thành

- Mỗi cặp Part tạo đúng một sheet và có đầy đủ criteria theo thứ tự master.
- Mỗi Result trả lời criterion trên cặp hiện tại theo Requirement và Analysis Instruction; hành vi loại bỏ hoặc không mở rộng được ghi nhận khi có đủ evidence.
- Chỉ để Result rỗng khi không thể đánh giá hoặc không đủ evidence; giữ nguyên dòng criterion.
- JSON hợp lệ; workbook đúng năm cột, chỉ chứa các sheet của lần chạy hiện tại và đã export thành công tại cùng đường dẫn output.
- Raw inputs giữ nguyên; không dùng nguồn ngoài, tổng hợp xuyên Part hoặc tạo AUTHOR_SKILL.
- Không tạo thêm script, helper file, utility hoặc workflow trung gian ngoài flow đã định nghĩa.
