# AGENTS.md — Author Skill System

## 1. Phạm vi áp dụng

Tệp này áp dụng cho toàn bộ project trong thư mục `analys/`.

Chỉ sử dụng dữ liệu, mã nguồn và tài liệu nằm trong `analys/`, trừ khi người dùng chỉ định rõ một nguồn bên ngoài. Không suy diễn kiến trúc hoặc quy ước từ các project lân cận.

## 2. Mục tiêu project

Xây dựng hệ thống phân tích cách một tác giả biến **Content A** thành **Content B**, rồi tổng quát hóa các hành vi có bằng chứng thành một **AUTHOR_SKILL** có thể tái sử dụng cho nội dung nguồn mới.

Luồng học:

```text
A + B + DEFAULT_CRITERIA
→ CREATE_SKILL
→ AUTHOR_SKILL
```

Luồng áp dụng:

```text
A1 + AUTHOR_SKILL → B1
A2 + AUTHOR_SKILL → B2
```

Hệ thống phải học quá trình ra quyết định và pattern chuyển đổi của tác giả. Không học thuộc, sao chép hoặc gắn cứng facts riêng của A/B vào rule dùng cho đối tượng mới.

## 3. Dữ liệu hiện có

```text
analys/
├── AGENTS.md
├── criteria/
│   └── DEFAULT_CRITERIA.xlsx
└── reference/
    ├── oldContent.txt
    └── newContent.txt
```

Vai trò mặc định:

- `reference/oldContent.txt`: Content A, nội dung nguồn ban đầu.
- `reference/newContent.txt`: Content B, nội dung tác giả tạo từ A và có thể có thông tin bổ sung.
- `criteria/DEFAULT_CRITERIA.xlsx`: master list các tiêu chí có thể dùng để phân tích quan hệ A ↔ B.

Không sửa ba tệp nguồn trên nếu người dùng không yêu cầu rõ ràng. Khi cần thử nghiệm, tạo đầu ra hoặc fixture riêng.

## 4. Khái niệm cốt lõi

- **DEFAULT_CRITERIA** là tập tiêu chí có thể kiểm tra, không phải template cố định của mọi AUTHOR_SKILL.
- **CREATE_SKILL** là quy trình đối chiếu A, B và từng criterion để tìm pattern có bằng chứng.
- **AUTHOR_SKILL** là tập rule động thực sự được kích hoạt từ A+B.
- **Observation** mô tả điều xảy ra trong cặp A/B cụ thể.
- **Generalized Rule** chuyển observation thành hướng dẫn có thể áp dụng cho đối tượng mới.
- **Executable Instruction** phải cho biết khi nào áp dụng, phải làm gì và chọn kết quả như thế nào.

## 5. Nguyên tắc bắt buộc

1. Không đưa toàn bộ DEFAULT_CRITERIA vào AUTHOR_SKILL.
2. Chỉ kích hoạt criterion khi có bằng chứng rõ trong quan hệ A → B.
3. Mọi rule sinh ra phải giữ `Criterion ID` nguồn.
4. Không biến tên riêng, con số, sự kiện hoặc đối tượng của A/B thành rule tổng quát.
5. Không giả định mọi loại đối tượng dùng chung một content schema.
6. Content schema phải được suy ra từ những dimensions tác giả thực sự khai thác trong B.
7. Không chỉ ghi retention ratio; phải mô tả loại thông tin được giữ, bỏ và ưu tiên.
8. Thông tin B bổ sung ngoài A phải được phân loại theo loại thông tin, không lưu fact cụ thể làm rule.
9. Không tạo rule khi bằng chứng yếu, mơ hồ hoặc không thể tổng quát hóa.
10. AUTHOR_SKILL cuối cùng phải chạy độc lập, không cần lại B, DEFAULT_CRITERIA hoặc CREATE_SKILL.

## 6. Quy trình CREATE_SKILL

Thực hiện theo thứ tự:

1. Đọc đầy đủ A và B.
2. Nhận diện đơn vị nội dung tương ứng giữa A và B.
3. Xác định đối tượng chính, loại chủ đề và các information dimensions.
4. Đọc schema và toàn bộ criterion trong DEFAULT_CRITERIA.
5. Kiểm tra từng criterion phù hợp với dữ liệu.
6. Ghi observation và evidence cụ thể từ A ↔ B.
7. Xác định `Activated: Yes/No` và confidence.
8. Loại criterion trùng lặp, không hữu ích hoặc không thể thực thi.
9. Chuyển observation thành pattern tổng quát.
10. Chuyển pattern thành rule độc lập, có điều kiện và hành động rõ ràng.
11. Gán rule vào đúng workflow step.
12. Xuất bảng phân tích trung gian để người dùng review.
13. Chỉ đưa rule đủ căn cứ vào AUTHOR_SKILL.

## 7. Bảng phân tích trung gian

Trước khi tạo AUTHOR_SKILL, phải tạo kết quả có tối thiểu các trường:

| Field | Ý nghĩa |
|---|---|
| Criterion ID | ID truy vết về DEFAULT_CRITERIA |
| Group | Nhóm criterion |
| Criterion | Tên criterion |
| Activated | Yes hoặc No |
| Evidence A↔B | Bằng chứng quan sát được |
| Result | Kết quả phân tích |
| Confidence | High, Medium hoặc Low |
| Generalized Rule | Rule đã loại bỏ chi tiết instance |
| Target Step | Workflow step sử dụng rule |

Evidence phải ngắn gọn nhưng đủ kiểm chứng. Không dùng nhận định chung chung như “văn phong tốt hơn” hoặc “nội dung hấp dẫn hơn”.

Mặc định:

- `High`: bằng chứng trực tiếp, lặp lại hoặc rất rõ.
- `Medium`: có bằng chứng hợp lý nhưng còn cách giải thích khác.
- `Low`: bằng chứng thiếu hoặc suy luận phụ thuộc nhiều giả định.

Không đưa rule `Low` vào AUTHOR_SKILL nếu chưa được người dùng duyệt.

## 8. Cấu trúc một Author Rule

Mỗi rule phải độc lập, dễ đọc và dễ sửa:

```markdown
## [C005 — Historical Context]

**Target step:** STEP 2 — RESEARCH & ENRICHMENT
**Confidence:** High

### WHEN
Điều kiện kích hoạt trên đối tượng hoặc source mới.

### ACTION
Hành động phân tích, nghiên cứu hoặc viết cần thực hiện.

### SELECTION
Điều kiện chọn, loại hoặc giới hạn kết quả.

### VALIDATION
Cách xác nhận output tuân thủ rule.
```

Không ghi facts của A/B vào `WHEN` hoặc `ACTION`. Evidence cụ thể thuộc bảng phân tích trung gian, không phải nội dung thực thi của rule.

## 9. Workflow của AUTHOR_SKILL

Workflow cấp cao có thể cố định; các rule bên trong phải động theo A+B.

### STEP 1 — SOURCE ANALYSIS

- Nhận diện đối tượng và loại chủ đề.
- Phân rã source thành facts và dimensions.
- Áp dụng retention, omission, priority và selection rules đã học.
- Tạo `Selected(A1)`.

### STEP 2 — RESEARCH & ENRICHMENT

- Chỉ chạy khi AUTHOR_SKILL có research rule được kích hoạt.
- Tìm dimensions còn thiếu mà tác giả thường bổ sung.
- Chỉ nghiên cứu thông tin tương ứng với đối tượng A1 hiện tại.
- Đánh giá độ liên quan, độ tin cậy và nguồn của thông tin.
- Không tìm lại facts riêng của A/B cũ.

### STEP 3 — CONTENT CONSTRUCTION & PRESENTATION

- Kết hợp `Selected(A1)` với external information đã chọn.
- Áp dụng các rule về biến đổi, diễn giải, reasoning, ordering, narrative, hook, transition, tone, vocabulary, sentence, paragraph, rhythm và ending nếu chúng thực sự tồn tại trong AUTHOR_SKILL.
- Tạo Draft B1.

### STEP 4 — VALIDATION

- Dùng chính các rule đang hoạt động làm rubric.
- Đánh giá mức tuân thủ author pattern, không so B1 với B theo độ giống câu chữ hoặc facts.
- Xác định rule vi phạm, sửa draft và kiểm tra lại.
- Tạo Final B1 khi các rule quan trọng đạt yêu cầu.

## 10. Content schema động

Trước khi áp dụng schema, phải nhận diện loại đối tượng: sự vật, con người, sự kiện, địa điểm, tổ chức, quy trình, khái niệm, hiện tượng hoặc loại khác được phát hiện.

Chỉ giữ những dimensions có bằng chứng cho thấy tác giả quan tâm. Không ép một schema về đồ vật lên sự kiện, con người hoặc quy trình.

Khi tạo rule schema, mô tả:

- loại đối tượng áp dụng;
- dimension cần phát hiện hoặc bổ sung;
- điều kiện dimension có giá trị;
- mức ưu tiên;
- giới hạn để tránh mở rộng lan man.

## 11. Source mapping và retention

Khi đối chiếu A → B, phân biệt tối thiểu:

- thông tin giữ gần nguyên bản;
- thông tin paraphrase;
- thông tin nén;
- thông tin mở rộng;
- thông tin thay đổi vị trí;
- thông tin bị bỏ;
- thông tin mới ngoài A.

Nếu ước lượng retention ratio, ghi cả phương pháp và selection pattern. Không coi một tỷ lệ đơn lẻ là rule hoàn chỉnh.

## 12. Research và tính xác thực

Khi rule yêu cầu nghiên cứu bên ngoài:

- ưu tiên nguồn gốc, nguồn chính thức hoặc nguồn có thẩm quyền;
- lưu URL, tiêu đề nguồn và ngày truy cập nếu hệ thống output hỗ trợ;
- phân biệt fact có nguồn với suy luận;
- không dùng một nguồn yếu để khẳng định fact quan trọng;
- không bổ sung thông tin chỉ vì thú vị nếu nó không phục vụ author pattern;
- không bịa dữ kiện để lấp dimension còn thiếu.

Nếu không xác minh được, bỏ fact hoặc đánh dấu cần review.

## 13. Chống overfitting và hallucination

Luôn thực hiện chuỗi chuyển đổi:

```text
INSTANCE OBSERVATION
→ RECURRING OR MEANINGFUL PATTERN
→ GENERALIZED RULE
→ EXECUTABLE INSTRUCTION
```

Trước khi chấp nhận rule, kiểm tra:

1. Có thể suy ra từ A+B không?
2. Evidence có trực tiếp và rõ ràng không?
3. Có trùng rule khác không?
4. Có áp dụng được cho đối tượng mới cùng loại không?
5. Có giúp xử lý A1/A2 không?
6. Có chuyển thành hành động kiểm chứng được không?
7. Có target workflow step rõ ràng không?

Nếu bất kỳ điểm cốt lõi nào không đạt, không kích hoạt rule.

## 14. Định dạng AUTHOR_SKILL.md

AUTHOR_SKILL cuối cùng nên có:

1. Tên và mô tả author pattern.
2. Phạm vi và loại đối tượng phù hợp.
3. Content schema đã học.
4. Các activated rule, nhóm theo workflow step.
5. Research policy nếu có research rule.
6. Construction instructions.
7. Validation rubric và threshold.
8. Những giới hạn hoặc trường hợp cần human review.

AUTHOR_SKILL không chứa criterion không kích hoạt. Không sao chép bảng evidence dài vào file thực thi nếu evidence đã được lưu ở artifact phân tích riêng.

## 15. Khả năng chỉnh sửa

Thiết kế output để hỗ trợ cả manual edit và AI-assisted edit:

- ID rule ổn định và duy nhất.
- Mỗi rule có thể thêm, sửa hoặc xóa độc lập.
- Threshold và retention ratio phải thể hiện rõ nếu có.
- Tránh phụ thuộc chéo không cần thiết.
- Khi sửa bằng yêu cầu tự nhiên, xác định rule theo ID trước khi thay đổi.
- Không tái đánh số ID nguồn trong DEFAULT_CRITERIA.

## 16. Quy tắc làm việc với file

- Đọc file văn bản theo UTF-8 và giữ nguyên ngôn ngữ gốc.
- Không silently sửa lỗi transcript trong source. Có thể chuẩn hóa ở lớp phân tích nhưng phải giữ raw source bất biến.
- Khi đọc workbook, giữ nguyên sheet, ID, header và kiểu dữ liệu nếu chỉ phân tích.
- Không ghi đè output đã được người dùng duyệt; tạo phiên bản mới hoặc thực hiện thay đổi có chủ đích.
- Tên output phải thể hiện rõ chức năng và không chứa facts riêng của đối tượng mẫu.
- Không tạo thêm cấu trúc thư mục chỉ để “đủ kiến trúc”; chỉ tạo khi có artifact thực sự cần lưu.

## 17. Kiểm thử và tiêu chuẩn hoàn thành

Một thay đổi chỉ hoàn thành khi các kiểm tra phù hợp đã đạt:

- mọi activated rule truy vết được về Criterion ID;
- evidence thực sự tồn tại trong A/B;
- rule không chứa chi tiết chỉ đúng với instance mẫu;
- rule có `WHEN`, `ACTION`, `SELECTION` và cách validation rõ;
- AUTHOR_SKILL không chứa toàn bộ DEFAULT_CRITERIA một cách máy móc;
- workflow chạy được với A1 mà không cần B;
- research chỉ chạy theo research rule;
- validation đo mức tuân thủ rule đang kích hoạt;
- raw inputs không bị thay đổi ngoài ý muốn.

Khi có fixture mới, ưu tiên kiểm thử ít nhất:

1. một đối tượng cùng loại với dữ liệu học;
2. một đối tượng có thiếu dimension cần research;
3. một trường hợp không thỏa điều kiện của một rule;
4. một trường hợp dễ gây sao chép facts từ mẫu cũ.

## 18. Cách Codex thực hiện nhiệm vụ trong project

- Trước khi sửa, xác định đúng file nguồn, artifact đầu ra và phạm vi người dùng yêu cầu.
- Chỉ đọc trong `analys/` trừ khi người dùng cho phép phạm vi khác.
- Bảo toàn thay đổi hiện có của người dùng và không sửa file không liên quan.
- Với yêu cầu phân tích, báo cáo evidence và kết luận; không tự động thay đổi dữ liệu nguồn.
- Với yêu cầu triển khai, thực hiện thay đổi, chạy kiểm tra phù hợp và báo rõ file đã tạo hoặc sửa.
- Nếu đặc tả và implementation mâu thuẫn, nêu bằng chứng cụ thể và ưu tiên ý định mới nhất của người dùng.
- Không tuyên bố một pattern đã được học nếu chưa có evidence hoặc kết quả validation.

## 19. Ngôn ngữ và thuật ngữ

- Trao đổi với người dùng bằng tiếng Việt, trừ khi họ yêu cầu ngôn ngữ khác.
- Giữ các tên chuẩn như `Content A`, `Content B`, `DEFAULT_CRITERIA`, `CREATE_SKILL`, `AUTHOR_SKILL`, `WHEN`, `ACTION`, `SELECTION`, `VALIDATION` và Criterion ID để tránh nhập nhằng.
- Nội dung rule có thể dùng tiếng Việt hoặc ngôn ngữ output mà người dùng chọn, nhưng cấu trúc và ID phải ổn định.

