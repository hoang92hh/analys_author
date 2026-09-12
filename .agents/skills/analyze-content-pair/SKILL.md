---
name: analyze-content-pair
description: Phân tích một hoặc nhiều cặp Content Aᵢ → Content Bᵢ theo từng criterion trong DEFAULT_CRITERIA.xlsx để tạo Result trung gian có thể tái sử dụng; Python phụ trách Excel và validation, còn LLM chỉ suy ra quy luật từ dataset. Dùng khi tạo hoặc cập nhật result_analysis.xlsx, không dùng để tạo AUTHOR_SKILL cuối cùng.
---

# Analyze Content Pair

Phân tích toàn bộ dataset `Content Aᵢ ↔ Content Bᵢ` để trả lời, với từng criterion: **“Tác giả triển khai criterion này như thế nào khi biến A thành B?”**

Đầu ra là các `Result` trung gian để bước sau tổng hợp thành transformation DNA. Không tạo `AUTHOR_SKILL` hoặc AUTHOR DNA cuối cùng trong workflow này.

Python sở hữu toàn bộ thao tác Excel, ánh xạ và lọc. LLM chỉ đọc dữ liệu và tạo JSON có định danh; không trực tiếp tạo, sửa hoặc lưu workbook.

## Cấu hình và ranh giới dữ liệu

Đọc [config.yaml](config.yaml). Resolve đường dẫn tương đối từ thư mục gốc project, trừ khi người dùng ghi đè trong yêu cầu hiện tại. Không sửa `content_a`, `content_b` hoặc `criteria`.

Chỉ dùng Content A, Content B và `criteria_payload` làm evidence. Không dùng web hay external knowledge để bổ sung, xác minh, giải thích hoặc suy ra transformation DNA. Nếu B chứa thông tin không có trong A, chỉ nhận diện loại nội dung mới quan sát được; không nghiên cứu thêm fact đó.

## Workflow bắt buộc

### 1. PREPARE — Python đọc Excel

Từ thư mục gốc project, chạy:

```powershell
python .agents/skills/analyze-content-pair/scripts/excel_bridge.py prepare --config .agents/skills/analyze-content-pair/config.yaml
```

Script đọc criteria theo đúng thứ tự workbook, không sửa file nguồn, và tạo `criteria_payload`. Mỗi dòng có `criterion_key` dạng `<sheet>!<row>:<criterion_id>` cùng `criterion_id`, `sheet`, `row`, `group`, `criterion` và `requirement`.

Nếu thiếu file đầu vào, ID trống/trùng, schema cột sai hoặc không có đúng một sheet criteria chứa `result_column`, dừng và báo lỗi. Không tự thay thế file hoặc cột khác.

### 2. LLM ANALYSIS — chỉ phân tích nội dung

Đọc đầy đủ Content A, Content B và toàn bộ `criteria_payload`. Nhận diện các segment tương ứng `Aᵢ → Bᵢ`; chỉ ghép cặp khi quan hệ đủ rõ, không ghép cưỡng ép. Với mỗi criterion, phân tích tất cả cặp áp dụng được rồi tạo **một Result tổng hợp**, không tạo Result riêng cho từng cặp.

#### Taxonomy biến đổi A → B

Khi criterion liên quan đến source mapping, coverage, transformation hoặc enrichment, phân biệt bốn trạng thái loại trừ nhau ở cấp semantic fact/proposition/idea unit:

1. **Giữ nguyên / gần nguyên văn:** nội dung A được đưa sang B gần trực tiếp.
2. **Giữ ý nhưng chuyển hóa:** ý/fact đã có trong A nhưng được paraphrase, rút gọn, tổng hợp, đổi ngôi, chuyển dialogue thành narration hoặc biến đổi hình thức khác. Không tính trạng thái này là nội dung mới.
3. **Nội dung mới thật sự:** fact, proposition, context, comparison, explanation hoặc significance trong B không tồn tại về mặt ngữ nghĩa trong A.
4. **Nội dung bị loại:** ý/fact có trong A nhưng không được dùng trong B.

`Reordering` là thuộc tính bổ sung của nội dung được giữ hoặc chuyển hóa, không phải trạng thái nội dung thứ năm. Một unit có thể vừa “giữ ý nhưng chuyển hóa” vừa “được chuyển vị trí”, nhưng không thể vì đổi cách diễn đạt mà trở thành “nội dung mới thật sự”.

Chỉ áp dụng taxonomy này khi nó phục vụ đúng `criterion` và `requirement`. Không bắt criterion về syntax, tone, engagement hoặc khía cạnh khác phải báo cáo giữ/bỏ/new/reordering.

#### Luật phân tích chung cho từng criterion

Thực hiện lần lượt theo thứ tự trong payload:

1. Đọc `criterion` và `requirement`; xác định câu hỏi riêng mà criterion đo, loại evidence liên quan và boundary với criteria khác.
2. Kiểm tra criterion có áp dụng cho dataset hay không. Chỉ quan sát các biến liên quan trực tiếp; `requirement` quyết định phạm vi phân tích.
3. Phân tích criterion riêng trên từng cặp áp dụng được trước khi tổng hợp cross-sample.
4. So sánh các cặp để tìm phần ổn định, khoảng biến thiên, trigger làm pattern thay đổi, conditional variation và ngoại lệ có hệ thống.
5. Khi evidence cho phép, rút ra các thành phần hữu ích: pattern chính; mức/range; priority; sequence; vị trí; trigger/điều kiện; decision rule; giới hạn; stop condition. Không ép mọi Result phải chứa tất cả thành phần này.
6. Tổng quát hóa observation thành chỉ dẫn có thể thao tác trên Content A mới. Chỉ kết luận những gì evidence trong dataset hỗ trợ.
7. Kiểm tra Result không trùng chức năng với Result của criterion khác và không suy diễn intention.
8. Nếu criterion không áp dụng, pattern chỉ là hiện tượng đơn lẻ không tổng quát hóa được, hoặc evidence không đủ để tạo rule hữu ích, trả `result: ""`.

Các nhóm criterion có phạm vi khác nhau. Ví dụ: coverage của A cần xét selection/omission; transformation cần xét cách material từ A được chuyển hóa; new-content criterion cần phân biệt bổ sung thật sự với paraphrase; sentence/syntax chỉ xét cấu trúc câu của narration; engagement chỉ xét mechanism quan sát được tạo engagement. Không kéo các biến ngoài phạm vi vào Result chỉ vì chúng xuất hiện trong dataset.

#### Tổng hợp cross-sample

Không biến hiện tượng chỉ xuất hiện ngẫu nhiên ở một sample thành rule chung. Với mỗi criterion:

- xác định biểu hiện trong từng cặp áp dụng được;
- tìm điểm lặp và phần ổn định;
- ghi nhận range hoặc phần co giãn;
- tìm đặc điểm trong A hoặc bối cảnh quan sát được làm pattern thay đổi;
- giữ trigger trong Result nếu rule chỉ đúng có điều kiện;
- coi ngoại lệ là có hệ thống chỉ khi dataset cho thấy điều kiện phân biệt hợp lý.

Khi chỉ có một cặp hoặc bằng chứng ít, chỉ giữ rule nếu quan hệ A → B trực tiếp, rõ và có thể thao tác; không gọi một lựa chọn ngẫu nhiên là stable pattern. Hạ mức khẳng định bằng điều kiện/giới hạn hoặc để Result rỗng.

#### Đơn vị đo

Với tỷ lệ về nội dung, ưu tiên semantic fact, proposition hoặc idea unit:

- tỷ lệ A được giữ dựa trên lượng ý/fact của A được giữ hoặc giữ ý nhưng chuyển hóa;
- tỷ lệ nội dung mới trong B dựa trên lượng ý/fact mới thật sự trong B;
- filler, lặp lời hoặc độ dài câu không được làm sai lệch kết luận về content retention.

Chủ yếu dùng word count hoặc sentence count cho độ dài, pacing, sentence rhythm, source run và narration run. Chỉ nêu số hoặc range khi cách đo nhất quán và evidence đủ mạnh; nếu không, mô tả pattern định tính có giới hạn rõ.

#### Chuẩn chất lượng Result

Result nên tiến gần logic sau trong phạm vi evidence cho phép:

`đặc điểm/trigger trong A → quyết định quan sát được → hành động chuyển đổi → mức độ/vị trí/thứ tự → điều kiện thay đổi → giới hạn hoặc điểm dừng`

Không suy diễn mục đích, cảm xúc mong muốn hoặc trạng thái tinh thần của tác giả. Mô tả behavior quan sát được từ A ↔ B. Ví dụ, ưu tiên “đưa rarity trước valuation và để valuation ở cuối” hơn “muốn người xem tò mò”.

Trước khi chấp nhận Result, tự kiểm tra: **Nếu một LLM khác nhận Result này cùng Content A mới, nó có biết phải quyết định hoặc làm gì không?** Những câu chung như “giữ fact quan trọng”, “mở rộng để hấp dẫn hơn”, “phần thân chi tiết” hoặc “giọng kể chuyên nghiệp” không đạt nếu thiếu tiêu chí chọn, hành động, vị trí, mức độ hoặc điều kiện có evidence.

Không kể lại dài dòng từng segment, không tổ chức Result theo kiểu `A nói..., B nói..., Part 1..., Part 2...`, và không giữ facts riêng của dataset làm rule. Viết trực tiếp thành chỉ dẫn chuyển đổi cô đọng. Không bịa evidence, tỷ lệ hoặc certainty.

#### Boundary và chống trùng giữa criteria

Mỗi Result phải trả lời đúng chức năng riêng của criterion. Cùng một evidence có thể hỗ trợ nhiều criterion, nhưng góc kết luận phải khác theo boundary; không copy hoặc paraphrase gần giống một Result sang nhiều criterion. Ví dụ: criterion về nguồn chọn material trả lời “lấy từ đâu”; criterion về macro-structure trả lời “B được tổ chức thành các phần nào”; criterion về hierarchy/reordering trả lời “fact nào được ưu tiên và chuyển vị trí ra sao”. Nếu criterion không bổ sung DNA riêng hữu ích, để Result rỗng.

#### Định dạng JSON

Tạo đúng một object cho mọi criterion, kể cả khi Result rỗng, và giữ nguyên thứ tự cùng định danh từ payload:

```json
{
  "schema_version": "1.0",
  "results": [
    {
      "criterion_key": "Temp!2:C001",
      "criterion_id": "C001",
      "result": "Quy luật có thể thao tác cho criterion này."
    }
  ]
}
```

Ghi document vào đường dẫn `analysis_results` trong config và tuân thủ [analysis_results.schema.json](schemas/analysis_results.schema.json). Không thêm field, không đổi thứ tự hoặc định danh criterion. Với criterion bị loại, dùng chính xác chuỗi rỗng; không ghi lý do, `Không áp dụng`, placeholder hoặc trạng thái `KEEP/DROP` vào `result`.

### 3. FINALIZE — Python ghi Excel

Chạy:

```powershell
python .agents/skills/analyze-content-pair/scripts/excel_bridge.py finalize --config .agents/skills/analyze-content-pair/config.yaml
```

Python kiểm tra document có đúng `schema_version` và `results`; đủ đúng một object cho mỗi criterion; không có field lạ, khóa lạ hoặc trùng; `criterion_key`, `criterion_id`, sheet và row vẫn khớp workbook; mọi `result` là chuỗi; và payload chưa lỗi thời.

Nếu validation thất bại, không tạo hoặc thay thế output. Nếu đạt, Python trim Result, loại criteria có Result rỗng và tạo workbook `output/result_analysis.xlsx` mới. Workbook chỉ chứa các cột `ID`, `Nhóm`, `Tiêu chí`, `Requirement`, `Result`, theo thứ tự tương đối trong `DEFAULT_CRITERIA.xlsx`.

## Điều kiện hoàn thành

- Đúng một result object cho mọi criterion; Result rỗng khi không có rule riêng, đáng tin cậy và tái sử dụng được.
- Mọi Result có nội dung trả lời cách tác giả triển khai đúng criterion khi biến A thành B và giúp ra quyết định trên A mới.
- Paraphrase/chuyển hóa không bị tính là nội dung mới; reordering chỉ là thuộc tính bổ sung.
- Kết luận cross-sample giữ stable pattern, conditional variation, trigger và range khi có evidence.
- Không dùng knowledge ngoài dataset, không suy diễn intention và không trùng Result giữa criteria.
- Output chỉ chứa criteria có Result khác rỗng, đúng thứ tự và đúng năm cột.
- Workbook criteria và hai Content nguồn không bị sửa.
- Không tạo `AUTHOR_SKILL` hoặc AUTHOR DNA cuối cùng.
