---
name: build-dna
description: Skill 2 — Build DNA. Tổng hợp độc lập từng Criterion ID xuyên các sheet Part của result_analysis.xlsx, lọc và chuẩn hóa pattern A→B có bằng chứng thành AUTHOR_DNA.xlsx. Dùng sau analyze-content-pair; không phân tích lại raw A/B hoặc tạo AUTHOR_SKILL.md.
---

# Skill 2 — Build DNA

Flow: `result_analysis.xlsx → GROUP BY Cxx → SYNTHESIZE → FILTER → NORMALIZE → EXPORT DNA`.

**Một đơn vị tổng hợp = một Criterion ID xuyên toàn bộ Part.** Python chỉ đọc, gom, kiểm tra cấu trúc và xuất dữ liệu. LLM so sánh Result, quyết định giữ/loại và viết DNA. Không gắn cứng danh sách C100…C460 hoặc số lượng 22.

## Đầu vào và artifact

Đọc [config.yaml](config.yaml), resolve đường dẫn từ project root. Đầu vào duy nhất là workbook Skill 1 đã xuất với các sheet `Part1`, `Part2`, … và năm cột `ID | Nhóm | Tiêu chí | Requirement | Result`.

- `grouped_payload`: Result gom theo ID, nguyên văn, kèm tọa độ nguồn.
- `dna_results`: một JSON do LLM viết, có quyết định của mọi ID, cả giữ và loại, theo [schemas/dna_results.schema.json](schemas/dna_results.schema.json).
- `output`: `AUTHOR_DNA.xlsx`, chỉ chứa ID giữ.

Không đọc lại Content A/B hoặc DEFAULT_CRITERIA, không dùng web hay kiến thức ngoài để bổ sung evidence. Result là observation do Skill 1 báo cáo, không phải raw evidence đã được Skill 2 xác minh. Nếu thiếu căn cứ, loại hoặc báo cần review. Chỉ sử dụng script có sẵn trong skill khi thực thi; không tạo helper mới. Không sửa workbook đầu vào.

## 1. PREPARE — GROUP BY Cxx

Chạy từ project root:

```powershell
python .agents/skills/build-dna/scripts/dna_bridge.py prepare --config .agents/skills/build-dna/config.yaml
```

Python đọc tất cả sheet Part theo số Part; gom cùng ID theo thứ tự criterion ở Part đầu tiên: `C100 = [Part1.C100, Part2.C100, ...]`. Giữ cả Result trống: trống là không có observation, **không** phải evidence hành vi vắng mặt. Không paraphrase, tính điểm hay tự chọn DNA.

Thiếu/trùng ID, tập criterion không khớp, metadata khác nhau, sai schema hoặc có công thức thay dữ liệu: báo lỗi. Cho phép thứ tự dòng khác nhau và ghép bằng ID. Sheet ngoài Part được liệt kê trong payload, không dùng làm evidence.

## 2. SYNTHESIZE EACH Cxx

Đọc đầy đủ payload. Hoàn thành một criterion trước khi chuyển sang criterion tiếp theo:

```text
FOR EACH criterion trong payload.criteria:
    Đọc Requirement và toàn bộ observations của ID hiện tại
    So sánh cách xử lý A→B, không chỉ từ khóa hoặc facts
    Tìm pattern lặp lại, variation có điều kiện, chi tiết riêng, mâu thuẫn
    FILTER → NORMALIZE criterion hiện tại
    Ghi một decision vào dna_results
END
```

Chỉ Result của ID hiện tại làm evidence cho quyết định đó. Không suy ra Cxx từ criterion khác hoặc viết DNA chung rồi nhân cho nhiều ID.

- **Lặp lại:** cùng cách chọn, bỏ, nén, bổ sung, sắp xếp hoặc trình bày thông tin trên nhiều đối tượng; facts có thể khác.
- **Variation:** cách xử lý đổi theo đặc điểm source/loại đối tượng/dimension; điều kiện phải được Result chứng minh và nhận diện được trên source mới. Không dùng Part ID, tên riêng hay số đo mẫu làm điều kiện.
- **Riêng một đối tượng:** fact hoặc cách xử lý chỉ thấy ở một Part, không có căn cứ xuyên đối tượng.
- **Mâu thuẫn:** khác cách xử lý; chỉ giải thích bằng variation khi có evidence về điều kiện, không suy diễn ý định tác giả.

Lưu `synthesis` cô đọng về pattern, phạm vi, variation và phản chứng. Mỗi evidence gồm `part_id`, `result_excerpt` trích nguyên văn ngắn từ Result và `role`: `support`, `variation`, `exception`, `unresolved`. `exception` phải có giới hạn phạm vi được giải thích; `unresolved` là mâu thuẫn chưa xác định điều kiện. Đọc cả Result không được trích, tránh chỉ chọn phần ủng hộ.

## 3. FILTER

Loại criterion chỉ chứa chi tiết instance, chỉ có một Part đủ evidence, evidence yếu/thiếu, mâu thuẫn chưa giải thích điều kiện, hoặc không tạo được rule A→B tái sử dụng và kiểm chứng được. Không ép đủ 22 Cxx hoặc một số lượng tối thiểu.

Rule giữ phải có ít nhất hai Part khác nhau chứng minh pattern hoặc variation có điều kiện. Đây là căn cứ tối thiểu cho hành vi xuyên đối tượng, không phải điểm số chất lượng: hai Result yếu vẫn không đủ. Nhiều trích dẫn cùng Part không thành nhiều đối tượng. Mỗi nhánh variation phải có evidence riêng; nhánh chỉ thấy một Part không được gọi là hành vi lặp lại.

`confidence`: `High` khi evidence trực tiếp, rõ giữa các Part; `Medium` khi có căn cứ hợp lý nhưng còn giới hạn cần ghi; `Low` khi phụ thuộc giả định. Loại `Low`. Không bỏ qua phản chứng vì đa số ủng hộ, không lấy retention ratio trung bình làm DNA.

Ghi `decision: "drop"`, `filter_reason` cụ thể; các trường thực thi để chuỗi rỗng. ID loại ở lại JSON để review, không vào workbook.

## 4. NORMALIZE DNA

Với `decision: "keep"`, viết hướng dẫn độc lập cho source mới:

- `when`: điều kiện, phạm vi; chỉ giới hạn loại đối tượng khi có căn cứ.
- `action`: thao tác chuyển A→B.
- `selection`: thông tin cần ưu tiên/giới hạn/loại, không chỉ tỷ lệ giữ.
- `validation`: dấu hiệu kiểm tra được trên source mới và output.
- `condition_variation`: nhánh điều kiện và giới hạn có evidence; ghi rõ nếu chưa thấy variation, không bịa nhánh.
- `target_step`: `STEP 1`, `STEP 2`, `STEP 3` hoặc `STEP 4` theo workflow AUTHOR_SKILL của project.

Không giữ “B giữ kích thước 12 inch.” Chỉ khi nhiều Result chứng minh cách chọn thông số, mới viết kiểu: “Khi A có nhiều thông số cùng loại, ưu tiên giữ thông số trực tiếp mô tả đối tượng chính; thông số so sánh phụ có thể được nén hoặc loại bỏ.” Thêm SELECTION và VALIDATION phù hợp; ví dụ không phải rule mặc định.

Không đưa tên riêng, số liệu hay sự kiện chỉ đúng với mẫu vào trường thực thi. Số liệu Result không tự thành threshold. Không ép schema đồ vật lên mọi source. B bổ sung ngoài A được tổng quát theo loại thông tin, không lưu fact mẫu; việc bổ sung chưa chứng minh tác giả có nghiên cứu bên ngoài.

Kiểm tra từng rule bằng tình huống source mới cùng phạm vi và tình huống ngoài điều kiện: có thể chọn áp dụng, xử lý thông tin và kiểm tra output mà không cần Part mẫu không? Đây là kiểm tra khả năng thực thi, không phải evidence mới.

Ghi JSON chung theo schema, sao `schema_version`, `payload_digest` từ payload; `criteria` có đúng một decision mỗi ID theo thứ tự payload. Giữ ID nguồn, không tái đánh số. `filter_reason` rỗng cho ID giữ. Python kiểm tra cấu trúc, trích dẫn và liên kết nguồn; LLM kiểm tra ngữ nghĩa, variation và loại facts mẫu.

## 5. EXPORT DNA

```powershell
python .agents/skills/build-dna/scripts/dna_bridge.py finalize --config .agents/skills/build-dna/config.yaml
```

Một sheet `AUTHOR_DNA`, một dòng mỗi ID giữ, bốn cột:

`ID | Tiêu chí | DNA Rule | Condition/Variation`

`DNA Rule` chứa `WHEN`, `ACTION`, `SELECTION`, `VALIDATION`. Cột cuối chứa điều kiện/variation. Confidence, target step, synthesis và evidence nằm trong JSON review. Nếu tất cả bị loại, workbook chỉ có header; báo rõ không có DNA đủ căn cứ.

Output tồn tại: chọn đường dẫn mới; chỉ thêm `--overwrite` khi thay thế có chủ đích đã thuộc yêu cầu người dùng. Script tạo file tạm, mở lại kiểm tra rồi mới thay output; lỗi giữ file cũ. Input đổi sau prepare phải prepare và tổng hợp lại, không chỉ đổi digest cho khớp.

## Điều kiện hoàn thành

- Mọi ID có decision sau khi tổng hợp độc lập đầy đủ Result xuyên Part.
- ID giữ có evidence xuyên đối tượng, không có mâu thuẫn chưa giải thích, confidence High/Medium, đủ bốn trường thực thi và không chứa facts mẫu.
- Trích dẫn tồn tại đúng ID/Part; LLM kiểm tra điều kiện variation và phạm vi.
- JSON hợp lệ; workbook mở lại đúng bốn cột, chỉ xuất ID giữ, không chèn placeholder cho ID loại.
- Input giữ nguyên; không tuyên bố đã tạo AUTHOR_SKILL hoặc xác minh raw A/B.
