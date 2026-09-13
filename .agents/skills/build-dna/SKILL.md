---
name: build-dna
description: Skill 2 — Build DNA. Tách và đánh giá semantic patterns từ Result xuyên Part trong từng Cxx, reconcile rồi gộp thành một kết quả DNA mỗi Cxx. Xuất AUTHOR_DNA.xlsx với sheet AUTHOR_DNA và bốn cột hiện có; không đọc lại raw A/B, research hoặc mapping workflow viết.
---

# Skill 2 — Build DNA

Flow: `GROUP BY Cxx → EXTRACT PATTERNS → COUNT/CONDITION → FILTER PATTERNS → NORMALIZE → RECONCILE → MERGE BACK TO ONE Cxx RESULT → EXPORT`.

**Cxx là nhóm tổng hợp; pattern là đơn vị đánh giá; một Cxx tối đa một dòng Excel.** Python gom dữ liệu, kiểm tra trích dẫn/phép tính/coverage và xuất workbook. LLM nhận diện semantic behavior, xác định eligibility, lọc, chuẩn hóa, reconcile và viết kết quả gộp. Python không tự kết luận behavior hoặc viết DNA.

## Đầu vào, review và contract Excel

Đọc [config.yaml](config.yaml); resolve đường dẫn từ project root. Đầu vào duy nhất là `result_analysis.xlsx` của Skill 1, gồm các sheet `Part1`, `Part2`, … với năm cột `ID | Nhóm | Tiêu chí | Requirement | Result`. Dùng mọi Part và criterion thực có; không cố định số Part hoặc danh sách ID.

- `grouped_payload`: JSON Result nguyên văn theo Cxx, có tọa độ nguồn.
- `dna_results`: JSON review **3.0** theo [schema](schemas/dna_results.schema.json), chứa mọi Cxx, candidates giữ/loại và kết quả gộp. Đây là dữ liệu nội bộ, không phải workbook kết quả bổ sung.
- `output`: duy nhất `AUTHOR_DNA.xlsx`; sheet kết quả cố định `AUTHOR_DNA`, đúng bốn cột:

`ID | Tiêu chí | DNA Rule | Condition/Variation`

Không đưa Rule ID, Level, Frequency, eligible/supporting Parts, confidence hoặc evidence thành cột Excel. Không tạo sheet phân tích trung gian. Các sheet khác có sẵn trong workbook được giữ nguyên.

Không đọc lại raw Content A/B hoặc DEFAULT_CRITERIA; không web, research, kiến thức ngoài hoặc xác minh facts. Chỉ Result của đúng Cxx làm evidence cho pattern thuộc Cxx đó. Result là observation Skill 1, không phải raw evidence đã được Skill 2 xác minh. Không sửa input. Khi thực thi chỉ dùng script có sẵn trong skill; không tạo helper mới.

## 1. GROUP BY Cxx

```powershell
python .agents/skills/build-dna/scripts/dna_bridge.py prepare --config .agents/skills/build-dna/config.yaml
```

Python đọc sheet Part theo số, ghép bằng Criterion ID và giữ thứ tự criterion ở Part đầu tiên. Sai schema, thiếu/trùng ID, metadata không khớp hoặc công thức thay dữ liệu: báo lỗi. Sheet ngoài Part không dùng làm evidence.

Giữ Result trống: đó là thiếu observation, không chứng minh thiếu material hoặc behavior vắng mặt. Không tự chuyển JSON review cũ hoặc đổi digest để bỏ qua việc tổng hợp lại.

## 2. EXTRACT PATTERNS — độc lập từng Cxx

Đọc Requirement và toàn bộ Result xuyên Part của một Cxx. Tách thành behaviors nhỏ có action riêng: lựa chọn material, nén, bỏ, trình bày, mở rộng, mức khẳng định… Ví dụ không phải danh sách bắt buộc.

Gộp các biểu hiện khác từ, fact hoặc cấu trúc bề mặt nếu cùng thực hiện một semantic behavior. Trước khi loại vì khác nhau giữa các Part, thử khái quát lên một cấp nhưng vẫn đúng Requirement và còn action cụ thể. Không khái quát thành “viết tốt” hoặc “chọn thông tin phù hợp”.

Tách các action trái nhau hoặc variation có trigger khác nhau. Mâu thuẫn một pattern không làm mất các pattern khác của Cxx. Không quyết định keep/drop cả criterion trước khi lọc từng pattern.

Mỗi criterion chứa `criterion_id`, `synthesis`, `patterns[]`, `dna_result`. Pattern có `pattern_id` nội bộ dạng `C430.P01`; ID giữ prefix Criterion ID và duy nhất trong JSON review, không đưa vào Excel. Khi sửa review đang có, giữ ID của pattern chưa đổi. Nếu không trích được candidate, dùng `patterns: []` và giải thích trong synthesis.

## 3. COUNT/CONDITION — từng pattern, xét mọi Part

Ghi `part_assessments[]` theo thứ tự corpus:

- `eligibility: eligible`: Result chứng minh có material/điều kiện để behavior có thể xuất hiện, kể cả khi B không thực hiện.
- `ineligible`: có evidence rõ không có material/điều kiện đó.
- `unknown`: chưa xác định được, kể cả Result trống. Không suy ineligible vì Result không nhắc material.
- `occurrence: present`: evidence cho thấy behavior thực sự được thực hiện.
- `absent`: eligible và có evidence rõ behavior không được thực hiện.
- `unknown`: chưa biết behavior có được thực hiện. Ineligible/unknown eligibility luôn có occurrence unknown.

Mỗi assessment có reason từ observation. `evidence[]` dùng trích dẫn nguyên văn từ Result đúng Cxx, có `part_id`, `result_excerpt`, `role`: `support | eligible | ineligible | counterexample | condition | unresolved`. Present cần support; absent cần counterexample; ineligible cần trích dẫn chứng minh thiếu material. Support có thể đồng thời chứng minh eligibility. Unknown không cần bịa quote.

Lưu các tập Parts theo thứ tự corpus:
- `eligible_parts`: tất cả eligible.
- `assessed_eligible_parts`: eligible có occurrence present/absent.
- `supporting_parts`: present.
- `unknown_occurrence_parts`: eligible có occurrence unknown.

Assessed và unknown occurrence phải tách biệt, hợp lại đúng eligible. Đếm mỗi Part một lần, không đếm quotes. Không chọn trigger bằng outcome B để tạo mẫu số chỉ gồm support.

```text
numerator = len(supporting_parts)
denominator = len(assessed_eligible_parts)
ratio = numerator / denominator
```

Unknown eligibility và eligible occurrence unknown không tham gia mẫu số; không tính unknown như absent. Nêu các giới hạn này trong limitations của pattern giữ. Denominator 0: ratio null, level UNASSESSED, drop vì chưa đo được.

| Tần suất trên assessed eligible Parts | Level nội bộ |
| --- | --- |
| 100% | CORE / Always trong phạm vi quan sát |
| >=80%, dưới 100% | DEFAULT |
| >=50%, dưới 80% | COMMON |
| >=30%, dưới 50% | POSSIBLE |
| <30% | POSSIBLE nếu có trigger rõ; nếu không INSTANCE_SPECIFIC và drop |

Frequency là tỷ lệ được observation xác nhận, không phải xác suất chắc chắn trên source mới. Với ba Part đánh giá được: 3/3 CORE, 2/3 COMMON, 1/3 POSSIBLE.

**Single observation:** support=1 và assessed=1 vẫn ghi 1/1 nhưng level thực thi POSSIBLE; chỉ giữ nếu có condition nhận diện được trên source mới và evidence cho condition. Nếu Part khác thiếu observation, ghi unknown, không tuyên bố chỉ một Part có material. Không auto-drop support đơn lẻ khi behavior tái sử dụng đủ căn cứ.

Condition phải nhận diện được trên source mới trước khi viết B; không dùng Part ID, tên vật, tên riêng, số liệu mẫu hoặc “khi B giữ…” làm trigger. Variation phải có condition được evidence của chính Cxx chứng minh. Không bịa condition để giải thích phản chứng.

## 4. FILTER PATTERNS

Xét giữ/loại từng pattern theo evidence, eligibility, action và tính tái sử dụng. Loại chi tiết instance, evidence yếu, khái quát rỗng, action không rõ, mâu thuẫn chưa giải quyết, hoặc pattern hiếm không có trigger.

Eligible nhưng không luôn thực hiện là counterexample để hạ level, không tự làm criterion bị loại. Hai action đối nghịch dưới cùng điều kiện phải tách/xử lý; không lấy đa số để thắng.

`confidence: High | Medium | Low` độc lập với level; 100% vẫn có thể evidence yếu. Low bị loại. Lưu `status: keep | drop` và filter_reason cho pattern drop; giữ cả candidates bị loại trong JSON. Không ép số patterns hoặc số Cxx xuất.

## 5. NORMALIZE

Pattern giữ bắt buộc có `action`: thao tác A→B cụ thể, áp dụng được trên source mới. `condition` là trigger riêng, dùng chuỗi rỗng khi không cần trigger. `selection`, `validation`, `limitations` chỉ ghi khi thực có căn cứ; không sinh filler.

Scope corpus ở `metadata.corpus_scope`, giới hạn chung ở `metadata.limitations`; không lặp scope đó trong từng WHEN. Không có target_step hay mapping vào các bước viết trong Skill 2; mapping thuộc Skill 3.

Loại facts mẫu khỏi instruction thực thi. Số liệu Result không tự trở thành threshold/quota; không ép schema của một loại đối tượng lên mọi source.

**Mọi expansion phải có điểm neo từ source; phần diễn giải mới được phép phát triển từ điểm neo đó.** Không viết “chỉ dùng nội dung có trong source”. Phân biệt paraphrase/tổng hợp với giải thích, suy luận, đối chiếu hoặc significance thật sự thêm. Fact mới trong B mẫu không chứng minh permission bịa facts hoặc research rule.

Thử instruction trên source giả định cùng condition và ngoài condition: có áp dụng được mà không cần facts mẫu không? Đây là kiểm tra tính tái sử dụng, không phải evidence mới.

## 6. RECONCILE — sau normalize

Rà patterns đã chuẩn hóa giữa các Cxx để phát hiện action trùng, overlap hoặc đối nghịch:

- Trùng: ghi `redundant_with` theo pattern_id; có thể giữ để truy vết hoặc drop pattern dư với lý do.
- Khác nhau có condition phân biệt được: ghi `conflicts_with`, condition có căn cứ riêng và `reconcile_note`.
- Đối nghịch cùng điều kiện chưa giải quyết: sửa/tách/drop; không xuất cả hai hoặc tự chọn theo tần suất cao hơn.

Không chuyển evidence hoặc frequency giữa Cxx. Criterion khác chỉ giúp nhận diện quan hệ, không là căn cứ tạo trigger cho pattern hiện tại. Sửa condition phải đánh giá lại eligibility/frequency. References phải tồn tại, không tự tham chiếu hoặc tạo vòng redundancy.

Chỉ ghi `metadata.reconciliation_completed: true` sau khi rà ngữ nghĩa thật sự. Script kiểm tra cấu trúc/liên kết, không chứng minh reconcile ngữ nghĩa.

## 7. MERGE BACK TO ONE Cxx RESULT

Sau reconcile, LLM gộp **toàn bộ patterns keep của từng Cxx** thành một `dna_result`:

- `pattern_ids`: đúng toàn bộ pattern giữ của Cxx, theo thứ tự; chỉ để truy vết trong JSON.
- `action`: bắt buộc; giữ nguyên trường chuỗi hiện tại, có thể chứa nhiều dòng với nhãn level của từng pattern giữ. Không nén tất cả thành một ACTION chung làm mất `CORE / DEFAULT / COMMON / POSSIBLE` hoặc condition riêng.
- `when`, `selection`, `validation`: optional, chỉ có nếu patterns giữ có căn cứ tương ứng; giữ rõ nội dung thuộc pattern nào. Chỉ dùng WHEN chung khi condition thực sự áp dụng cho toàn bộ patterns giữ.
- `condition_variation`: chỉ ghi các trigger/variation và giới hạn cần thiết để áp dụng đúng, gắn rõ với hành vi tương ứng; không chứa action mới hoặc filler. Trigger riêng phải được giữ tại đây hoặc WHEN phù hợp; riêng POSSIBLE phải nêu trigger ngay trong dòng action để không bị đọc thành mặc định.

**Bảo toàn level và condition khi merge:**

- `CORE`: ghi rõ là hành vi chính/bắt buộc trong phạm vi quan sát, vẫn giữ condition và giới hạn của pattern nếu có; không nâng thành bắt buộc ngoài phạm vi đó.
- `DEFAULT`: ghi là hành vi mặc định, trừ khi condition khác có căn cứ từ patterns giữ áp dụng; không tự tạo ngoại lệ.
- `COMMON`: ghi là hành vi thường dùng, không diễn đạt như bắt buộc.
- `POSSIBLE`: luôn giữ rõ trigger/condition đã có của pattern và diễn đạt là hành vi có thể dùng khi trigger đó thỏa; không viết như rule mặc định.

Hai pattern khác level hoặc khác trigger phải được diễn đạt thành các câu/dòng riêng, không gộp thành cùng một câu làm mất khác biệt. Chỉ gộp cách diễn đạt của patterns cùng level và cùng trigger khi vẫn giữ đủ từng behavior, selection, validation và giới hạn có căn cứ. Không đổi level hoặc condition đã xác định ở các bước trước để làm câu gộp ngắn hơn.

Trong cùng một cell `DNA Rule`, phần ACTION có thể trình bày như sau; chỉ dùng các level thực có trong patterns keep:

```text
CORE: Hành vi chính/bắt buộc trong phạm vi quan sát: …
DEFAULT: Mặc định …, trừ khi điều kiện khác đã nêu áp dụng.
COMMON: Thường dùng …
POSSIBLE: Khi [trigger có căn cứ], có thể …
```

Các nhãn này là nội dung của chuỗi `action`, không phải field JSON hoặc cột Excel mới. `AUTHOR_DNA.xlsx` vẫn đúng bốn cột `ID | Tiêu chí | DNA Rule | Condition/Variation`, tối đa một dòng mỗi Cxx.

Không thêm behavior mới khi merge, không đưa patterns drop vào action, không mất pattern giữ hoặc điều kiện có căn cứ. Không ghép evidence giữa Cxx. Một Cxx không còn pattern giữ thì `dna_result: null`, không xuất Excel. Synthesis giải thích kết quả sau lọc/reconcile.

Chỉ ghi `metadata.merge_completed: true` sau khi rà từng pattern keep với `dna_result`: behavior còn đầy đủ, level được ghi và diễn đạt đúng, trigger/condition còn nguyên và gắn đúng hành vi, selection/validation/giới hạn có căn cứ không bị mất. Không bỏ pattern keep chỉ để câu ngắn hơn; không có behavior mới hoặc pattern drop trong kết quả gộp. Python kiểm tra pattern_ids; LLM kiểm tra coverage ngữ nghĩa, từng level và condition. JSON review giữ nguyên schema_version 3.0, schema hiện tại và digest payload hiện tại.

## 8. EXPORT — replace sheet trong chính AUTHOR_DNA.xlsx

```powershell
python .agents/skills/build-dna/scripts/dna_bridge.py finalize --config .agents/skills/build-dna/config.yaml
```

Một dòng mỗi Cxx có dna_result, đúng bốn cột hiện có. DNA Rule chứa ACTION; chỉ thêm WHEN/SELECTION/VALIDATION có nội dung thực. Cột cuối chứa condition_variation hoặc để trống. Level, frequency, evidence, confidence và Part lists chỉ ở JSON review.

Nếu output đã tồn tại, mở workbook đó, xóa sheet `AUTHOR_DNA` cũ rồi tạo lại sheet cùng tên tại vị trí cũ. Giữ nguyên các sheet khác, dữ liệu, công thức và định dạng của chúng. Nếu chưa có workbook/sheet này, tạo sheet AUTHOR_DNA. Không thêm sheet review và không tạo tên output có hậu tố phiên bản. Tất cả patterns bị loại thì sheet AUTHOR_DNA chỉ có header; các sheet khác vẫn giữ.

Việc thay sheet là hành vi mặc định mỗi lần chạy, không cần --overwrite hoặc xác nhận lại. Có thể ghi file tạm nội bộ, mở lại kiểm tra bốn cột/coverage và các sheet được giữ, rồi thay chính AUTHOR_DNA.xlsx. Sai JSON, payload stale, lỗi ghi hoặc input/output đổi trong lúc export: báo lỗi và giữ file cũ; không xuất sang path khác để né lỗi.

## Điều kiện hoàn thành

- Mọi Cxx được đọc đầy đủ và lọc ở cấp pattern; assessments/frequency/evidence riêng.
- Mẫu số assessed eligible đúng; unknown không thành absent; single observation giữ ở POSSIBLE khi có trigger đủ căn cứ.
- Rules giữ có ACTION, confidence đủ căn cứ, không facts mẫu/filler/mapping workflow.
- Reconcile không dùng evidence chéo; merge bao phủ đúng patterns keep, giữ mức áp dụng/conditions.
- Excel vẫn sheet AUTHOR_DNA, đúng bốn cột, tối đa một dòng/Cxx; replace sheet trong chính AUTHOR_DNA.xlsx và giữ các sheet khác.
- Input bất biến; không tuyên bố xác minh raw A/B hoặc đã tạo AUTHOR_SKILL.
