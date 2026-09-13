---
name: apply-dna
description: Skill 3 — Apply DNA. Đọc Content A_new và AUTHOR_DNA.xlsx, chọn các DNA rules phù hợp với material hiện có trong A_new, lập cấu trúc và viết Content B_new theo DNA đã học. Chỉ cho phép Derived expansion; không phân tích lại các Part cũ, xây lại DNA hoặc bổ sung external facts.
---

# Skill 3 — Apply DNA

Flow: `A_new + AUTHOR_DNA.xlsx → READ DNA → MATCH RULES → CLASSIFY EXPANSION → PLAN B → WRITE B → VALIDATE → OUTPUT B_new`

**Mục tiêu duy nhất:** biến một Content A mới thành Content B mới bằng AUTHOR_DNA đã được tạo bởi Skill 2.

## Input

Đầu vào gồm:

- `A_new`: nội dung source mới cần chuyển đổi.
- `AUTHOR_DNA.xlsx`: DNA đã được Skill 2 tổng hợp, sheet `AUTHOR_DNA` với các cột:

`ID | Tiêu chí | DNA Rule | Condition/Variation`

Không đọc lại `result_analysis.xlsx`, các Part A/B cũ hoặc `DEFAULT_CRITERIA`.

Mode cố định của skill này: không dùng web, research hoặc kiến thức ngoài để bổ sung fact. Research phải thuộc mode/bước riêng, không tự chuyển mode trong luồng Apply DNA.

Không được dùng kiến thức nhớ sẵn của model để thêm factual claim không có trong A_new. Nếu một expansion cần fact ngoài source, đánh dấu nhu cầu đó nhưng không tự bổ sung fact. Nhu cầu này chỉ ghi nội bộ, không đưa vào B_new khi người dùng chỉ yêu cầu B_new.

Đọc file văn bản theo UTF-8. Không sửa `A_new` hoặc workbook đầu vào. Nếu thiếu đầu vào hoặc không đọc được sheet/các cột yêu cầu, báo rõ phần thiếu; không dùng các Part cũ thay cho `A_new`.

## 1. READ DNA

Đọc đầy đủ tất cả DNA rules.

Nhận diện trong từng rule các mức:

- `CORE`
- `DEFAULT`
- `COMMON`
- `POSSIBLE`

và các condition/variation đi kèm.

Một dòng Cxx có thể chứa nhiều hành vi khác level hoặc trigger. Đọc từng hành vi cùng condition, selection, validation và giới hạn tương ứng trong `DNA Rule` và `Condition/Variation`; không gán một level chung cho cả dòng. Giữ `ID` nguồn để truy vết nội bộ.

Không tự thay đổi level, không tổng hợp lại DNA và không tạo rule mới. Nếu level hoặc condition cần thiết không rõ, yêu cầu làm rõ phần đó trước khi áp dụng; không tự gán mặc định.

## 2. MATCH RULES TO A_new

Đọc toàn bộ `A_new` trước khi áp dụng rule.

Với từng rule:

- `CORE`: áp dụng khi `A_new` nằm trong scope và có material cần thiết.
- `DEFAULT`: áp dụng mặc định khi condition phù hợp và không có lý do rõ từ DNA để dùng variation khác.
- `COMMON`: áp dụng khi material trong `A_new` hỗ trợ cách xử lý đó.
- `POSSIBLE`: chỉ trở thành candidate khi trigger/condition được nêu trong DNA thực sự xuất hiện trong `A_new`. Trigger thỏa không bắt buộc áp dụng; chỉ chọn khi có đủ căn cứ từ source và giúp `B_new` rõ hơn, mạch lạc hơn hoặc thực hiện tốt hơn cách xử lý đã nêu trong DNA.

Không cố áp rule khi `A_new` không có material tương ứng.

Không tạo trigger từ nội dung B đang viết. Khi chưa xác định được condition từ `A_new`, không coi condition đã thỏa.

### Ưu tiên khi rules xung đột

Trong các rules cùng phù hợp, dùng thứ tự `CORE > DEFAULT > COMMON > POSSIBLE`. Nếu một rule có condition cụ thể hơn và condition đó thực sự phù hợp với `A_new`, ưu tiên rule đó trong phạm vi condition; ngoài phạm vi đó dùng thứ tự level.

Không cố thỏa cả hai rule nếu chúng mâu thuẫn. Không tự tạo condition hoặc variation để giải quyết xung đột, không thay level của DNA. Mọi lựa chọn vẫn phải tuân thủ giới hạn factuality của mode hiện tại.

## 3. CLASSIFY EXPANSION

Phân biệt hai loại expansion:

- **Derived expansion**: giải thích, suy luận hoặc significance có thể rút trực tiếp từ `A_new`, không cần tiền đề factual ngoài source.
- **External factual expansion**: facts, lịch sử, bối cảnh, rarity, thông tin kỹ thuật hoặc factual claim khác không có trong `A_new` và không thể suy ra trực tiếp.

Mode hiện tại chỉ cho phép **Derived expansion** khi DNA phù hợp cho phép cách xử lý đó. Một điểm neo cùng chủ đề không đủ để hợp thức hóa external fact.

Với mỗi điểm neo và cách triển khai dự kiến, phân loại:

- `SOURCE_ONLY`: paraphrase, nén, tổng hợp hoặc tái cấu trúc material có trong `A_new`.
- `DERIVED`: suy luận/giải thích trực tiếp từ `A_new`; xác định rõ anchor và cách rút ra ý mới từ anchor đó.
- `NEEDS_EXTERNAL_FACT`: muốn triển khai nhưng cần thông tin ngoài `A_new`; đánh dấu nhu cầu nội bộ, không tự bổ sung fact.

`NEEDS_EXTERNAL_FACT` không được viết vào `B_new` hoặc đưa vào plan như nội dung sẽ viết. Nếu một đoạn có cả derived và external factual expansion, chỉ giữ phần derived có thể đứng độc lập với căn cứ từ source.

## 4. PLAN B

Trước khi viết prose, xác định:

- material nào trong `A_new` được giữ;
- material nào được nén, tổng hợp hoặc bỏ;
- thứ tự các cụm nội dung;
- phần mở và phần kết;
- điểm neo nào có thể phát triển thêm theo phân loại `DERIVED`;
- rule nào đang được áp dụng cho từng cụm.

Xác định explicit source anchor cho mọi expansion dự kiến: đoạn/câu hoặc cụm facts cụ thể trong `A_new`, cách suy ra ý mới và DNA rule cho phép cách xử lý đó. Expansion không xác định được anchor thì loại khỏi plan; expansion cần external fact cũng loại khỏi phần nội dung sẽ viết.

Plan chỉ dùng nội bộ để viết; không đưa vào output nếu người dùng chỉ yêu cầu `B_new`.

## 5. WRITE B

Viết `B_new` theo plan và DNA.

Các nguyên tắc:

- ưu tiên CORE trước, xét ngoại lệ condition cụ thể theo cách xử lý xung đột ở bước MATCH RULES;
- dùng DEFAULT khi phù hợp;
- COMMON chỉ dùng khi material/condition hỗ trợ; POSSIBLE còn phải vượt qua bước chọn candidate, không tự áp dụng chỉ vì trigger thỏa;
- không ép tất cả rule vào cùng một nội dung;
- không bắt `A_new` phải có cùng cấu trúc với các Part mẫu;
- giữ facts của `A_new` chính xác theo source;
- được paraphrase, nén, sắp xếp lại và narration hóa theo DNA;
- chỉ viết Derived expansion đã được phân loại `DERIVED`, có explicit source anchor từ `A_new` và được DNA phù hợp cho phép;
- không bịa fact mới;
- không dùng fact từ các Part cũ;
- không sao chép nội dung riêng của đối tượng mẫu.

Derived expansion được DNA cho phép có thể phát triển diễn giải, suy luận, đối chiếu hoặc significance trực tiếp từ điểm neo; đối chiếu chỉ dùng material trong `A_new`. Không dùng kiến thức nhớ sẵn để thêm factual claim, kể cả khi DNA mô tả kiểu mở rộng cần external fact.

Không được biến đánh giá thành fact, suy luận thành lịch sử thực tế, khả năng thành khẳng định, hoặc correlation thành causation.

Không được ghép các lượt nói không liền nhau trong `A_new` thành một quotation duy nhất trong `B_new`.

Nếu nhiều câu của cùng người nói xuất hiện ở các vị trí khác nhau trong source:

- hoặc giữ thành các quote riêng đúng theo từng lượt gốc;
- hoặc tổng hợp/narration hóa mà không đặt trong dấu ngoặc kép.

Quote trong `B_new` phải phản ánh một đoạn lời nói liên tục có thật trong `A_new`. Không được tạo “composite quote” bằng cách nối các câu rời nhau rồi trình bày như nguyên văn.

Khi giữ lời thoại/reaction trực tiếp, phải giữ đúng quan hệ vị trí với phát biểu mà lời thoại/reaction đó phản hồi trong `A_new`. Áp dụng cả với reaction bằng lời và biểu hiện như cười, ngạc nhiên, dù có hoặc không có dấu ngoặc kép.

Không được chuyển một reaction từ vị trí khác trong `A_new` sang sau một phát biểu khác trong `B_new` rồi trình bày như reaction trực tiếp của phát biểu mới, kể cả để tạo kết thúc đẹp. Nếu sắp xếp lại một cụm hội thoại, phải giữ reaction cùng phát biểu gốc mà nó phản hồi và không để cách trình bày gán reaction sang câu khác.

Nếu muốn thay đổi vị trí riêng của reaction, phải narration hóa, giữ đúng ngữ cảnh phản hồi gốc và không trình bày như reaction trực tiếp của câu mới. Nếu bỏ phát biểu gốc, không giữ reaction trực tiếp như thể nó phản hồi phát biểu thay thế.

Khi sau valuation có một reaction trực tiếp của người nói và reaction đó tiếp tục thành một đoạn kết ngắn, liên tục, có chức năng khép hội thoại, không được tự động rút còn một từ/câu đầu. Giữ toàn bộ cụm reaction liên tục nếu nó ngắn và không mở sang chủ đề mới.

DNA quy định **cách xử lý**, không phải nội dung cụ thể cần xuất hiện.

## 6. VALIDATE

Sau khi viết `B_new`, kiểm tra:

- CORE rule phù hợp có bị bỏ sót không, hoặc có được thay bằng rule condition cụ thể phù hợp theo quy tắc xử lý xung đột;
- DEFAULT/COMMON/POSSIBLE có bị áp sai condition không;
- POSSIBLE có được chọn vì thực sự giúp `B_new` và có đủ căn cứ, hay chỉ vì trigger xuất hiện;
- rules xung đột có được xử lý theo level và condition cụ thể phù hợp, thay vì cố thỏa cả hai;
- có material quan trọng trong `A_new` bị mất ngoài DNA không;
- có factual claim nào không có căn cứ trong `A_new`, kể cả claim lấy từ kiến thức nhớ sẵn của model;
- có expansion nào không có điểm neo, không rút trực tiếp từ source, không được DNA cho phép hoặc thuộc `NEEDS_EXTERNAL_FACT`;
- có đánh giá bị biến thành fact, suy luận thành lịch sử thực tế, khả năng thành khẳng định, hoặc correlation thành causation;
- có câu nào bắt chước fact/tên/số liệu của Part mẫu không;
- với mọi đoạn đặt trong dấu ngoặc kép, đối chiếu `A_new` để xác nhận đó là lời thoại liên tục có thật;
- nếu quote được ghép từ nhiều lượt không liền nhau, phải sửa thành narration hoặc tách quote;
- với mọi lời thoại/reaction trực tiếp, xác định phát biểu mà nó phản hồi trong `A_new` và kiểm tra vị trí trong `B_new` có giữ đúng quan hệ đó không, kể cả ở phần kết và khi reaction không có dấu ngoặc kép;
- nếu reaction bị chuyển sang sau một phát biểu khác, phải đưa về cùng phát biểu gốc, narration hóa với đúng ngữ cảnh phản hồi gốc hoặc bỏ; không giữ như reaction trực tiếp của câu mới;
- Kiểm tra đoạn kết sau valuation: nếu source có một cụm reaction liên tục gồm nhiều câu ngắn cùng chức năng kết, không được chỉ giữ câu đầu rồi bỏ phần còn lại nếu việc bỏ làm mất sắc thái hoặc chức năng kết;
- cấu trúc, ngôi kể, mức khẳng định, xử lý hội thoại và expansion có phù hợp DNA không.

Với từng câu/ý không phải paraphrase trực tiếp từ `A_new`, xác định nó là `DERIVED` từ anchor nào. Với nén/tổng hợp/tái cấu trúc thuần `SOURCE_ONLY`, chỉ rõ material nguồn tương ứng; không ép gán nhãn `DERIVED` nếu không thêm diễn giải. Nếu không chỉ ra được anchor hoặc cần external fact, xóa hoặc viết lại. Anchor phải đủ để rút ra ý đó mà không cần tiền đề factual ngoài source.

Kiểm tra dựa trên `A_new` và DNA đang có; không mở lại các Part mẫu để đối chiếu. Chi tiết mẫu nằm trong DNA, nếu có, không trở thành material để viết.

Nếu có lỗi, sửa `B_new` trước khi trả kết quả.

## 7. OUTPUT

Mặc định chỉ trả nội dung `B_new`.

Không trả plan, rule matching, expansion classification, nhu cầu external fact, validation report hoặc giải thích nội bộ trừ khi người dùng yêu cầu.

Không tạo AUTHOR_DNA mới và không sửa `AUTHOR_DNA.xlsx`.
