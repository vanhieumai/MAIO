# Nâng cấp frmAskAI — Groq / OpenRouter
Ngày: 03/10/2026

## Cài đặt
Bộ này là mã thay thế, chưa được ghi vào workbook/add-in đang chạy.

1. Sao lưu add-in/workbook đang dùng. Trong VBE, chọn đúng VBAProject của CPMS (ảnh trước của bạn là CPMS AR.xlam), không chọn workbook dữ liệu.
2. Insert > Module, đặt tên **modAskAIJson**. Dán toàn bộ **modAskAIJson_CODE.txt** vào module mới.
   - Hoặc File > Import File > chọn **modAskAIJson.bas**. Chỉ dùng MỘT cách; không tạo hai module trùng.
3. Mở code của **frmAskAI**, Ctrl+A thay toàn bộ bằng **frmAskAI_FULL_CODE.txt**.
   - Giữ nguyên form và các control trong Designer. Không xóa/import lại UserForm.
   - Hai file *_CODE.txt không có các dòng Attribute/VERSION, phù hợp để dán trực tiếp.
4. Debug > Compile VBAProject, lưu add-in/workbook. Đóng rồi mở lại form AskAI.
5. Chọn API key Groq/OpenRouter, bấm **Refresh Models**, chọn model rồi gửi câu hỏi.

Không cần tạo thêm control bằng tay. Ba control mới được tạo khi mở form, có font theo cboModel:
- Refresh Models.
- OpenRouter: free only (mặc định bật, chỉ xuất hiện khi dùng OpenRouter).
- Cancel request (chỉ hiện khi đang chờ mạng).

## Thay đổi
- Groq: danh sách dự phòng cập nhật gồm openai/gpt-oss-20b, openai/gpt-oss-120b, qwen/qwen3.8-27b, cùng hai model Llama còn được tài liệu liệt kê. Qwen là preview; quyền truy cập thực tế phụ thuộc tài khoản.
- Refresh Models đọc catalog hiện tại từ API. Groq loại các dòng audio/guard và chỉ nhận các họ chat được hỗ trợ trong bộ lọc. OpenRouter kiểm tra input/output text và dùng metadata image để bật đính kèm.
- OpenRouter mặc định dùng openrouter/free. Refresh ở chế độ free only lọc theo giá prompt/completion/request bằng 0. Bỏ chọn free only rồi Refresh Models để xem cả model trả phí.
- Nhớ catalog theo provider/chế độ trong SETTINGS. Mở lại form không tự gọi mạng. Model đã chọn được giữ nếu còn trong danh sách; không tự gửi lại yêu cầu và không tự chuyển provider khi có lỗi.
- Hủy được yêu cầu đang chờ; đóng cửa sổ khi đang chờ sẽ hủy trước rồi đóng. Timeout: 30 giây lấy catalog, 120 giây trả lời.
- Sửa xử lý JSON có khoảng trắng, Unicode, ký tự escape, content dạng mảng, lỗi API cấp cao nhất; không coi từ “error” trong câu trả lời là lỗi API.
- Sửa lời gọi btnAskAI_Click không có thủ tục tương ứng; hỗ trợ cả tên nút btnAsk và btnAskAI.
- Sửa kiểm tra chỉ số ngoài mảng khi chọn/xóa API key. Không hiển thị toàn bộ key ngắn.
- PNG/WebP vẫn đính kèm được nếu LoadPicture của Excel không xem trước được. Giới hạn ảnh của ứng dụng: 10 MB.
- Lịch sử hỏi đáp, lưu ghi chú và cơ chế lưu key hiện có được giữ lại.

## Phạm vi
Ưu tiên Groq/OpenRouter theo lựa chọn của bạn. Danh sách model và endpoint của các nhà cung cấp khác được giữ từ code cũ; chưa xác minh toàn bộ nhà cung cấp khác. Bộ HTTP/JSON dùng chung được cải thiện. Không đổi sang GPT-6 hay bắt buộc dùng OpenAI.

Catalog cho biết model đang được nhà cung cấp liệt kê, không bảo đảm key của bạn có quyền gọi tất cả model. HTTP 401/403/404/429 sẽ có thông báo cụ thể. Cancel chỉ dừng việc chờ tại máy; không bảo đảm nhà cung cấp ngừng tính phí một yêu cầu đã nhận.

## Kiểm tra sau khi dán
- Compile không báo lỗi. Mở AskAI từ ribbon.
- Groq: Refresh Models, hỏi một câu tiếng Việt.
- OpenRouter: Refresh khi free only bật; chọn model; hỏi câu có dấu ngoặc kép.
- Bỏ free only rồi Refresh: có thể thấy model trả phí; chọn đúng model mong muốn.
- Hủy yêu cầu; gửi lại. Đóng bằng X lúc đang chờ rồi mở lại.
- Chọn model có ảnh, thử PNG; xác nhận ảnh vẫn có tên nếu không hiện preview.
- Đổi key, xem lịch sử và lưu một ghi chú như trước.

## Nguồn kiểm chứng
- Groq models + GET /models: https://console.groq.com/docs/models
- Groq Qwen text/image: https://console.groq.com/docs/model/qwen/qwen3.8-27b
- OpenRouter model metadata: https://openrouter.ai/docs/api/api-reference/models/list-all-models-and-their-properties
- OpenRouter free router: https://openrouter.ai/docs/guides/routing/routers/free-router

Mã gốc trích từ D:\OneDrive\MyProject\CPMS\CPMS AR Plus.xlsm, SHA256:
7c90ab77e274bec4d166918a3d614a175c859eb05b22666f740695b97531ac63

Không đọc API key từ SETTINGS, không gửi câu hỏi thật tới API trong quá trình chuẩn bị bản sửa.


## Kết quả kiểm tra đã thực hiện
Đã qua kiểm tra tĩnh và 30 trường hợp kiểm tra hành vi của các helper được chuyển cú pháp sang VBScript. Xem VALIDATION.txt. Chưa Compile hoặc chạy trực tiếp trong Excel, chưa gọi API thật bằng key của bạn.
