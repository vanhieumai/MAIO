# AskAI — bản cập nhật tất cả nhà cung cấp (03/10/2026)

## Cài đặt
Bản này THAY THẾ bộ fixes/askai-upgrade trước đó. Mã đã được chuẩn bị trong thư mục này; chưa ghi vào workbook/add-in đang chạy.

1. Sao lưu workbook/add-in CPMS đang dùng.
2. Trong VBE, chọn đúng VBAProject của CPMS (không chọn workbook dữ liệu).
3. Trong form **frmAskAI**, thay toàn bộ code bằng **frmAskAI_FULL_CODE.txt**. Giữ nguyên UserForm và các control trong Designer.
4. Nếu đã cài **modAskAIJson** từ bộ trước: giữ nguyên, module không thay đổi. Nếu chưa cài: Insert > Module, đặt tên modAskAIJson, dán **modAskAIJson_CODE.txt**. Hoặc import modAskAIJson.bas; không dùng cả hai cách.
5. Debug > Compile VBAProject, lưu và mở lại AskAI.
6. Chọn key của nhà cung cấp cần dùng, bấm Refresh Models và chọn model.

## Nhà cung cấp
| Nhà cung cấp | Danh sách khởi đầu mới | Cách cập nhật |
|---|---|---|
| OpenAI (API dùng các model GPT) | gpt-6-luna, gpt-6.1-sol, gpt-6-astra; giữ gpt-4.1-mini/4.1 để tương thích | GET /v1/models, lọc các model chat phù hợp |
| Gemini | gemini-3.5-flash-lite, gemini-3.8-flash, gemini-3.1-pro-preview | ListModels, đọc nextPageToken đến hết; yêu cầu generateContent |
| NVIDIA NIM | Nemotron 3.5 Lightning, Nemotron 3 Super, DeepSeek V4 Flash/Pro, GPT-OSS 120B | Catalog https://integrate.api.nvidia.com/v1/models; lọc các họ chat đã hỗ trợ |
| Cerebras | gpt-oss-120b, qwen-3.8-27b | GET https://api.cerebras.ai/v1/models |
| Azure OpenAI | Deployment do bạn cấu hình | Nút Deployments... nhập endpoint và tên deployment |
| GitHub Models | Dịch vụ đã ngừng | Nút Service info giải thích; không gọi API cũ, giữ key và ghi chú |
| Groq / OpenRouter | Giữ nâng cấp của bộ trước | Refresh Models; OpenRouter có free only |

Danh sách khởi đầu được đối chiếu tài liệu hiện tại. Catalog và quyền truy cập của API key mới quyết định model nào gọi được. Các bộ lọc chat bảo thủ, không bảo đảm bao gồm mọi model mới hoặc model tùy chỉnh. Catalog lớn quá giới hạn ô SETTINGS vẫn dùng trong phiên hiện tại nhưng không lưu trọn bộ để mở lại.

OpenAI dùng Chat Completions cho hỏi đáp/ảnh, không có tool calling. GPT-6 ở chế độ hỏi đáp này không cần chuyển sang Responses. Không gửi temperature hoặc các tham số cũ không tương thích. Giữ model đã lưu nếu model còn trong danh sách; không tự gọi lại hoặc tự chuyển sang nhà cung cấp khác khi có lỗi. Gói ChatGPT không thay thế API key/quyền API.

## Azure: dùng tên deployment
Bấm **Deployments...**:
- Endpoint có thể là URL gốc resource, URL kết thúc /openai/v1/, hoặc URL đầy đủ chat/completions.
- Nhập các tên deployment có sẵn trong resource, phân cách bằng |.
- Ví dụ tên do bạn tự đặt: cpms-fast|cpms-quality. Đây chỉ là ví dụ, không phải model được tự tạo.
- Với URL cũ có /deployments/old-name/chat/completions?api-version=..., code thay old-name bằng deployment đang chọn.
- Không dùng URL /responses cho form này.
- Form không tạo/nâng cấp deployment trên Azure. Tên model gốc không chứng minh có deployment tương ứng.
- Chưa suy đoán khả năng xem ảnh của deployment Azure chỉ từ tên: Azure hiện dùng text-only trong bản này.

## GitHub Models
Theo GitHub Docs, GitHub Models đã ngừng hoàn toàn từ 30/07/2026, bao gồm catalog và inference API. Không thể sửa bằng cách thay model ID hoặc tiếp tục gọi endpoint cũ.
Bản này giữ key/lịch sử, hiện thông báo và chặn lời gọi tới API đã ngừng. Không tự chuyển token GitHub sang OpenAI hoặc nhà cung cấp khác. GitHub Copilot là dịch vụ riêng.

## Ảnh và tính tương thích
OpenRouter lấy khả năng ảnh từ catalog. OpenAI/Gemini dùng các họ model đã xác định; Groq giữ Qwen đã xác minh ở bộ trước. NVIDIA/Cerebras/Azure chưa có xác nhận khả năng ảnh trong catalog đang dùng nên được giữ text-only, tránh gửi ảnh vào model không hỗ trợ.
Giữ cơ chế timeout/cancel, đọc JSON và lưu lịch sử của bộ trước.

## Kiểm tra
Xem VALIDATION.txt cho kiểm tra đã chạy. Chưa Compile trong Excel, chưa thử UI hoặc API thật bằng key của bạn.
Sau khi dán: Compile; thử Refresh Models và câu hỏi tiếng Việt trên từng key; thử Cancel; thử chọn deployment Azure thứ hai; xác nhận GitHub hiện thông báo thay vì gọi endpoint cũ.

## Tài liệu đối chiếu
- OpenAI: https://developers.openai.com/api/docs/guides/latest-model (phần Migration quickstart; Chat Completions không dùng tools được hỗ trợ).
- Gemini model: https://ai.google.dev/gemini-api/docs/models
- Gemini catalog/phân trang: https://ai.google.dev/api/models
- NVIDIA chat: https://docs.api.nvidia.com/nim/reference/llm-apis
- NVIDIA catalog: https://integrate.api.nvidia.com/v1/models
- Cerebras model: https://inference-docs.cerebras.ai/models/overview
- Cerebras catalog: https://inference-docs.cerebras.ai/api-reference/models/list-models
- Azure endpoint/deployment: https://learn.microsoft.com/en-us/azure/foundry/foundry-models/concepts/endpoints
- GitHub ngừng dịch vụ: https://docs.github.com/en/github-models

Công cụ phân giải model OpenAI tự động không kết nối được sau hai lần thử; dùng nội dung trang tài liệu chính thức vừa truy xuất để chọn model. Không đọc key trong SETTINGS hoặc gọi API tính phí khi chuẩn bị mã.

