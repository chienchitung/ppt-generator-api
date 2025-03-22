# PPT Generator API

一個基於 FastAPI 的服務，用於自動生成競品分析 PowerPoint 簡報。

## 目錄
- [功能特點](#功能特點)
- [系統要求](#系統要求)
- [快速開始](#快速開始)
- [API 使用說明](#api-使用說明)
- [輸入格式](#輸入格式)
- [開發指南](#開發指南)
- [測試](#測試)
- [常見問題](#常見問題)
- [貢獻指南](#貢獻指南)

## 功能特點

- 自動生成結構化的競品分析 PowerPoint 簡報
- 支援繁體中文內容
- RESTful API 接口
- Docker 容器化部署
- 完整的測試套件
- 美觀的簡報模板
- 自動生成多個應用的分析頁面

## 系統要求

- Docker 和 Docker Compose
- Python 3.9+ (本地開發用)
- 支援繁體中文的系統環境

## 快速開始

1. 克隆倉庫：
```bash
git clone <repository-url>
cd ppt-generator-api
```

2. 使用 Docker 運行：
```bash
docker-compose up --build
```

服務將在 `http://localhost:8000` 啟動

## API 使用說明

### 1. 使用 Swagger UI（推薦方式）

1. 訪問 API 文檔：
   - 打開瀏覽器訪問 `http://localhost:8000/docs`
   
2. 生成 PPT：
   - 點擊 `POST /generate-ppt/` 端點
   - 點擊 "Try it out"
   - 上傳 JSON 檔案（可以使用 `sample_input.json`）
   - 點擊 "Execute"
   - 獲取生成的 PPT 檔案路徑，會得到類似這樣的回應：
     ```json
     {
       "message": "PPT generated successfully",
       "file_path": "generated_ppts/your_input_analysis.pptx"
     }
     ```

3. 下載 PPT：
   - 使用 `GET /download/{filename}` 端點
   - 重要：使用完整的檔案名（包含 `_analysis.pptx` 後綴）
   - 例如：如果你上傳的檔案叫 `test.json`，下載時應使用 `test_analysis.pptx`
   - 在 Swagger UI 中輸入完整檔案名並執行
   - 或直接在瀏覽器訪問 `http://localhost:8000/download/test_analysis.pptx`

### 2. 使用 curl 命令

```bash
# 生成 PPT
curl -X POST "http://localhost:8000/generate-ppt/" \
     -H "accept: application/json" \
     -H "Content-Type: multipart/form-data" \
     -F "input_file=@sample_input.json"

# 從回應中獲取檔案路徑
# 回應格式：{"message": "PPT generated successfully", "file_path": "generated_ppts/sample_input_analysis.pptx"}

# 下載 PPT（注意：使用完整的檔案名，包含 _analysis.pptx）
curl -O -J -L "http://localhost:8000/download/sample_input_analysis.pptx"
```

### 3. 使用 Python requests

```python
import requests

# 生成 PPT
files = {'input_file': open('sample_input.json', 'rb')}
response = requests.post('http://localhost:8000/generate-ppt/', files=files)
result = response.json()
file_path = result['file_path']

# 從 file_path 中獲取完整檔案名
filename = file_path.split('/')[-1]  # 會得到類似 'sample_input_analysis.pptx'

# 下載 PPT
response = requests.get(f'http://localhost:8000/download/{filename}')
if response.status_code == 200:
    with open(filename, 'wb') as f:
        f.write(response.content)
    print(f"檔案已下載: {filename}")
else:
    print(f"下載失敗: {response.status_code}")
```

### 檔案命名規則

生成的 PPT 檔案名遵循以下規則：
- 格式：`{輸入檔案名}_analysis.pptx`
- 例如：
  - 輸入：`test.json` → 輸出：`test_analysis.pptx`
  - 輸入：`sample_input.json` → 輸出：`sample_input_analysis.pptx`

## 輸入格式

### JSON 結構說明

```json
{
    "title": "電商App競品分析報告",
    "date": "2024-03-21",
    "apps": [
        {
            "name": "應用名稱",
            "ratings": {
                "ios": 4.5,
                "android": 4.3
            },
            "reviews": {
                "stats": {
                    "positive": 85,
                    "negative": 15
                },
                "analysis": {
                    "advantages": ["優點1", "優點2"],
                    "improvements": ["改進點1", "改進點2"],
                    "summary": "總結說明"
                }
            },
            "features": {
                "core": ["核心功能1", "核心功能2"],
                "advantages": ["優勢1", "優勢2"],
                "improvements": ["待改進1", "待改進2"]
            },
            "uxScores": {
                "memberlogin": 90,
                "search": 85,
                "product": 88,
                "checkout": 92,
                "service": 87,
                "other": 86
            },
            "uxAnalysis": {
                "strengths": ["優點1", "優點2"],
                "improvements": ["改進點1", "改進點2"],
                "summary": "用戶體驗總結"
            }
        }
    ]
}
```

### 欄位說明

- `title`: 簡報標題
- `date`: 報告日期
- `apps`: 應用列表（可包含多個應用）
  - `name`: 應用名稱
  - `ratings`: 評分資訊
  - `reviews`: 評論分析
  - `features`: 功能特點
  - `uxScores`: 用戶體驗評分
  - `uxAnalysis`: 用戶體驗分析

## 開發指南

### 本地開發環境設置

1. 創建虛擬環境：
```bash
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
```

2. 安裝依賴：
```bash
pip install -r requirements.txt
```

3. 運行開發服務器：
```bash
uvicorn app:app --reload
```

### 專案結構

```
ppt-generator-api/
├── app.py              # FastAPI 主應用
├── scripts/
│   └── generate_ppt.py # PPT 生成核心邏輯
├── requirements.txt    # Python 依賴
├── Dockerfile         # Docker 配置
├── docker-compose.yml # Docker Compose 配置
├── test_app.py        # 測試套件
├── sample_input.json  # 範例輸入檔案
├── generated_ppts/    # 生成的簡報存放目錄
└── README.md         # 文檔
```

## 測試

運行測試：
```bash
pytest test_app.py
```

測試覆蓋的功能：
- API 端點可用性
- JSON 驗證
- PPT 生成功能
- 檔案下載功能

## 常見問題

### 1. 中文顯示問題
確保系統和 Docker 容器都已正確設置中文語言環境：
```bash
# 檢查容器中的語言設置
docker exec -it <container-id> locale
```

### 2. 'Slides' object has no attribute 'add_shape' 錯誤
在使用 python-pptx 時，需要注意 `slide.shapes.add_shape()` 與 `shapes.add_shape()` 的區別：

- **錯誤寫法**：`slide.shapes.add_shape()`
- **正確寫法**：使用 `shapes = slide.shapes` 後再調用 `shapes.add_shape()`

**示例修正**：
```python
# 錯誤寫法
background = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)

# 正確寫法
shapes = slide.shapes
background = shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
```

### 3. WebP 格式圖片轉換問題
當處理 WebP 格式的圖片時，需要確保正確轉換為 PNG 或 JPEG 格式：

1. 首先安裝 Pillow 庫：
```bash
pip install Pillow
```

2. 使用以下代碼處理 WebP 圖片轉換：
```python
from PIL import Image
from io import BytesIO

def convert_webp_to_png(webp_data):
    try:
        image = Image.open(webp_data)
        # 處理透明通道
        if image.mode in ('RGBA', 'LA') or (image.mode == 'P' and 'transparency' in image.info):
            background = Image.new('RGBA', image.size, (255, 255, 255))
            if image.mode == 'RGBA':
                background.paste(image, mask=image.split()[3])
            else:
                background.paste(image)
            image = background.convert('RGB')
        else:
            image = image.convert('RGB')
        
        output = BytesIO()
        image.save(output, format="PNG")
        output.seek(0)
        return output
    except Exception as e:
        logger.error(f"Error converting WebP to PNG: {str(e)}")
        return None
```

### 4. PowerShell 和 Docker 命令執行問題
在 Windows PowerShell 中執行 Docker 命令時，需要注意以下幾點：

1. PowerShell 不支持 `&&` 連接命令，應使用分號 `;` 代替：
```powershell
# 錯誤
cd directory && docker-compose up

# 正確
cd directory; docker-compose up
```

2. Windows 路徑格式在 Docker 命令中的使用：
```powershell
# 使用雙反斜線或單正斜線
docker run -v C:/path/to/directory:/app image_name
```

### 5. 橫向與垂直置中布局設計
如需實現元素在幻燈片中橫向和垂直置中，可使用以下方法：

```python
# 橫向置中計算
element_width = Inches(2)
element_left = (prs.slide_width - element_width) / 2

# 垂直置中計算
element_height = Inches(2)
element_top = (prs.slide_height - element_height) / 2

# 多個元素垂直排列且整體置中
first_element_height = Inches(1.5) 
second_element_height = Inches(1.0)
gap = Inches(0.5)
total_height = first_element_height + gap + second_element_height

# 計算第一個元素的起始位置
start_y = (prs.slide_height - total_height) / 2
first_element_top = start_y
second_element_top = start_y + first_element_height + gap
```

## 貢獻指南

請參考 [CONTRIBUTING.md](CONTRIBUTING.md) 文件。