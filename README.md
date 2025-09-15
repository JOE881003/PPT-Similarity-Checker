# 📄 PPT Similarity Checker

一個基於 **Python + Streamlit** 的簡報比對工具，能夠：  
- 從 PPTX 檔案中擷取文字  
- 使用 **TF-IDF + Cosine Similarity** 計算簡報相似度  
- 串接 **OpenAI GPT API** 自動產生摘要，快速掌握重點  
---

## 🚀 功能特色
- 🔍 **PPT 文字比對**：輸入兩份簡報，即可計算相似度  
- 📝 **AI 摘要**：利用 GPT 生成條列式重點與關鍵字  
- 📊 **差異比較**：比較兩份簡報的異同，並提供合併建議  
- 🌐 **Web UI**：使用 Streamlit，介面簡單直覺，免安裝額外軟體  

---

## 📦 安裝需求

請先安裝以下 Python 套件：

```bash
streamlit
openai >= 1.0.0
python-pptx
scikit-learn
（已在 requirements.txt 列出）
```
## 專案結構建議
```bash
PPT_Similarity_Checker/
├── app_gpt.py # 主程式（Streamlit App）
├── requirements.txt # 相依套件清單
├── .streamlit/
  └── secrets.toml # 本地測試用的 API Key
```
## 相似度計算說明
使用 TF-IDF 向量將文字轉換為數值形式，接著以餘弦相似度計算向量間相似程度：

similarity = (A · B) / (||A|| * ||B||)

相似度介於 0（完全不同）~ 1（完全相同）之間。

## 使用範例

上傳兩份 PPTX 檔案

系統會顯示相似度分數

點選「產生 GPT 摘要」即可獲得重點與差異比較

