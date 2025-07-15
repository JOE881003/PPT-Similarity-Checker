# 📄 PPT Similarity Checker

一個用來比對一份「主 PPT」與其他多份 PPT 檔案間**文字相似度**的工具。  
基於 Python 實作，使用 `TF-IDF` 向量與 `餘弦相似度`（cosine similarity）做比對分析，適合用於議題關聯性比較、問題相似性篩選等情境。

---

## 🔧 功能簡介

- ✅ 提取 `.pptx` 檔案中的文字內容
- ✅ 指定一個主 PPT 與多個被比較的 PPT
- ✅ 使用 TF-IDF + Cosine Similarity 計算相似度
- ✅ 輸出每份 PPT 的相似度分數（0 ~ 1）

---

## 📦 安裝需求

請先安裝以下 Python 套件：

```bash
pip install python-pptx scikit-learn
```
## 專案結構建議
```bash
PPT_Similarity_Checker/
│
├─ main.py                 # 主程式
├─ main_issue.pptx         # 主比對的 PPT
├─ compare/                # 放入被比對的 PPT
│   ├─ compare1.pptx
│   ├─ compare2.pptx
│   └─ ...
└─ README.md
```
## 相似度計算說明
使用 TF-IDF 向量將文字轉換為數值形式，接著以餘弦相似度計算向量間相似程度：

similarity = (A · B) / (||A|| * ||B||)

相似度介於 0（完全不同）~ 1（完全相同）之間。
