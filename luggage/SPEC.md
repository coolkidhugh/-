"""
拍照存行李 + 拍照找位置（简化版）

## 你怎么用

1. **拍照存档**：拍行李 → 标卡联号 → 写位置 + 备注  
2. **拍照查找**：以后从 10 件里找 1 件，再拍一张 → 告诉你位置，并给出当初存档的原图  
3. **存档列表**：翻全部、改备注、标记取出  

就这些。不强制房号、不强制 QR。

## 技术

- Streamlit 三页 + SQLite（`data/luggage/luggage.db`）
- 照片本地存 `data/luggage/photos/`
- 查找用感知哈希（外形 + 颜色），无需 API Key

## 运行

```bash
pip install -r requirements.txt
streamlit run app.py
```

货架名改 `luggage/config.py` 里的 `STORAGE_ZONES`。
