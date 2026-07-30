"""
编号 → 实体图（核心）

## 要求

发 / 查卡联号（如 `56469`、`0056469`）时，**必须带出现场实体照片**，不是标注卡。

## 用法

1. **按编号看图** — 输入编号，大图展示 `luggage_records/<编号>/photo.jpg`  
2. **拍照存档** — 存档时必须上传现场实体图；同时写入编号目录  
3. **补传实体图** — 已有编号缺图时，在「按编号看图」页下方补传  
4. **拍照查找** — 可选：再拍一张按外形找  

## 存储

- `luggage_records/<编号>/photo.jpg` — 实体图（查编号用这个）  
- `luggage_records/<编号>/record.json` — 位置 / 备注  
- SQLite `data/luggage/luggage.db` — 索引与哈希  

## 运行

```bash
pip install -r requirements.txt
streamlit run app.py
```
