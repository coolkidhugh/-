"""
行李寄存 QR + 拍照找行李 — 实现说明（云端落地版）

> 原路径 `老板IP打造/行李寄存QR系统/BRIEF_行李寄存系统_实现规格_给云端Cursor.md`
> 未随仓库同步，本文件按「拍照找行李」目标落地的可运行规格。

## 目标

酒店行李房：存件拍照建档 + QR 票号；取件可扫码 / 房号 / **拍照相似检索** 定位货架位置。

## 角色与流程

### 存入（行李员）
1. 拍摄行李正面照（必填）
2. 填姓名、房号、件数、颜色/类型、存放位置
3. 系统生成票号 `L月日XXXX`、QR（内容 `LUGGAGE:<票号>`）、感知哈希入库
4. 打印/下载 QR 贴行李或交客人

### 拍照找行李
1. 对着目标行李（或客人手机里的旧照）再拍一张
2. 用 phash + dhash + colorhash 与在库件比对，按相似度排序
3. 展示候选：照片、票号、房号、**货架位置**
4. 可一键标记取出

### 扫码 / 房号取件
- 扫 QR 或手输票号直达
- 按房号 / 姓名筛选在库列表后取出

## 技术选型

| 项 | 选择 | 说明 |
|---|---|---|
| UI | Streamlit multipage | 与本仓库 Excel 工具同栈，前台平板可开 |
| DB | SQLite `data/luggage/luggage.db` | 零运维，单店够用 |
| 图 | JPEG 本地存 `data/luggage/photos/` | 不上传第三方 |
| QR | `qrcode` | 票号载荷 |
| 找图 | `imagehash` | 无 GPU / 无 API Key，可离线 |

## 页面

- `app.py` — 首页（Excel 分析 + 入口说明）
- `pages/1_存入行李.py`
- `pages/2_拍照找行李.py`
- `pages/3_扫码取件.py`
- `pages/4_在库一览.py`

## 本地运行

```bash
export PATH="$HOME/.local/bin:$PATH"
pip install -r requirements.txt
streamlit run app.py --server.address 0.0.0.0 --server.port 8501
```

侧边栏进入「存入行李 / 拍照找行李 / …」。

## 存放区

默认货架标签见 `luggage/config.py` 的 `STORAGE_ZONES`，可按现场平面图改名。

## 后续可接（未做）

- 热敏小票打印
- 微信/短信把 QR 推给客人
- CLIP / 向量库提高跨角度召回
- 多门店与账号权限
