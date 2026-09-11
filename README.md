# TK 印尼表格转化工具

EasyBoss 印尼站「导出 #SKU」xlsx → TikTok Shop Indonesia 批量上传模板（V5.0.2）的本地转化工具。
**最终产物**：单个 Windows `.exe`，双击即用。

## 适合谁

- 在 **TikTok Shop Indonesia**（Tokopedia 已合并）开店的卖家
- 用 **EasyBoss** 多店管理系统管理 SKU
- 需要把 EasyBoss 导出 xlsx 转成 TikTok 后台可批量上传的格式
- **每天出 1-几十个产品**，不想每次手动填写 41 列 + 9 个 property 字段

## 跟兄弟仓库的关系

| 仓库 | 目标市场 | 当前版本 |
|---|---|---|
| `yuqingyangxt1-wq/tk-ph-converter` | TikTok Shop Philippines | v3.2.4 |
| `yuqingyangxt1-wq/tk-th-converter` | TikTok Shop Thailand | v1.0.2 |
| **`yuqingyangxt1-wq/tk-id-converter`** ← **本仓库** | **TikTok Shop Indonesia** | **v1.0.0** |

代码主体共用；只在 `app/config.py` + `app/tiktok_writer.py` + `app/main.py` 改市场相关默认值。

## ID 版相对 PH/TH 的差异

| 项 | ID 版默认 |
|---|---|
| 货币 | Rp（IDR）`115000` |
| 标题前缀 | `Kaos Unisex Oversize `（印尼语中性） |
| 模板列数 | **41 列**（PH/TH 39 列，多了 `minimum_order_quantity` + `shipping_insurance`） |
| 类目格式 | 英文（PH/TH 也是英文） |
| 类目来源 | EasyBoss 源表给印尼语 → 工具自动翻译映射到英文 |
| 颜色 property_value_1 | 印尼语 → 英文（`Putih`→`White`） |
| shipping_insurance 值 | 字面量 `"Optional"`（PH/TH 没有这列，ID 后台特殊） |
| 预订单列 | 保留（ID 模板有这列，PH/TH 已删；用户不填留空） |

## 快速开始

1. 拉 **Actions 最新的 `TK-ID-Converter-windows` artifact**，解包到任意目录。
2. 双击 `TK-ID-Converter.exe` 运行。
3. **第一次运行**：在 EasyBoss 导出「导出 #SKU」xlsx → 在工具里选这个 xlsx → 选输出目录 → 点「开始转化」。
4. 输出：`output/<源名>_TKID_<时间戳>.xlsx`，直接拿到 TikTok Seller Center 上传。

### 进阶选项（GUI「转化设置」里）

- **每产品输出份数**：默认 2（防查重）；设 1 = 不复制
- **拆分文件**：勾上后每份输出独立 xlsx
- **下载产品图片**：✅ 默认勾上（8 线程并发）；不勾秒出表
- **价格 / 品牌 / 类目 / 描述 / 尺码图**：都可以单独覆盖

### 重置默认值

GUI 右上角「设置」页的修改会**自动保存**到同目录 `config.json`。
想恢复出厂默认值：关掉工具 → 删 `config.json` → 重启。

## 已知限制

- 模板只有 **25 个男装类目**（女装/童装需要换其他模板）
- **印尼语类目翻译**只覆盖了模板 Category sheet 的 25 个男装路径——女装/童装需自扩展
- **product_property/* 9 列 ID** 工具不自动填（按类目匹配 HiddenAttr 表的值，规则复杂，PH/TH 也没做；卖家手动填或后批处理）
- **图片 URL** 默认走 EasyBoss CDN（`https://p16-oec-sg.ibyteimg.com/...`），TikTok 后台通常能直拉但有时效性——建议卖家上传后到 TikTok Media Center 把图片换成本站 URL

## 打包（开发者）

```bash
# 触发云端打包（默认）
git push origin main    # GitHub Actions 自动打 windows-latest + python-3.11

# 本地 macOS 打包参考（不推荐给卖家用）
pip install -r requirements.txt
pyinstaller TK-ID-Converter.spec
```

依赖：`openpyxl`、`Pillow`、`tkinterdnd2`、`pyinstaller`（仅打包时）。

## License

内部工具，未授权不得二次分发。