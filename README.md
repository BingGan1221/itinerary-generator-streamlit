# 自动生成行程单

## 在线部署

这个项目可以作为 Streamlit 应用部署。部署入口文件是：

```text
streamlit_app.py
```

部署到 Streamlit Community Cloud 时，把仓库连接到 Streamlit，选择对应分支和上面的入口文件即可。

本地预览网页版本：

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/streamlit run streamlit_app.py
```

注意：不要把真实订单 Excel、生成后的行程单、`.venv/` 或本地 `config/trip_config.json` 上传到 GitHub。仓库里只保留 `config/trip_config.example.json` 作为示例配置。

## 文件夹用途

- `input/`：放订单 Excel，文件名保持 `订单列表_*.xlsx`。
- `templates/`：放 Word 模板，默认使用 `行程单通用模版.docx`。
- `config/`：放本地行程配置，默认使用 `trip_config.json`；可从 `trip_config.example.json` 复制。
- `output/`：生成后的行程单会放在这里。
- `examples/`：参考 PDF、旧模板等样例文件。

## 使用方法

1. 把新的订单 Excel 放进 `input/`，也可以直接放项目根目录。
2. 在项目根目录运行：

```bash
./run.sh
```

3. 按终端提示填写旅游路线、领队、领队电话。直接按回车会使用 `config/trip_config.json` 里的默认值。

脚本会自动查找最新的 `订单列表_*.xlsx`，生成 DOCX 到 `output/`。
如果订单表里有 `路线` 和 `出行时间`，文件名会自动使用它们，例如：

```text
060523青甘大环线行程单.docx
```

开始时间和结束时间也会自动从订单表的 `出行时间` 填入。

`config/trip_config.json` 只需要维护：

```json
{
  "route": "旅游路线",
  "leaders": "领队",
  "leader_phone": "领队电话"
}
```

如果不想每次在终端填写，可以直接使用配置文件：

```bash
./run.sh --no-prompt
```

## 手动指定表格

```bash
./run.sh --excel input/订单列表_2026-05-21_812378473774845952.xlsx
```
