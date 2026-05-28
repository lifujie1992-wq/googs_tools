# Excel 拆分生成数据包 MVP

本地网页工具，用于把上传的 `.xlsx` 按“商品名称”拆分成多个商品文件夹，并打包下载 ZIP。

## 功能

- 上传 `.xlsx` 文件并解析预览
- 按 `商品名称` 分组
- 每个商品生成独立文件夹
- 每个商品文件夹内生成 `商品数据.xlsx`
- 可选上传图片 ZIP，按 Excel 中的图片文件名复制到对应目录
- 生成总数据包 ZIP 下载

## Mac 调试启动

需要 Python 3.11 或更新版本。

```bash
cd excel-splitter
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app:app --host 127.0.0.1 --port 8123 --reload
```

然后打开：

```text
http://127.0.0.1:8123/
```

也可以直接运行：

```bash
chmod +x start-mac.sh
./start-mac.sh
```

## Excel 格式要求

第一行必须是表头，并且包含以下任一商品名称字段：

- `商品名称`
- `商品名`
- `产品名称`
- `产品名`
- `name`
- `product name`

图片字段支持：

- `主图`
- `主图文件名`
- `商品主图`
- `详情图`
- `详情图文件名`
- `详情页图`
- `商品详情页图`
- `商品信息图`
- `信息图`
- `颜色图`
- `颜色图文件名`

图片 ZIP 中的文件名需要和 Excel 单元格中的文件名一致。一个单元格里可以写多个图片文件名，用逗号、中文逗号、分号或中文分号分隔。

## 目录说明

```text
app.py              FastAPI 后端
xlsx_tools.py       轻量 xlsx 读写工具
static/             前端页面
work/               运行时临时目录，首次启动会自动创建
requirements.txt    Python 依赖
start-mac.sh        macOS 一键启动脚本
Dockerfile          服务器容器镜像配置
docker-compose.yml  服务器一键启动配置
deploy/             Ubuntu 部署和更新脚本
```

## Ubuntu 服务器一键部署

先把本项目推到 GitHub，然后在 Ubuntu 服务器执行：

```bash
curl -fsSL https://raw.githubusercontent.com/你的用户名/你的仓库/main/deploy/install-ubuntu.sh | sudo bash -s -- https://github.com/你的用户名/你的仓库.git
```

默认端口是 `8123`。如需改端口，例如 `80`：

```bash
curl -fsSL https://raw.githubusercontent.com/你的用户名/你的仓库/main/deploy/install-ubuntu.sh | sudo APP_PORT=80 bash -s -- https://github.com/你的用户名/你的仓库.git
```

部署后访问：

```text
http://服务器IP:8123/
```

后续更新：

```bash
cd /opt/excel-splitter
sudo ./deploy/update-ubuntu.sh
```

## 输出数据包格式

生成的 ZIP 会按商品名称建父文件夹，Excel 文件名使用资料编码：

```text
商品名称/
  商品数据-资料编码.xlsx
  商品主图/
  商品详情页图/
  商品信息/
  颜色图/
```

四个图片子文件夹即使为空，也会写入 ZIP，文件夹名字保持一字不差。

## 说明

这版为了方便迁移，没有依赖 `openpyxl`，只实现了当前业务需要的 `.xlsx` 读写能力。后续正式版如果要兼容更复杂的 Excel 样式、公式、合并单元格，建议换成 `openpyxl`。
