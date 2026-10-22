# 样品寄送单生成器

一个基于 Python + Streamlit 的本地/可部署小工具，用于读取 `样品寄送模版.xlsx`，自动拆分中文地址、匹配商品 69 码，并生成样品寄送 Excel。

## 功能

- 读取模板中的：
  - `Sheet0`：寄送单输出模板
  - `Sheet1`：商品名称与 69 码映射
- 支持手动录入：
  - 多地址
  - 每个地址多个外部平台单号
  - 每个外部平台单号多个商品
- 支持批量上传多地址（同商品）
- 自动解析地址中的：
  - 收货人
  - 手机号
  - 省
  - 市
  - 区
  - 详细地址
- 商品可直接在下拉框中搜索
- 自动匹配并展示产品 69 码
- 生成前自动校验：
  - 地址解析失败
  - 外部单号缺失
  - 重复外部单号
  - 商品未匹配
  - 数量非正整数
- 生成结果自动保存到 `output` 文件夹

## 项目文件

- `sample_shipping_app.py`
  页面入口
- `address_utils.py`
  地址解析与校验
- `excel_utils.py`
  模板读取、批量上传解析、Excel 生成与保存
- `ui_components.py`
  页面组件
- `time_utils.py`
  统一北京时间
- `github_sync.py`
  线上更新模板时把文件提交回 GitHub
- `样品寄送模版.xlsx`
  Excel 模板
- `requirements.txt`
  Python 依赖
- `run.command`
  macOS 启动脚本

## 本地运行

### macOS

```bash
cd "/Users/xuxiaofeng/Desktop/样品寄送项目"
python3 -m pip install -r requirements.txt
python3 -m streamlit run sample_shipping_app.py
```

或直接双击：

- `run.command`

## 模板维护

请在 `样品寄送模版.xlsx` 中维护商品映射：

- `Sheet1`
  - A列：线上商品名称
  - B列：产品六九码

更新模板并保存后，回到页面点击：

- `刷新商品数据`

## 在线更新商品模板

页面底部有个「更新商品模板（管理员）」折叠区，可以直接上传新的 `样品寄送模版.xlsx`。

**线上默认隐藏这个入口**，需要打开带参数的链接才会出现：

```text
https://你的应用地址.streamlit.app/?admin=1
```

打开后仍是折叠状态，展开并输入管理密码后才能上传。普通用户直接访问分享链接时看不到这一块。

上传行为：

- **本地运行**：直接替换项目里的模板文件，立即生效。
- **线上（Streamlit Cloud）**：把文件提交回 GitHub 仓库，Streamlit Cloud 收到提交后自动重新部署，约 1 分钟后生效。

线上使用需要在 Streamlit Cloud 的 **Settings → Secrets** 里配置：

```toml
GITHUB_TOKEN = "ghp_你的token"
ADMIN_PASSWORD = "自定的管理密码"
```

- `GITHUB_TOKEN`：GitHub 个人访问令牌，需要该仓库的 Contents 读写权限
- `ADMIN_PASSWORD`：上传入口的密码；**线上必须配置**，否则上传入口不开放（应用是公开的，防止别人改模板）

## 批量上传格式

批量上传只支持 `.xlsx`。

表头必须为：

- `外部平台单号`
- `地址`

## 输出说明

- 生成的 Excel 会保存到 `output` 文件夹
- 支付日期会写入生成当天的静态值
- 客户信息支持选择：
  - `KH00832 / 第三方样品`
  - `KH00007 / 小红书`
- 固定写入字段：
  - 分销渠道 = 空
  - 子订单类型 = `样品`

## 部署到 Streamlit Community Cloud

这个项目可以直接上传到 GitHub 后部署到 Streamlit Community Cloud。

部署时：

1. 将整个项目上传到 GitHub 仓库
2. 在 Streamlit Community Cloud 中选择该仓库
3. 入口文件填写：

```text
sample_shipping_app.py
```

4. 依赖文件使用：

```text
requirements.txt
```

## 注意事项

- 不要删除模板文件 `样品寄送模版.xlsx`
- `output` 文件夹中的内容是生成结果，删掉不影响程序本身
- 首次启动时会自动安装依赖，可能需要等待几十秒
- 所有时间（刷新时间、输出文件名、支付日期）统一按北京时间（Asia/Shanghai）生成，不受服务器所在时区影响。Streamlit Cloud 服务器默认是 UTC，若不用本项目的时间处理会慢 8 小时
