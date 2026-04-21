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
- `样品寄送模版.xlsx`
  Excel 模板
- `requirements.txt`
  Python 依赖
- `run.command`
  macOS 启动脚本
- `run_windows.bat`
  Windows 启动脚本

## 本地运行

### macOS

```bash
cd "/Users/xuxiaofeng/Desktop/样品寄送项目"
python3 -m pip install -r requirements.txt
python3 -m streamlit run sample_shipping_app.py
```

或直接双击：

- `run.command`

### Windows

双击：

- `run_windows.bat`

如果窗口一闪而过，请打开项目文件夹，在顶部地址栏输入 `cmd`，回车后执行：

```bat
run_windows.bat
```

## 模板维护

请在 `样品寄送模版.xlsx` 中维护商品映射：

- `Sheet1`
  - A列：线上商品名称
  - B列：产品六九码

更新模板并保存后，回到页面点击：

- `刷新商品数据`

## 批量上传格式

批量上传只支持 `.xlsx`。

表头必须为：

- `外部平台单号`
- `地址`

## 输出说明

- 生成的 Excel 会保存到 `output` 文件夹
- 模板中的支付日期会保留 `TODAY()` 公式，不会被写成静态值
- 固定写入字段：
  - 客户code = `KH00832`
  - 客户名称 = `第三方样品`
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
