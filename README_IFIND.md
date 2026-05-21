# iFinD API 与 ETF 投资助手 Demo

本项目使用同花顺 iFinD Python SDK `iFinDPy` 获取 ETF 数据，并提供一个 Streamlit demo：

- 投资者偏好问卷
- ETF 推荐候选
- 单只 ETF 质量评估
- 投资前自查问题

所有结果仅用于学习和决策辅助，不构成投资建议。

## 1. 安装依赖

```powershell
python -m pip install -r requirements.txt
```

如果只想验证 iFinD SDK：

```powershell
python -m pip install iFinDAPI
python -c "import iFinDPy; print('iFinDPy import ok')"
```

## 2. 设置 iFinD 账号密码

账号密码不要写入代码，请在当前 PowerShell 窗口设置环境变量：

```powershell
$env:IFIND_USERNAME="你的iFinD账号"
$env:IFIND_PASSWORD="你的iFinD密码"
```

如果没有设置账号密码，demo 仍能打开，但只使用 Excel 静态样本和已有缓存。

## 3. 运行登录检查

```powershell
python .\ifind_setup_check.py
```

成功时会看到：

```text
iFinD login succeeded.
```

## 4. 启动 ETF 投资助手 demo

```powershell
streamlit run .\app.py --browser.gatherUsageStats false
```

打开 Streamlit 给出的本地地址，一般是：

```text
http://localhost:8501
```

## 5. 数据与缓存

第一版固定读取 `上证ETF(1).xlsx` 的前 10 只 ETF 作为样本池。

缓存文件会自动生成在：

```text
data/cache/etf_sample_profile.csv
data/cache/etf_sample_prices.csv
```

如果缓存存在，demo 默认优先读取缓存，不重复调用 iFinD。需要刷新时，在页面侧边栏点击“重新拉取 iFinD 行情”。

## 6. 可选：接入 DeepSeek 解释润色

DeepSeek 只用于把规则和数据结论润色成更易读的中文解释，不参与评分和排序。未设置 API Key 时，demo 会自动使用本地规则模板。

```powershell
$env:DEEPSEEK_API_KEY="你的DeepSeek API Key"
$env:DEEPSEEK_BASE_URL="https://api.deepseek.com"
$env:DEEPSEEK_MODEL="deepseek-v4-flash"
streamlit run .\app.py --browser.gatherUsageStats false
```

## 7. 免费版注意事项

iFinD 免费版可能限制可用字段、调用频率或数据量。第一版只拉取近 1 年日线行情，并优先使用本地缓存，降低额度消耗。

常见问题：

- `No module named 'iFinDPy'`：执行 `python -m pip install iFinDAPI`。
- 登录失败：检查账号密码、接口权限和网络。
- 字段为空或权限错误：该字段可能不在免费权限内，先换少量 ETF 或更基础字段测试。
