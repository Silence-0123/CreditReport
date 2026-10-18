# 信贷提案智能生成器（Credit Proposal Generator）

一个面向银行客户经理的轻量级 Web 工具：选择贷款品种、录入尽调与财务素材，调用 GPT-4o 生成结构化授信提案内容，并回填到 Word 模板中。生成的分析段落会以**黄色高亮**标注，方便人工复核。

## 功能

- 贷款品种：流动资金贷款、固定资产贷款、普惠金融贷款。
- 模板：内置三种银行公文风格模板，也可上传自定义 `.docx` 模板。
- 素材：支持直接粘贴，或导入 `.txt` / `.md` / `.docx`。
- AI 输出：严格 JSON 结构，Key 与 Word 占位符一一对应。
- Word 处理：递归处理正文、表格、嵌套表格、页眉页脚；支持占位符被 Word 拆分到多个 run 的情况。
- 格式保留：替换时继承原 run 的字体、字号等字符格式，不重建整篇文档。
- 人工复核：四个 AI 分析段落默认黄色高亮，可在侧边栏关闭。
- 下载：生成 `.docx`，文件名自动包含公司名称、贷款品种和日期。

## 快速开始

```bash
# 1. 创建虚拟环境（建议 Python 3.10+）
python -m venv .venv
.venv\Scripts\activate        # Windows
# source .venv/bin/activate   # macOS / Linux

# 2. 安装依赖
pip install -r requirements.txt

# 3. 配置环境变量
copy .env.example .env        # Windows
# cp .env.example .env        # macOS / Linux
# 编辑 .env，填写 OPENAI_API_KEY；如使用兼容网关，同时填写 OPENAI_BASE_URL

# 4. 启动
streamlit run app.py
```

浏览器打开终端提示的地址（默认 <http://localhost:8501>）。

也可以在侧边栏直接输入 API Key；留空时读取 `.env` 或 Streamlit secrets。

## 部署到 Streamlit Community Cloud

本项目已按 Streamlit Community Cloud 的运行方式设计：不依赖本地磁盘持久化，Word 生成与下载均在内存中完成，API Key 支持 `st.secrets`。

### 1. 推送到 GitHub

确保仓库根目录至少包含：

```text
app.py
requirements.txt
.streamlit/config.toml
```

`requirements.txt` 必须位于仓库根目录或 `app.py` 同目录。`packages.txt` 不需要，因为项目没有 apt 系统依赖。

不要把以下文件提交到仓库：

```text
.env
.streamlit/secrets.toml
```

### 2. 创建应用

1. 打开 <https://share.streamlit.io>，登录后点击 **Create app**。
2. 选择 **Yup, I have an app**。
3. 选择 GitHub 仓库、分支（通常 `main`）和入口文件 `app.py`。
4. 可选：设置自定义子域名。

### 3. 配置 Python 版本与 Secrets

在部署页点击 **Advanced settings**：

- **Python version**：选择 `3.12`（当前默认）或其他仍受支持的版本。建议与本地开发版本保持一致。
- **Secrets**：粘贴以下 TOML 内容（替换为真实密钥）：

```toml
OPENAI_API_KEY = "sk-..."
OPENAI_BASE_URL = "https://api.openai.com/v1"
OPENAI_MODEL = "gpt-4o"
```

保存后点击 **Deploy**。应用代码读取顺序为：侧边栏输入 → 环境变量 → `st.secrets`，因此 Secrets 中的配置会自动生效。

如果使用银行内部或自建的 OpenAI 兼容网关，请确认该网关允许从 Streamlit Cloud 的公网出口访问，并且地址是 HTTPS。内网地址通常无法从 Community Cloud 直接访问。

### 4. 部署后更新

- 修改代码：推送到 GitHub 后，Community Cloud 会自动重新运行。
- 修改依赖：`requirements.txt` 变更后会自动重装，可能需要几分钟。
- 修改密钥：进入应用的 **Settings → Secrets** 更新，保存后应用会重启。
- 查看日志：应用右下角的 **Manage app** 中可查看部署与运行日志（需要仓库写权限）。

### 5. 重要安全提示

Streamlit Community Cloud 应用通常通过 URL 公开访问。即使 Secrets 不会直接显示在前端，任何访问者都可以使用你配置的 API Key 发起生成请求并消耗额度。银行真实客户资料场景请特别注意：

- 不建议在公开的 Community Cloud 应用中配置长期、高额度的生产 API Key；
- 仅使用演示数据或脱敏数据测试；
- 为 API Key 设置额度、速率限制和独立项目；
- 生产环境应部署在行内受控平台，接入 SSO、访问审计、数据脱敏和网络隔离；
- 客户敏感资料会发送到所配置的大模型服务，须符合所在机构的数据分级、外包与个人信息保护要求。

官方文档：[Secrets management](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/secrets-management) · [Deploy your app](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy)

## `.env` 配置

```dotenv
OPENAI_API_KEY=sk-...
OPENAI_BASE_URL=https://api.openai.com/v1
OPENAI_MODEL=gpt-4o
```

- `OPENAI_BASE_URL`：使用 OpenAI 官方接口时可保留默认值；使用兼容网关时填写网关地址，通常以 `/v1` 结尾。
- `OPENAI_MODEL`：默认 `gpt-4o`；如网关使用其他模型名，可在侧边栏“高级连接设置”中覆盖。
- 也支持 Streamlit Cloud 的 `st.secrets`，Key 名称相同。

## Word 模板占位符

模板中使用以下六个占位符（支持 `{{Key}}` 和 `{{ Key }}` 两种写法）：

| 占位符 | 来源 | 是否高亮 |
| --- | --- | --- |
| `{{公司名称}}` | AI 从素材提取 | 否 |
| `{{拟申请金额及期限}}` | AI 从素材提取 | 否 |
| `{{行业及经营状况分析}}` | AI 撰写 | 是 |
| `{{财务状况及还原调整结论}}` | AI 撰写 | 是 |
| `{{还款来源与压力测试分析}}` | AI 撰写 | 是 |
| `{{主要风险点及防控措施}}` | AI 撰写 | 是 |

说明：

- 黄色高亮仅作用于 AI 生成的文字，不改变模板的标题、字号、字体、行距等段落格式。
- 如果占位符在 Word 中被拆成多个 run（例如中途改过字体或拼写检查），替换逻辑会跨 run 处理。
- 如果上传的模板缺少部分占位符，界面会提示；缺少的内容不会写入文档。
- 如果模板中的占位符位于 Word 文本框、内容控件或特殊域中，可能无法识别；建议将其放在普通段落或表格单元格中。

## 使用建议

1. 尽调素材尽量包含：企业沿革、股权与实际控制人、行业与竞争、采购销售、订单、授信需求、还款来源、风险线索。
2. 财务数据尽量包含：近三年及最近一期资产负债表、利润表、现金流量表关键科目，以及周转率、负债结构、对外担保、还原调整说明。
3. 生成后重点核对黄色高亮段落中的数字、结论和风险措施；AI 不会替代尽调、审查与审批。
4. 如需调整公文风格，可下载内置模板后修改，再上传使用；不要删除六个标准占位符。

## 安全与合规

- API Key 只在当前进程内存中使用，不写入磁盘、不写入生成文档。
- `.env` 已加入 `.gitignore`，请勿提交真实密钥。
- 客户敏感资料会发送至所配置的大模型服务；请遵守所在机构的网络安全、数据分级、外包与个人信息保护规定。
- 本工具输出仅供授信审批辅助参考，不构成授信承诺；最终方案以有权审批人审批意见和正式合同为准。

## 常见问题

**Q：提示 `response_format` 不支持？**  
A：代码会自动降级为普通 Chat Completions 调用，仍通过 Prompt 约束 JSON 输出。若网关仍不支持，请检查其 OpenAI 兼容性。

**Q：生成内容不是 JSON / 解析失败？**  
A：应用会自动重试一次。若仍失败，请检查模型是否支持 JSON Mode，或换用 `gpt-4o` 等指令遵循能力较强的模型。

**Q：下载的 Word 中占位符没有被替换？**  
A：确认占位符拼写与上表完全一致；确认其位于普通段落或表格单元格中，而不是文本框、图片或内容控件中。

**Q：黄色高亮在哪里设置？**  
A：侧边栏“AI 分析段落黄色高亮”复选框，默认开启。高亮使用 Word 的文本突出显示（Yellow Highlight），不是单元格底纹。
