# XL FreightPilot 服务端（Cloudflare Worker）

`worker.js` 负责三件事：

- **授权**：每个同事一个授权码，总开关和每个人的开关都在这里控制
- **AI 整理中转**：Gemini Key 和提示词只放在这里，exe 里没有
- **产品托盘参数**：所有人共用一份，带版本历史；同事提交修改申请，管理员审核

全部在 Cloudflare 网页后台操作，不用装任何工具。免费额度为每天 10 万次请求，用不完。

## 一、部署（只做一次）

1. 注册并登录 https://dash.cloudflare.com 。
2. **建 KV**：左侧 Storage & Databases → KV → Create，名字填 `xlfp`。
3. **建 Worker**：左侧 Compute (Workers) → Create → 选 "Hello World" 模板，名字填 `xlfp` → Deploy。
   然后点 Edit code，把 `worker.js` 的全部内容粘贴进去，替换原来的代码，再点 Deploy。
4. **绑定 KV**：在这个 Worker 的 Settings → Bindings → Add → KV namespace。
   Variable name 填 **`XLFP`**（必须大写），namespace 选第 2 步建的 `xlfp`。
5. **放 Gemini Key**：Settings → Variables and Secrets → Add。
   Type 选 **Secret**，Name 填 `GEMINI_API_KEY`，Value 填 Google AI Studio 里的 Key。
   （可选）再加一个 Text 类型的变量 `MODEL`，可以换模型，不加就用 `gemini-3.5-flash-lite`。
6. **记下 Worker 地址**：类似 `https://xlfp.<你的子域名>.workers.dev`。
   把它填进 `license_check.py` 的 `SERVER_URL`，然后重新打包（只需要这一次）。

## 二、给同事发授权码

1. 生成一个随机授权码：
   ```
   python -c "import secrets; print(secrets.token_urlsafe(16))"
   ```
2. 在 KV `xlfp` 里点 Add entry：
   - Key：`user:<授权码>`，例如 `user:Q3v9...`
   - Value：`{"name": "张三", "enabled": true, "daily": 100}`
     - `daily`：每天最多做几次 AI 整理（一张表算一次），写 0 或不写表示不限
3. 把授权码发给同事。同事在软件里点"授权码"按钮填进去即可。

给自己也建一个授权码，并加上 `"role": "admin"`（管理员：产品参数直接保存、审核申请、回滚版本）：
`{"name": "刘旭", "enabled": true, "role": "admin"}`

## 三、日常管理（改完 1 分钟内生效，不用重新打包）

| 要做的事 | 怎么改 |
|---|---|
| 停用某个同事 | 把他那条的 `enabled` 改成 `false`，可加 `"message": "提示文字"` |
| 停用所有人 | 加/改键 `config`，值为 `{"enabled": false, "message": "提示文字"}` |
| 重新开放 | 改回 `true` |
| 看谁今天用了多少次 | KV 里 `usage:<日期>:<授权码>` 这些键（2 天后自动删除） |
| 换 Gemini Key | Settings → Variables and Secrets 里改 `GEMINI_API_KEY` |
| 改提示词 | Edit code 改 `PROMPT`，再点 Deploy；同时把 `worker.js` 存回仓库 |

## 四、产品托盘参数和审核

KV 里的键（程序自动维护，不用手改）：

| 键 | 内容 |
|---|---|
| `products` | 当前生效的参数：`{"version", "updated", "by", "note", "data"}`，`data` 就是 `pallet_params.json` 的内容 |
| `products:v:<版本号>` | 每个版本的快照，回滚用 |
| `proposal:<id>` | 修改申请（提交人、内容、状态、驳回原因）；审完 90 天后自动删除 |

- **第一次**：部署新的 `worker.js`、给自己加上 `role` 后，在软件里点"产品参数"，会提示把本机的
  `pallet_params.json` 上传为第 1 版。之后以服务器为准，不要再手改同事电脑上的文件（会被覆盖）。
- **审核**：同事提交后，你点"开始"询价时运行记录里会提示"有 N 条待审核"。在"产品参数"→"审核"页
  能看到修改前后的对比和校验结果，点"通过"立即对所有人生效，点"驳回"可以写原因。
- **回滚**："历史版本"页选一个版本 →"回滚到选中的版本"（另存为新版本，旧记录都保留）。
- 校验规则在程序里（`pallet_plan.validate_product`），服务器只检查格式和权限。
- KV 是最终一致的：刚提交的申请可能要几十秒后才出现在审核列表里。

## 五、费用和条款

- 在欧盟把程序给同事使用，Gemini 必须用**付费服务**（Gemini API 条款："You may use only Paid Services when making API Clients available to users in the European Economic Area"）。
  在 Google AI Studio 的项目里开通 Billing 即可，按用量计费。
- 建议在 Google Cloud 的 Billing → Budgets 里设一个每月预算提醒（比如 5 欧元）。
- Worker 发给 Gemini 的只有：地址、House No、Address2、邮编、城市、国家、收件人、公司，询价再加编号和尾板。不包含电话和邮箱。
