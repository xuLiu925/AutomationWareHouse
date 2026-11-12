# FreightPilot 服务端（Cloudflare Worker）

`worker.js` 负责三件事：

- **服务开关**：程序每次运行前问这里服务是否开着；关掉或连不上都不能用（同事不用授权码，窗口底部有说明）
- **AI 整理中转**：Gemini Key 和提示词只放在这里，exe 里没有
- **产品托盘参数**：所有人共用一份，带版本历史；同事提交修改申请，管理员审核
- **飞书邮件授权**（对账功能）：保管飞书应用的 App Secret，把授权 code 换成每个人自己的 token；邮件内容不经过这里

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

## 二、同事不用授权码；管理员码只给自己

同事装上就能用：程序每次运行前问服务器服务是否开着，提交产品参数修改时用 Windows 用户名显示是谁提交的。

管理员（产品参数直接保存、审核申请、回滚版本）靠一个只放在自己电脑上的码：
1. 生成一个随机码：
   ```
   python -c "import secrets; print(secrets.token_urlsafe(16))"
   ```
2. 在 KV `xlfp` 里点 Add entry：Key `user:<码>`，Value `{"name": "刘旭", "role": "admin"}`
3. 在自己电脑的程序目录建 `xlfp_config.json`：`{"token": "<码>"}`（不纳入 Git、不打包；同事电脑上没有这个文件）

## 三、日常管理（改完 1 分钟内生效，不用重新打包）

| 要做的事 | 怎么改 |
|---|---|
| 暂停服务（所有功能、所有人） | 加/改键 `config`，值为 `{"enabled": false, "message": "提示文字"}`（不写 message 显示"服务暂停使用，请联系刘旭"） |
| 重新开放 | `enabled` 改回 `true` |
| AI 每日总次数 | `config` 里加 `"ai_daily": 500`（全软件每天最多几次，一张表 / 一次对账通常算一次；0 或不写 = 不限）。用完当天改用规则处理，不报错 |
| 看今天用了多少次 | KV 里 `usage:<日期>:all`（总数）和 `usage:<日期>:win:<Windows 用户名>`（每人）这些键（2 天后自动删除） |
| 换 Gemini Key | Settings → Variables and Secrets 里改 `GEMINI_API_KEY` |
| 改提示词 | Edit code 改 `PROMPT`，再点 Deploy；同时把 `worker.js` 存回仓库 |
| 欢迎画面的每日一句 | 自动：每天（德国日期）第一个打开软件的人触发 1 次 Gemini，存在 `greeting:<日期>`，当天所有人看到同一句；想换一句就删掉当天这个键。提示词在 `GREETING_PROMPT` |

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

## 五、飞书邮件（对账功能）

对账要读每个人自己的飞书邮箱。飞书应用的 App ID / App Secret 只放在 Worker 里，不写进程序和公开仓库。

1. Worker 的 Settings → Variables and Secrets → Add：
   - Text 变量 `FEISHU_APP_ID`：飞书开发者后台的 App ID
   - Secret `FEISHU_APP_SECRET`：App Secret
   - （可选）Text 变量 `FEISHU_REDIRECT_URI`，不加就是 `http://localhost:8080/callback`
2. 飞书开发者后台：
   - 「安全设置」→ 重定向 URL 里有 `http://localhost:8080/callback`
   - 「权限管理」→ 用户身份权限开通：`mail:user_mailbox.message:readonly`、`mail:user_mailbox.message.body:read`、
     `mail:user_mailbox.message.subject:read`、`mail:user_mailbox.message.address:read`、`mail:user_mailbox.folder:read`（读自建文件夹，报价邮件常被邮箱规则收进去），发布版本并审批通过
   - 应用的「可用范围」要包含所有要用对账的同事，否则他们授权时会被拒绝
3. 把新的 `worker.js` 粘贴进去 Deploy。

同事第一次对账时，程序会打开浏览器，他用自己的飞书账号点"授权"即可；token 只存在他自己电脑的 `feishu_token.json`，约 2 小时有效。
以后在飞书后台开通 `offline_access` 并发布新版本，再把 `feishu_mail.py` 的 `SCOPES` 加上 `offline_access`，程序就会自动续期，不用每次授权。

## 六、在线更新（R2）

程序打开时会问服务器有没有新版本，有就提示同事一键更新（也可以点"检查更新"）。更新包放在 Cloudflare R2，服务开着时程序才能下载。

**开通（只做一次）**
1. 左侧 Storage & Databases → **R2** → 开通（免费额度：10 GB 存储，下载流量免费；开通时 Cloudflare 要求绑定付款方式）。
2. Create bucket，名字填 `xlfp-releases`（Location 默认即可，不要开公开访问）。
3. 回到 Worker `xlfp` → Settings → Bindings → Add → **R2 bucket**，Variable name 填 **`RELEASES`**，bucket 选 `xlfp-releases` → Deploy。
4. 把最新的 `worker.js` 粘贴进去 Deploy（里面有 `/update/latest`、`/update/download`）。

**每次发布新版本**
1. 打包：`python build.py 这次更新了什么`（后面的文字可选，会显示在同事的更新提示里）。
2. 打包完 `dist\release\` 里有两个文件：`FreightPilot-<版本>.zip` 和 `latest.json`。
3. R2 → `xlfp-releases` → Upload：**先传 zip，再传 `latest.json`**（覆盖旧的；顺序反了，同事可能下载到校验不一致的包，重试即可）。
4. 同事下次打开程序就会提示更新。旧的 zip 可以留着，也可以删掉（只会下载 `latest.json` 指向的那个）。

**命令行自动上传（远程发版用）**
- 一次性准备：`npm install -g wrangler`；Cloudflare → My Profile → API Tokens → 自定义 token，只给 Account 的 **Workers R2 Storage: Edit**；存成 Windows 用户环境变量 `CLOUDFLARE_API_TOKEN` 和 `CLOUDFLARE_ACCOUNT_ID`（只在本机，不进仓库）。
- `python build.py 这次更新了什么 --upload`：打包、自检通过后自动先传 zip 再传 `latest.json`，读回 R2 上的 `latest.json` 核对版本，再问一次 Worker `/update/latest`。zip 传失败就不动 `latest.json`。
- `python build.py --upload-only`：只把上次打好的 `dist\release\` 再传一次（上传失败后重试）。

**说明**
- 版本号是打包时间（如 `2026.11.07.1530`），写在程序目录的 `version.txt`，只比新旧，不用手动改。
- 更新时程序下载 → 校验 sha256 → 关闭 → 覆盖程序文件 → 重新打开。管理员码、飞书授权、邮件缓存、Dachser 登录信息、危险品文件不在更新包里，不会动；本机已有的 `pallet_params.json` 不覆盖。
- 服务暂停时查不到更新、也下载不了。
- 没有在线更新功能的旧版本（2026-11-07 之前打包的）需要手动换一次新版，之后就能在线更新。

## 七、售后助手（单独的程序，不在 exe 里）

`/aftersales/read`（读 Dachser 售后邮件，起草给客户的中文消息）和 `/aftersales/reply`（把客户回复起草成德语邮件）给开发者自己的售后助手用，
不用授权码，每次调用计入 AI 每日总次数。不用额外配置，把最新的 `worker.js` 粘贴进去 Deploy 即可。

## 八、费用和条款

- 在欧盟把程序给同事使用，Gemini 必须用**付费服务**（Gemini API 条款："You may use only Paid Services when making API Clients available to users in the European Economic Area"）。
  在 Google AI Studio 的项目里开通 Billing 即可，按用量计费。
- 建议在 Google Cloud 的 Billing → Budgets 里设一个每月预算提醒（比如 5 欧元）。
- Worker 发给 Gemini 的只有：地址、House No、Address2、邮编、城市、国家、收件人、公司，询价再加编号和尾板。不包含电话和邮箱。
- 售后助手发给 Gemini 的是 Dachser 售后邮件的主题、发件人、正文，以及开发者转给自己的客户回复。
- 对账时，只有表格读不出报价的邮件，才会把那封邮件的主题和正文（只含回复本身，去掉引用部分，最多 6000 字）发给 Gemini。
