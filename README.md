# XL FreightPilot

**Freight Quotation & Order Automation for Dachser eLogistics** · Developed by Xu Liu（刘旭）

把物流同事每天在 Dachser eLogistics 网站上手动做的两件事自动化：

- **询价**：读取询价表 Excel，自动计算托盘方案，逐单填写 freightquotation 的 Enquiry 表单
- **建单**：读取领星导出的出库单，按备注里的 MR 号找到对应报价并点 create，自动填写运输订单（地址、联系人、单号、危险品、上传危险品文件）

脚本只负责**填写**，最后的 Enquiry / Save 由人确认后手动点击。

---

## 功能概览

### 询价（`dachser_fill.py`）
1. 读询价表：`运费报价(欧元)` 列合并的几行属于同一订单，会合并成一张询价单
2. 按 `pallet_params.json` 里的产品参数计算托盘方案（整托盘 + 零散托盘，最多 2 行）
3. 自动登录 → 进入 freightquotation → 填国家、邮编、城市、地址、货物行、危险品、尾板、电话预约等
4. 你检查后点 Enquiry 提交，脚本检测到结果页后自动做下一单

### 建单（`create_orders.py`）
1. 读领星出库单的 `Remark/备注`，取出 MR 编号
2. 在 My enquiries 按单号搜索（找不到时扩大到整个日期范围再本地筛选），忽略已 created 的报价
   - 唯一匹配：自动点 create（地址和领星不一致只提醒，不拦截）
   - 多个匹配：列出候选，由你在浏览器里选
3. 在 Transport Order 页面自动填写：
   - 收货地址 Street / Country / PCD / City 以领星为准（County 不改）
   - Name 1 = 公司名（没有则收件人），Name 2 = 邮箱，Name 3 / Telephone / Mobile = 电话
   - Delivery options 选 Phone delivery booking，Contact person = 收件人
   - Order number 后加 `+出库单号`（超长时截掉原单号前部，出库单号保持完整）
   - 货物行 Content = `battery`，并为每行添加危险品 UN3480（重量、件数、Fibreboard box）
   - 上传与 SKU 同名的危险品文件
4. 你检查后点 Save；离开前脚本会确认是否已保存，避免 Dachser 锁单
5. 结果写入 Excel 旁边的 `<Excel名>_建单结果.csv`

---

## 给同事使用（打包版）

收到 `XL FreightPilot.zip` 后：

1. 解压到任意位置，**整个文件夹一起用**，不要只拷贝 exe
2. 电脑需要安装 **Google Chrome**
3. 双击 `XL FreightPilot.exe`，第一次在弹出的 Chrome 里登录 Dachser（建议让 Chrome 记住密码，以后会自动登录）
4. 选功能 → 选 Excel → 开始

详细说明见 [`使用说明.txt`](使用说明.txt)（会一起打包进去）。

---

## 开发：直接运行源码

```powershell
pip install playwright openpyxl
playwright install chromium     # 可选：脚本默认使用本机已安装的 Chrome
```

```powershell
python gui.py                                     # 窗口界面（和打包版一样）

python dachser_fill.py 询价表.xlsx --all --dry      # 询价：只预览托盘方案
python dachser_fill.py 询价表.xlsx --all            # 询价：批量填写

python create_orders.py 领星导出.xlsx --dry         # 建单：只搜索匹配，不点 create
python create_orders.py 领星导出.xlsx               # 建单：搜索 + create + 填写
python create_orders.py 领星导出.xlsx --html eLogistics.html   # 建单：离线匹配预览

python pallet_plan.py US2000 44                   # 单独算托盘方案
```

命令行模式下，需要确认的地方在终端里输入；窗口模式下是弹窗。改完代码直接运行即可验证，满意后再打包。

---

## 打包

```powershell
python build.py
```

- 生成 `dist\XL FreightPilot\`（发给同事的文件夹）和 `dist\XL FreightPilot.zip`
- 打包前请关闭正在运行的 `XL FreightPilot.exe`
- 用 Nuitka 编译为原生程序，不包含 `.py` 源码，也**不包含你的登录信息**（`dachser_profile`）
- Nuitka 不支持中文路径：项目在纯英文路径（`D:\XLFreightPilot`）时直接用当前 Python 在 `build/` 里编译；否则把源码复制到 `D:\liuxu_build`（环境变量 `LIUXU_BUILD_DIR` 可改）并在那里建专用虚拟环境编译；第一次约 10–20 分钟，之后几分钟
- 打包后自动运行 `--selftest` 检查 Chrome、托盘参数和授权；看到 `自检结果: OK ...` 即成功
- 新增 `.py` 模块时，要把文件名加到 `build.py` 的 `SOURCES` 里
- 只改了 `pallet_params.json` 不需要重新打包，直接替换同事文件夹里的那个文件

---

## 配置

| 文件 / 目录 | 说明 |
|---|---|
| `pallet_params.json` | 产品托盘参数（单箱尺寸重量、每层个数、最大层数、Dachser 包装类型）；`_defaults` 为通用默认值。放在 exe 旁边，可直接修改 |
| `危险品文件/` | 建单时上传的危险品文件，文件名（不含扩展名）= 领星 SKU，例如 `11US2000ZZD-20001.pdf` |
| `dachser_profile/` | Chrome 登录信息（自动生成，已 gitignore，不会打包） |
| `license_check.py` | 授权开关地址，见下文 |

### 授权开关

软件每次点"开始"前读取公开仓库 [`xuLiu925/liuxu-license`](https://github.com/xuLiu925/liuxu-license) 里的 `license.json`：

```json
{"enabled": true, "message": ""}
```

- `enabled: false` 时停止使用，显示 `message`（留空则用默认提示）
- 连不上 GitHub 或文件格式不对时也不能使用
- 在 GitHub 网页上修改即可，无需重新打包；GitHub 缓存约 5 分钟
- 停用只是不再运行自动化，不删除、不修改任何文件

---

## Excel 格式

### 询价表（第 1 行表头）
`编号`、`国家`、`邮编`、`城市`、`地址`、`数量`、`Unit/单位`、`单个物料（长）`、`单个物料（宽）`、`单个物料（高）`、`单个物料（重量）`、`Dangerous Goods Type/危险品类型`、`产品物料`、`是否需要尾板`（前缀匹配，如 `是否需要尾板车辆`）；可选 `运费报价(欧元)`（合并单元格 = 同一订单）、`备注…`

- 一个订单的每项信息都写在一个单元格里（需要多行就合并单元格，或在格内换行），不要把补充信息拆到下面几行
- `编号`：取开头连续的字母/数字/`-_/.` 作为 Order number，后面的文字当备注；认不准（开头是中文、太短、空格后接字母数字、紧跟 `#` 等符号）时开始前逐个询问
- `是否需要尾板`：先找否定词（不/无/否/没/no/nein/nicht/kein），再找肯定词（需要/要/是/有/尾板/yes/ja/tail）；空或认不出时开始前询问
- 编号格里的备注和 `备注` 列脚本不处理，只在运行记录里列出来提示人工查看
- `地址` → Street 栏（只填街道 + 门牌号，保留原顺序）：去掉括号备注、末尾国家（粘连的也去）、邮编、城市、`z.H./c/o/Attn` 前缀、开头的收件人/公司名；剩下像邮编的数字或没有门牌号时提示核对

### AI 整理（可选）
设置环境变量 `GEMINI_API_KEY`（Google AI Studio 的 Key）后，询价会先把每单的编号、地址、邮编、城市、国家、尾板、收件人、公司发给 Gemini 整理（不发电话、邮箱），模型默认 `gemini-3.5-flash-lite`，可用 `XLFP_AI_MODEL` 覆盖。
- AI 结果要过规则校验才用：编号必须原样出自编号格；Street 的每个词必须出自原地址且不含邮编；与规则结果冲突的编号/尾板开始前询问
- 没有 Key 或调用失败时自动使用规则识别，行为与之前相同
- 设置环境变量后要重新打开程序才生效；Key 不要写进代码或打包进 exe
- 在欧盟给同事使用前需开通 Gemini 付费（条款要求面向 EEA 用户只能用付费服务），并考虑用中转服务保管 Key

### 领星出库单
领星直接导出即可。用到的列：`Outbound Order No`、`Remark/备注`、`Recipient/收件人`、`Company`、`Country/Region`、`City`、`Post code`、`House No.`、`Address1`、`Address2`、`Province Name`、`Telephone`、`Email`、`Total Qty of SKU`、`SKU 1\nSKU`（多 SKU 时 `SKU 2\nSKU` …）

选错表格时软件会提示应改选"询价"还是"建单"。

---

## 项目结构

```
gui.py              窗口界面（打包入口），后台线程运行自动化，print → 运行记录，ask() → 弹窗
app_ui.py           APP_DIR（程序目录）、ask() 交互、停止标记 —— 命令行和窗口共用
license_check.py    授权开关
dachser_fill.py     询价：Excel 解析、Filler（Playwright 页面封装）、登录、批量填写
ai_clean.py         询价：可选的 Gemini 整理（编号/备注/Street/尾板），没有 Key 时不启用
pallet_plan.py      托盘方案计算
create_orders.py    建单：搜索报价、create、填写运输订单、危险品弹窗、上传文件、防锁单
match_orders.py     领星表读取、地址比对（也可单独生成匹配结果 CSV）
fetch_dachser.py    match_orders.py live 模式用的简单抓取
build.py            Nuitka 打包
使用说明.txt         给同事的说明（打包进去）
eLogistics.html     保存的 My enquiries 列表页，用于离线测试（含业务数据，勿公开）
```

---

## 常见问题

- **"The order is being processed by someone else"**：建单编辑页没点 Save/Cancel 就离开或关浏览器，Dachser 会锁单。现在脚本离开前会确认；已被锁的单要等 Dachser 会话过期后才能再打开
- **Chrome 顶部"不受支持的命令行标记 --no-sandbox"**：已通过 `chromium_sandbox=True` 去掉
- **杀毒软件 / "已保护你的电脑"**：未签名的 exe 可能被拦截，点"更多信息 → 仍要运行"或加白名单
- **弹窗找不到**：看任务栏里 "XL FreightPilot - 请确认"，或点一下主窗口，弹窗会回到最前面

---

© 2026 Xu Liu（刘旭）. All rights reserved. · xuliu00925@gmail.com
