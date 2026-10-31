# AGENTS.md — Voice Edge Auth Sync

本文件只放**规范**：协作规则、目录约定、构建与验证流程、硬性禁令。
**不要在这里写实现细节**（协议参数、状态机、常量、踩坑记录一律放 `.memory/MEMORY.md`）。
需要了解"现在是怎么实现的"，去看 `.memory/MEMORY.md`。

## 1. 项目与模块职责

`voice-edge-auth-sync` 是配合本地 Voice Edge 服务的跨浏览器扩展（Firefox MV2 / Chromium MV3），单一源码树构建两端。模块职责（只到"这块负责什么"，不写实现方式）：

| 文件                        | 职责                                                                                           |
| --------------------------- | ---------------------------------------------------------------------------------------------- |
| `src/background.js`         | 后台核心：Native Messaging、各站点凭证采集与调度、M365 会话桥与令牌获取、SharePoint 预续期调度 |
| `src/browser-polyfill.js`   | `browser.*` Promise 化兼容垫片（Firefox 无操作 / Chromium 包装）                               |
| `src/content-doubao.js`     | 豆包页面侧的 SSE 捕获与转发（仅 Chromium 需要）                                                |
| `src/content-deepseek.js`   | DeepSeek 会话与模型模式快照                                                                    |
| `src/content-m365.js`       | M365 页面主世界钩子：会话、保活、令牌事件、页面侧请求改写                                      |
| `src/content-m365-login.js` | Entra ID 登录页的受控自动填密与失败识别                                                        |
| `src/content-sharepoint.js` | SharePoint 上传、回读校验、Cookie/Digest/Token 预续期                                          |
| `src/popup.*`               | 面板视图、状态渲染、设置项                                                                     |
| `src/icons/`                | 图标资源                                                                                       |

## 2. 目录与构建

```text
build.mjs              # 构建脚本：自动运行 Prettier、合并 manifest，按引用复制 src → dist/<target>/
.prettierrc.json       # Prettier 格式配置
.prettierignore        # Prettier 忽略列表
manifest.base.json     # 两端共享字段
manifest.firefox.json  # Firefox MV2 覆盖字段
manifest.chrome.json   # Chromium MV3 覆盖字段
src/                   # 唯一权威源码
test/reauth.mjs        # 端到端 harness（单文件，见 §5）
.memory/               # 本地 Agent 笔记（不入库）：
                       #   MEMORY.md        当前实现的唯一事实来源
                       #   YYYY-MM-DD.md    当天变更、证据、结论
dist/                  # 构建产物（已 gitignore）
```

```bash
npm run build          # 自动 Prettier 格式化并构建两端（等价 node build.mjs all）
npm run build:chrome   # 只构建 chrome（harness 校验的是 dist/chrome）
npm run format         # 全仓 Prettier 格式化（prettier --write .）
npm run format:check   # 全仓 Prettier 格式检查（prettier --check .）
npm run test:reauth    # 离线 harness
```

## 3. 硬性禁令

1. **只改 `src/` 与顶层 `manifest.*.json`**。`dist/` 由构建生成，手改会被下次构建覆盖。
2. **每次改完必须构建并语法自检**：`npm run build` + `node --check src/*.js`。
3. **未经明确许可，不得改变任何 Git 状态**：禁止 `git add` / `git stage` / `git commit`；禁止 `git restore --staged`、`git reset`、`git rm --cached` 等任何触碰暂存区的操作。用户已暂存的内容由其全权掌控。
4. **不得用 `checkout` / `HEAD` / `restore` / `reset --hard` / `clean` / `stash` 覆盖工作区里未提交的改动**；需要对照历史版本时只用只读手段（`git show HEAD:<path>`、`git diff HEAD -- <path>`）。
5. **绝对不要暂存或提交 `.memory/`**。
6. **注释禁令**：注释/docstring 里不得引用任何外部文件（抓包、bundle、外部源码、二进制地址），不得写外部文件的版本号，不得写日期。注释只描述协议事实本身。

## 4. 跨浏览器约束（约束，不是实现说明）

- 统一走 `src/browser-polyfill.js` 的 Promise 风格 `browser.*`，禁止裸调回调式 `chrome.*`。
- 两端的**响应体读取能力不同**（Firefox 有 `webRequest.filterResponseData`，Chromium 没有），所以任何"读响应体"的需求都必须准备页面世界钩子这条兜底路径。
- 需要读页面全局变量、包装 `fetch`/`XHR`/`WebSocket` 时，脚本要注入主世界；目标站点可能有严格 CSP，注入必须能拿到 nonce，并要能判定"注入是否成功"。
- MV3 的 service worker 会随时被回收：任何跨重启还要用的状态都必须持久化，不能只放内存。

## 5. 验证

- **harness 测的是 `dist` 产物，不是 `src` 的另一份拷贝**。改完代码先构建，再跑；`src` 与 `dist` 的 sha256 不一致时 harness 直接 exit 3。
- `test/reauth.mjs` 是**唯一**的 harness 文件，不要再拆模块。三种模式：
  - `npm run test:reauth`：离线、无窗口，可进 CI；
  - `npm run test:reauth:live`：驱动真实已加载的扩展，**占用 5003**（该端口同一时刻只能有一个监听者，被 Voice Edge 服务占着时需要先停它）；
  - `npm run test:reauth:visual`：逐个弹出真实窗口演示登录页。
- **行为有变更时同步 bump `M365_BUILD_TAG`**：live 模式靠它判断扩展是否需要热重载；标签没变时会认为"已是最新"，从而测到旧代码。
- live 模式可以通过 5003 下发 `M365_RELOAD_SELF` 让扩展自己重载，不必人工去 `about:debugging` 点重载。

## 6. 调试

- 后台控制台执行 `__veM365SetDebug(true)`（或 Python 端 `M365_DEBUG=1`）打开详细追踪；默认只输出 `[VE-m365-bg]` 主日志。

## 7. 文档分工

| 文件                    | 写什么                                                                       | 不写什么                               |
| ----------------------- | ---------------------------------------------------------------------------- | -------------------------------------- |
| `AGENTS.md`（本文件）   | 规范、目录、构建/验证流程、禁令                                              | 实现细节                               |
| `.memory/MEMORY.md`     | **当前实现的事实**：不变量、协议参数、常量及其依据、已知坑、明确否决过的做法 | 过程流水账                             |
| `.memory/YYYY-MM-DD.md` | 当天做了什么、证据（日志/实测数字）、结论                                    | 已被推翻的旧结论（直接改掉，不要堆叠） |

写文档时保持"一份事实只有一个出处"：同一个机制不要在两处各写一遍，否则迟早互相矛盾。
