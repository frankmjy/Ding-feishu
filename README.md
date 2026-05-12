# DingTalk To Feishu Sync Console

本项目用于把钉钉文档/多维表导出为本地 Excel，再同步到飞书多维表。当前包含两条链路：

- 演练问题同步：默认读取 `.env`
- EA118 问题表模块：读取 `.env.aliyun_problem`，用于无附件导出钉钉问题表并同步到飞书【阿里问题登记簿】

演练链路当前配置为辅助下载：控制台负责打开钉钉目标页并等待下载；登录和导出按钮由人工点击，下载完成后自动读取 Excel 并继续同步。EA118 问题表仍保留自动下载并同步。

## 启动与停止

双击或命令行运行：

```powershell
.\start_dashboard.bat
```

控制台地址：

```text
http://127.0.0.1:8111
```

停止服务并清理安全缓存：

```powershell
.\stop_dashboard.bat
```

停止脚本只清理调试截图、候选下载文件、空日志、`__pycache__` 等临时内容；不会删除 `.env`、`.browser` 登录态、正式下载 Excel 或源代码。

## 命令行同步

```powershell
.\run_sync.ps1 --env .env.aliyun_problem --download-only
.\run_sync.ps1 --env .env.aliyun_problem --skip-download --dry-run
.\run_sync.ps1 --env .env.aliyun_problem
```

## 配置

复制样例文件并填入本机密钥：

```powershell
Copy-Item .env.example .env
Copy-Item .env.aliyun_problem.example .env.aliyun_problem
```

真实 `.env*` 文件已被 `.gitignore` 忽略，不要提交密钥、钉钉登录态、下载文件和日志。

## EA118 问题表规则

- 钉钉多维表导出走 `DINGTALK_EXPORT_KIND=bitable`
- 导出 Excel 时取消“包含所选范围中的附件”
- 同步模式使用 `insert_missing`，按 `序号` 去重
- 派生 `自动编码`：`楼栋-发现时间-风险等级-风险状态`
- `楼栋` 会把 `A/B/C...` 转为 `A楼/B楼/C楼...`
- `跟进人（飞书）` 会按 `现场跟进人` 在配置的飞书群成员中查找账号，查不到则留空
- 成功后可通过飞书应用发送同步结果和多维表链接

## 运行经验

钉钉导出和登录点位的复盘记录在 `docs/dingtalk-export-runbook.md`。后续如果钉钉 UI 再调整，优先更新这份文档和对应的视觉兜底点位。
