# 钉钉导出自动化经验固化

## 演练问题文档

- 演练链路在 `sync_tasks.json` 中配置为 `download_mode=assisted`：控制台打开钉钉目标页并等待下载事件；登录和导出按钮由人工点击，下载完成后自动保存到 `downloads/dingtalk_drill_latest.xlsx` 并继续同步飞书。
- 如果只想读取本地 Excel，可把演练链路临时改为 `download_mode=manual`，或点击“仅同步本地 Excel”。
- 如果确实要恢复全自动下载，可把演练链路改为 `download_mode=auto`，并在 `.env` 中打开 `DINGTALK_AUTO_LOGIN=true` 和 `DINGTALK_AUTO_CLICK_EXPORT=true`；登录页优先点击账号名，备用点位点击页面中右侧头像卡片区域。
- 登录后如果被带到钉钉文档首页/最近列表，则按 `DINGTALK_DOC_TITLE` 点击目标文档。
- 演练链路使用独立本地文件 `downloads/dingtalk_drill_latest.xlsx`，避免和 EA118 问题表下载文件混用。
- 演练链路默认读取工作表 `演练问题`，按 Excel 里的 `序号` 去重，只新增缺失记录。
- 演练链路成功后也发送飞书消息，内容包括新增条数、跳过条数、Excel 行数和目标多维表链接。

## EA118 问题表多维表

- 钉钉多维表走右上角“分享”右侧三个点，再选择“导出 -> 导出为 Excel”。
- 出现高级权限提示时点击“继续下载”。
- 出现“导出为 Excel”弹窗时先取消勾选“包含所选范围中的附件”，再点击最终“导出”。
- 下载文件使用 `downloads/dingtalk_bitable_no_attachment_latest.xlsx`，同步到飞书【阿里问题登记簿】。
- EA118 链路固定使用 `SYNC_MODE=replace`：同步前删除飞书目标表原有记录，再按最新 Excel 全量新建，适合让飞书表完全跟随钉钉问题表。

## 易踩坑

- 不要把浏览器右上角三个点当成钉钉文档的三个点；导出入口是蓝色“分享”按钮右侧的钉钉文档菜单。
- 钉钉菜单有时是画布渲染，文字选择器读不到；需要保留视觉点位兜底。
- `.env` 和 `.env.aliyun_problem` 必须用 UTF-8 保存，否则中文文档标题、工作表名和字段名会匹配失败。
- 飞书报 `131005 not found` 时，通常不是 Excel 问题，而是当前飞书应用看不到目标 wiki/Base。处理方式是把应用加入目标 Base/wiki 的协作者，或在 `.env` 中直接填写 `FEISHU_BITABLE_APP_TOKEN`。
- 停止脚本只清理候选下载、空日志、调试截图和 Python 缓存，不删除 `.env`、正式 Excel、登录态或源代码。
