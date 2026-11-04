# 实测 RED/GREEN 记录

- RED 1：写入首批行为测试后运行 `.venv\Scripts\python.exe -m pytest -q`，测试收集失败：`ModuleNotFoundError: No module named 'app'`。
- GREEN 1：实现路由、存储和基础合并后，首轮 3 项失败、1 项通过。失败均由合法 PPTX 自带 `ppt/printerSettings/printerSettings1.bin` 被宏校验误拦截引起。调整为只拦截 VBA/ActiveX 后，`4 passed in 5.38s`。
- RED 2：新增真实文件合并测试后，`1 failed, 5 passed`；失败是测试把空白占位符当成第一个文本框。改为检查非空文本框后，`6 passed in 9.26s`。
- GREEN 2：新增重启持久化与历史顺序测试后，`8 passed in 17.51s`。
- GREEN 3：增加上传 CSRF、整日删除和空 ZIP 检查后，全量运行 `.venv\Scripts\python.exe -m pytest -q`，结果 `9 passed in 7.63s`。
- HTTP 冒烟：使用合成临时凭据启动真实 Waitress 进程，访问 `/` 和 `/admin/login`，得到 `home=200 chinese=True login=200`，随后终止进程。
- Docker：当前环境没有 `docker` 命令，镜像构建和 Compose 部署未验证。

这些结果仅对应上述本地测试，不代表 Docker 已验证。

## 2026-11-04 PowerPoint 损坏修复

- 依据 [Microsoft 演示文稿结构文档](https://learn.microsoft.com/en-us/office/open-xml/presentation/structure-of-a-presentationml-document)、[幻灯片母版 ID 示例](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.slidemasterid) 和 [幻灯片 ID 范围](https://learn.microsoft.com/en-us/openspecs/office_protocols/ms-pwvpdp/f18e6e1d-8b23-4662-8cc6-86be5248e7d2) 核对了演示文稿子元素顺序、ID 范围及每个包最多一个备注母版的限制。旧实现把第二份母版分配成 `257`（应在 `2147483648` 以上），并会复制第二套备注母版；`ElementTree` 还丢失 `mc:Ignorable` 引用的命名空间声明。
- 新增最小回归用例后，修复前 `tests/test_merge.py` 为 `2 failed, 2 passed`，失败点分别是母版 ID 范围及 `mc` 命名空间声明。修复后该文件的首轮运行 `4 passed`。随后补充绝对关系目标、未引用部件、冲突备注母版用例，全量运行 `.venv\Scripts\python.exe -m pytest -q -p no:cacheprovider --basetemp .test-tmp\suite` 为 `14 passed in 21.48s`。此前一次全量运行因 `.test-tmp` 父目录尚未建立产生 8 个测试环境错误，创建该目录后重跑成功。
- 合成富内容合并文件通过 `pptx/scripts/office/validate.py`（使用首份合成文件作 `--original` 基线）：`All validations PASSED!`。该结果覆盖关系、内容类型、XML 架构等静态检查，不等同于 PowerPoint 打开。
- 当日 3 份原件各自通过同一校验器的静态检查，均为 `All validations PASSED!`。
- 使用只读 SQLite URI 读取 `20261104` 的提交顺序，只读加载 3 份上传文件。前两份均包含备注页和不同的备注母版；按 OOXML 限制无法无损保留两套备注母版，当前合并明确抛出 `notesMasterIdLst 与首份演示文稿不兼容，无法无损合并`，未写出当日合并文件，也未改动原件。
- 发现 `POWERPNT.EXE`，且探测前没有运行中的 PowerPoint。新实例 COM 创建在当前登录会话失败，HRESULT `0x80070520`；未进入文件打开步骤，未产生可声称的 PowerPoint 兼容性验证。探测后仍无 PowerPoint 进程。

## 备注母版显式统一修复

- 先添加冲突备注母版默认拒绝、勾选后合并、备注文字及图片保留、单个备注母版、首份没有备注母版时从后续文件导入，以及管理页面 POST 参数转发测试。实现前运行 `tests/test_merge.py tests/test_app.py` 得到 `5 failed, 13 passed`；其中两项新夹具本身的设置有误，修正后再实施功能。
- 管理页新增默认未勾选的“统一备注母版（保留备注内容，备注排版可能变化）”。勾选时后续备注页保留自身 XML 和关系，备注母版关系指向顺序中第一份可用的备注母版；未勾选时冲突仍拒绝并提示勾选。讲义母版冲突仍拒绝。补全了来源只有备注母版关系、没有 ID 列表时的导入路径。
- 全量测试：`.venv\Scripts\python.exe -m pytest -q -p no:cacheprovider --basetemp .test-tmp\notes-suite`，`20 passed in 18.31s`。
- 只读查询 `data/hub.sqlite3` 并只读加载 `20261104` 的三份原件，按记录顺序合并，写出被忽略的 `.test-tmp/fixed/20261104-fixed.pptx`。三份均为 1 页；输出 3 页，两份原有备注正文逐项一致，所有可见幻灯片的文本、形状类型及顺序逐项一致。以首份原件作 `--original` 基线运行 `pptx/scripts/office/validate.py`，结果 `All validations PASSED!`。另用含图片、原生图表、备注的冲突母版夹具合并，基于原始夹具的同一校验器结果也为 `All validations PASSED!`。这些是静态检查；未实际用 PowerPoint 打开新文件，不声称 PowerPoint 接受或备注排版保真。
