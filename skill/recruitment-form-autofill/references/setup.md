# 本地安装与初始化

需要 Python 3.10 或更新版本。网页填写还需要可用的浏览器控制工具，以及已安装的 `profile-use` companion；本 Skill 不会自动安装 companion，也不会自行启动浏览器或提交申请。

把下载的 `skill/recruitment-form-autofill` 整个文件夹放入你的 Codex Skills 目录，保留其中的 `scripts`、`references`、`assets` 和 `requirements.txt`。如果同名 Skill 已存在，先备份并核对，不要直接覆盖个人版本。

以下是 Windows PowerShell 示例。`$skillRoot` 指向实际安装目录，`$localRoot` 是你选定的私人资料目录，必须在仓库和 Skill 目录之外。

```powershell
$skillRoot = 'C:\Users\你的用户名\.codex\skills\recruitment-form-autofill'
$localRoot = 'C:\Users\你的用户名\Documents\recruitment-local'
python -m venv "$localRoot\.venv"
$profilePython = "$localRoot\.venv\Scripts\python.exe"
& $profilePython -m pip install -r "$skillRoot\requirements.txt"
$profileConfig = "$localRoot\profile-config.json"
& $profilePython "$skillRoot\scripts\init_profile.py" --workbook "$localRoot\my-profile.xlsx" --config $profileConfig
& $profilePython "$skillRoot\scripts\validate_profile.py" --config $profileConfig
& $profilePython "$skillRoot\scripts\read_profile.py" --config $profileConfig inventory
```

在 Excel 的“个人资料”第 2 行填写单值信息；“经历清单”一行一条经历，CET-4、CET-6、IELTS 各一行；附件路径使用绝对路径或相对于资料库目录的路径。保留固定表头，不增加第二个人资料行。日期使用 Excel 日期、`YYYY-MM` 或 `YYYY-MM-DD`，经历结束时间可填“至今”。编辑后保存 Excel，再运行校验。

初始化保留已有资料库，不覆盖其内容。读取退出码 `0` 为成功、`2` 为错误、`3` 为请求值为空；校验退出码 `0` 为无问题、`1` 为发现问题、`2` 为配置错误。

迁移旧版资料库使用同一个解释器：

```powershell
& $profilePython "$skillRoot\scripts\migrate_profile.py" --source '<旧资料库.xlsx>' --destination '<尚不存在的新资料库.xlsx>'
```

迁移保留源文件，目标必须是新路径。成功返回 `0` 并输出不含原始值的 `warnings`；错误返回 `2`。未映射字段需要自行补入可支持的资料字段或保留在原资料库中。迁移后把本地配置中的 `workbook_path` 更新为新路径，再校验。

只有你明确要求“帮我填写”时才使用资料填写网页；保存草稿、上传和正式提交仍分别需要你的指令。
