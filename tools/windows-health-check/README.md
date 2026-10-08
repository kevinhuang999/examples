# 工具：Windows 电脑体检与安全优化脚本

[中文](#中文) | [English](#english)

---

<a id="中文"></a>
# 中文

> **一句话**：双击运行，一次查完 CPU、内存、磁盘、网络、启动项、防火墙，并顺手做一遍可逆的清理与优化——改注册表前自动备份，不碰任何个人文件。

配套 CSDN 文章的可运行脚本。它不是 Python 示例，而是一个独立的 Windows 批处理文件：**下载到本机就能用，不依赖 Python，也不装任何第三方软件**。

---

## 一、怎么用

1. **下载** `windows-health-check.bat`（怎么下、为什么不要复制网页内容，见第四节「编码说明」）
2. **双击运行** → 弹出 UAC 提示 → 点「是」（脚本要改启动项和系统设置，需要管理员权限）
3. 在主菜单**按对应数字键**即可（按键即响应，不用按回车）

| 选项 | 作用 | 是否改动系统 |
| --- | --- | --- |
| `1` | 一键安全优化：清理缓存 + 刷新 DNS + 关闭广告推荐 + SSD TRIM + 出体检报告 | 是（全部可逆） |
| `2` | 只做体检报告 | **否，纯只读** |
| `3` | 磁盘垃圾清理（含回收站，单独二次确认） | 是 |
| `4` | 网络诊断与优化（刷新 DNS / 优化 TCP / 换公共 DNS / 重型重置） | 可选 |
| `5` | 开机启动项检查（列出全部自启 + 启用禁用状态 + 禁用残留项） | 可选 |
| `6` | 系统文件修复（sfc + DISM，耗时 10–30 分钟） | 是（官方修复流程） |
| `7` | 隐私与广告优化（关广告 ID、推荐内容、锁屏推广；开存储感知） | 是（可逆） |
| `8` | C 盘 SSD 维护（TRIM + 磁盘健康检测） | 是（仅垃圾回收） |

体检报告生成到桌面 `电脑体检报告.txt`，共十节：系统概况、CPU、内存、磁盘空间、物理磁盘健康、可清理缓存、启动项数量、网络、安全状态、建议。

---

## 二、安全边界（写在代码里的，不是口头承诺）

- **不删除任何个人文件**——文档、照片、下载、资料一律不动。只清理系统可以自动重建的缓存（用户临时目录、Windows 临时目录、错误报告、更新下载缓存、传递优化缓存）。
- **改注册表前自动备份**——导出到脚本同目录的 `reg_backup\`，想恢复双击 `.reg` 文件即可。
- **不可逆操作单独确认**——清空回收站、重置网络协议栈这类，会再问一次；且这两个的确认框在无操作时**默认选「否」**（超时自动跳过，不会误执行）。
- **耗时步骤提前告知**——清理前会提示"文件较多时约需 1–2 分钟"，不会让人以为卡死。
- **不修改个人设置**——不改桌面、不改输入法、不改浏览器主页、不装任何软件。

---

## 三、实测记录：这个脚本踩过哪些坑

这些坑都真实发生过，写在这里是因为**它们比脚本本身更值得参考**——自己做批处理时大概率也会撞上。

### 1. 批处理里的中文，必须 GBK 编码 + CRLF 换行

- **UTF-8 编码**：cmd 按 GBK 解释，中文全乱码。
- **只有 LF 换行**（Unix 风格）：cmd 的行解析会错乱，标签被截断（`:REPAIR` 被读成 `EPAIR`），报一堆「不是内部或外部命令」。
- **UTF-8 + `chcp 65001`**：看着是最现代的方案，实测**在长脚本上会出字节偏移错位**——部分命令行被拦腰截断，报「`化` 不是内部或外部命令」这种莫名其妙的错。所以最终没用它。

### 2. 正文里的字面量 `%` 必须写成 `%%`

脚本里有段生成报告的文字是「C 盘剩余空间保持在 15% 以上」。这一句会让 cmd 把 `% 以上…%` 当成变量名去解析，**把中间整段文本吃掉**，导致后面拼接的命令被截断、报语法错误。

批处理里 `%` 是变量定界符，字面量百分比一律写 `%%`。

### 3. 启动项的「禁用」标记放哪，32 位程序不一样

任务管理器里点「禁用」启动项，写的不是 `Run` 键本身，而是并行的一张状态表：

| 自启项所在位置 | 禁用标记写在哪 |
| --- | --- |
| `HKLM\...\CurrentVersion\Run` | `HKLM\...\Explorer\StartupApproved\Run` |
| `HKCU\...\CurrentVersion\Run` | `HKCU\...\Explorer\StartupApproved\Run` |
| `HKLM\...\WOW6432Node\...\Run`（32 位程序） | `HKLM\...\Explorer\StartupApproved\Run32` |

**注意最后一行**：32 位程序的自启项虽然挂在 `WOW6432Node` 下，但它的禁用标记在**上一层**的 `StartupApproved\Run32` 里，不在 `WOW6432Node` 路径下。写错位置会静默失败——命令返回成功，启动项照样开机自启。

标记本身是 `REG_BINARY`：首字节 `0x03` = 已禁用，`0x02` = 已启用，后面跟 8 字节时间戳。

### 4. 读启动项用 PowerShell 比 `reg query` 省事

`reg query` 的输出要靠 `for /f` 切分，中文还可能因代码页不对而乱码。改用 PowerShell 读注册表，输出格式自己控制，还能顺手把「已启用 / 已禁用」一起打出来。

### 5. ★ 确认框别用 `set /p` 等回车——中文输入法会把回车吃掉

这是**最影响体验的一个坑**。原本「确认执行？(Y/N):」用的是 `set /p`：需要输入字符再按回车。

实际使用中，中文输入法处于激活状态时按回车，回车会被输入法截获——**屏幕上 `y` 已经显示出来了，光标却在原地不动，程序在安静地等一个永远不会来的回车**，看起来就像脚本卡死了。

改法有两条，本项目两条都用了：

- **改用 `choice` 命令**：按单个键立即响应，不需要回车。
- **给每个确认加超时兜底**：`choice /c yn /n /t 20 /d y`——20 秒内没有任何按键就自动按默认值走。**安全操作默认「是」（如清理缓存），危险操作默认「否」（如清空回收站）**。这样最坏情况下脚本也能自己走完，不会停在那里。

> 结论：给中文用户写交互式批处理，`set /p` 做 Y/N 确认是不可靠的。

---

## 四、编码说明（重要）

这是一个 **GBK 编码**的批处理文件。中文版 Windows 的命令行默认按 GBK 解释文本，而 UTF-8 的中文在 cmd 里会出现行被截断的解析错误（原因见第三节第 1 条）。

因此：

- **网页上直接预览会显示为乱码，这是正常的**，不代表文件损坏；
- **请下载原始文件使用**（页面上「克隆/下载」→ 下载该文件，或直接取 raw 链接）；
- **不要**把网页上显示的内容复制出来另存——存成 UTF-8 后运行会出错。

---

## 五、验证环境与已知限制

| 项 | 说明 |
| --- | --- |
| 验证系统 | Windows 11 专业版（Build 26300），管理员权限窗口 |
| 依赖 | 无。只用 Windows 自带命令（`choice` / `del` / `netsh` / `DISM` / `sfc` / `defrag`）与系统自带 PowerShell 5.1 |
| 不适用 | Windows 7 及更早版本（`choice` 行为、`Get-PhysicalDisk` 等 cmdlet 不可用）；Windows 家庭版部分策略项可能被忽略 |
| 杀软提示 | 脚本会读写注册表与启动项，个别安全软件可能弹窗询问，选择「允许」即可 |

---

## 六、排错表

| 现象 | 原因 | 怎么办 |
| --- | --- | --- |
| 打开就乱码，全是方块 | 文件被当成 UTF-8 存过 | 重新下载原始文件，别复制网页内容 |
| 提示「不是内部或外部命令」且夹杂半句话 | 文件换行被改成 LF，或编码被转过 | 重新下载，别用编辑器另存 |
| 双击后一闪而过 | 没有以管理员身份运行，或被杀软拦截 | 右键 →「以管理员身份运行」；看安全软件拦截记录 |
| 停在某个确认框不动 | 中文输入法截获了回车 | 按一下 `Y` 键即可（不必回车）；或等 20 秒自动继续 |
| 清理后可用空间没明显变化 | NTFS 删除大文件后空间释放会延迟约 20 秒 | 稍等再刷新「此电脑」查看 |
| 某项显示「操作失败，可能该项不存在」 | 该启动项本机没有，属正常 | 无需处理 |
| 防火墙/杀软报警 | 脚本会修改启动项与注册表 | 选择「允许」，或先看第二节的安全边界说明 |

---

## 七、与文章的对应关系

脚本里的每一段都对应文章里的一个步骤。想改脚本的，建议先读文章里的「为什么这么写」——几个关键决策（为什么清理前先算一遍大小、为什么危险项默认否、为什么写 `Run32`）在文章里有说明。

---

<a id="english"></a>
# English

> **In one sentence**: double-click to run — it checks CPU, memory, disk, network, startup items, and firewall in a single pass, and performs a reversible cleanup and optimization along the way. The registry is backed up before any change, and no personal files are touched.

Runnable script accompanying a CSDN article. It isn't a Python example but a standalone Windows batch file: **download it and it works — no Python, no third-party software**.

---

## 1. How to Use

1. **Download** `windows-health-check.bat` (see section 4, "Encoding Notes", for why you must not copy the web page contents)
2. **Double-click to run** → a UAC prompt appears → click "Yes" (the script modifies startup items and system settings, so administrator rights are required)
3. In the main menu, **press the corresponding number key** (keys respond immediately, no Enter needed)

| Option | Purpose | Modifies the system? |
| --- | --- | --- |
| `1` | One-click safe optimization: clean caches + flush DNS + disable ad recommendations + SSD TRIM + produce a health report | Yes (all reversible) |
| `2` | Health report only | **No, strictly read-only** |
| `3` | Disk junk cleanup (includes recycle bin, with a separate second confirmation) | Yes |
| `4` | Network diagnostics and optimization (flush DNS / tune TCP / switch public DNS / heavy reset) | Optional |
| `5` | Startup item inspection (lists all autostart entries + enabled/disabled state + disabled leftovers) | Optional |
| `6` | System file repair (sfc + DISM, takes 10–30 minutes) | Yes (official repair procedure) |
| `7` | Privacy and ad optimization (turn off ad ID, recommended content, lock screen promotions; enable Storage Sense) | Yes (reversible) |
| `8` | C: drive SSD maintenance (TRIM + disk health check) | Yes (garbage collection only) |

The health report is written to the desktop as `电脑体检报告.txt`, in ten sections: system overview, CPU, memory, disk space, physical disk health, cleanable caches, startup item count, network, security status, and recommendations.

---

## 2. Safety Boundaries (enforced in code, not just promised)

- **Never deletes any personal file** — documents, photos, downloads, and data are untouched. Only caches the system can rebuild are cleaned (user temp, Windows temp, error reports, update download cache, Delivery Optimization cache).
- **Registry backed up automatically before changes** — exported to `reg_backup\` next to the script; to restore, just double-click the `.reg` file.
- **Irreversible operations confirmed separately** — emptying the recycle bin or resetting the network stack asks again; both confirmation dialogs **default to "No"** when idle (they time out and skip rather than execute by accident).
- **Slow steps announced in advance** — cleanup warns that it "may take 1–2 minutes with many files", so no one assumes it has hung.
- **Never modifies personal settings** — no changes to the desktop, IME, browser homepage, and no software installed.

---

## 3. Field Notes: Pitfalls This Script Hit

These pitfalls really happened. They're documented here because **they're more valuable than the script itself** — anyone writing batch scripts will likely run into them too.

### 1. Chinese text in a batch file requires GBK encoding + CRLF line endings

- **UTF-8 encoding**: cmd interprets it as GBK, and all Chinese text turns into mojibake.
- **LF-only line endings** (Unix style): cmd's line parsing breaks down, labels get truncated (`:REPAIR` read as `EPAIR`), and you get a flood of "not recognized as an internal or external command" errors.
- **UTF-8 + `chcp 65001`**: looks like the most modern approach, but in testing it **caused byte-offset misalignment in long scripts** — some command lines were cut in half, producing bizarre errors like "`化` is not recognized as an internal or external command". So it was ultimately not used.

### 2. Literal `%` in body text must be written as `%%`

One report-generating line reads "Keep C: drive free space above 15%". This single line makes cmd treat `% above…%` as a variable name, **swallowing the entire text in between** and truncating the commands concatenated after it, producing syntax errors.

In batch files, `%` is the variable delimiter — always write literal percent signs as `%%`.

### 3. Where the "disabled" marker for startup items lives differs for 32-bit programs

Clicking "Disable" on a startup item in Task Manager doesn't write to the `Run` key itself, but to a parallel state table:

| Location of the autostart entry | Where the disable marker goes |
| --- | --- |
| `HKLM\...\CurrentVersion\Run` | `HKLM\...\Explorer\StartupApproved\Run` |
| `HKCU\...\CurrentVersion\Run` | `HKCU\...\Explorer\StartupApproved\Run` |
| `HKLM\...\WOW6432Node\...\Run` (32-bit programs) | `HKLM\...\Explorer\StartupApproved\Run32` |

**Note the last row**: although 32-bit autostart entries live under `WOW6432Node`, their disable marker goes in the **parent-level** `StartupApproved\Run32`, not under the `WOW6432Node` path. Writing to the wrong location fails silently — the command returns success while the item still launches at boot.

The marker itself is `REG_BINARY`: first byte `0x03` = disabled, `0x02` = enabled, followed by an 8-byte timestamp.

### 4. Using PowerShell to read startup items beats `reg query`

`reg query` output has to be parsed with `for /f`, and the Chinese text may be mojibake if the code page is wrong. Reading the registry with PowerShell gives you control over the output format and lets you print "enabled / disabled" alongside each entry.

### 5. ★ Don't use `set /p` for confirmation prompts — a Chinese IME will swallow the Enter key

This is **the single most disruptive pitfall**. The original "Execute? (Y/N):" prompt used `set /p`, which requires typing a character and pressing Enter.

In real use, when a Chinese input method is active, pressing Enter gets intercepted by the IME — **the `y` shows up on screen but the cursor doesn't move, and the program sits quietly waiting for an Enter that will never come**, looking exactly like a hung script.

There are two fixes, and this project uses both:

- **Switch to the `choice` command**: a single keypress responds immediately, no Enter required.
- **Add a timeout fallback to every confirmation**: `choice /c yn /n /t 20 /d y` — if no key is pressed within 20 seconds, it proceeds with the default. **Safe operations default to "Yes" (e.g. cache cleanup); dangerous ones default to "No" (e.g. emptying the recycle bin)**. Worst case, the script finishes on its own instead of stalling.

> Conclusion: for interactive batch scripts aimed at Chinese users, `set /p` is an unreliable way to do Y/N confirmation.

---

## 4. Encoding Notes (Important)

This is a **GBK-encoded** batch file. Command lines on Chinese versions of Windows interpret text as GBK by default, while UTF-8 Chinese causes truncated-line parse errors in cmd (see section 3, item 1).

Therefore:

- **Viewing it directly on the web page shows mojibake — this is normal** and does not mean the file is corrupted.
- **Please download the original file to use it** (click "Clone/Download" → download the file, or use the raw link).
- **Do not** copy the text shown on the web page and save it yourself — saving as UTF-8 will cause errors when run.

---

## 5. Verified Environment and Known Limitations

| Item | Notes |
| --- | --- |
| Verified on | Windows 11 Pro (Build 26300), administrator console |
| Dependencies | None. Uses only built-in Windows commands (`choice` / `del` / `netsh` / `DISM` / `sfc` / `defrag`) and the bundled PowerShell 5.1 |
| Not applicable to | Windows 7 and earlier (`choice` behavior and cmdlets like `Get-PhysicalDisk` are unavailable); some policy items may be ignored on Windows Home |
| Antivirus prompts | The script reads and writes the registry and startup items; some security software may prompt — choose "Allow" |

---

## 6. Troubleshooting

| Symptom | Cause | What to do |
| --- | --- | --- |
| Mojibake on open, all boxes | The file was saved as UTF-8 at some point | Re-download the original; don't copy from the web page |
| "Not recognized as an internal or external command" with half-sentences | Line endings were converted to LF, or the encoding was converted | Re-download; don't re-save in an editor |
| A flash and it's gone after double-clicking | Not run as administrator, or blocked by antivirus | Right-click → "Run as administrator"; check your security software's block log |
| Stuck at a confirmation prompt | The Chinese IME intercepted Enter | Press `Y` once (no Enter needed), or wait 20 seconds for it to continue automatically |
| No noticeable free space after cleanup | NTFS delays space release by about 20 seconds after deleting large files | Wait a moment and refresh "This PC" |
| An item reports "operation failed, item may not exist" | That startup item isn't present on this machine — normal | No action needed |
| Firewall/antivirus alert | The script modifies startup items and the registry | Choose "Allow", or read the safety boundaries in section 2 first |

---

## 7. Relationship to the Article

Every section of the script corresponds to a step in the article. If you want to modify the script, read the article's "why it's written this way" first — several key decisions (why sizes are computed before cleaning, why dangerous items default to No, why `Run32` must be written) are explained there.
