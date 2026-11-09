# project-reliability Specification

## Purpose
约束表格工具、工作簿操作、持久化文件和 WebView2 面板在异常、取消、切换宿主与并发场景下保持可恢复且不会误写。

## Requirements

### Requirement: Range limits remain effective
工具 SHALL 用不会在完整工作表尺寸下溢出的计数描述范围；SHALL 根据操作采用分页、分块或宿主批量调用，不按固定单元格总数拒绝有效范围。多区域 SHALL 按操作语义处理，不得只处理首个区域或将不支持的语义静默改写。

#### Scenario: Full worksheet or multiple areas
- **WHEN** 工具收到完整工作表范围或多区域并集
- **THEN** 完整范围计数不溢出，读取分页面返回，支持的多区域操作覆盖全部目标；不支持的组合在修改前解释具体原因

### Requirement: Workbook and cancellation boundaries

对话轮次、审批和撤销记录 SHALL 绑定发起时的工作簿实例。执行前目标已切换或请求已取消时 SHALL 拒绝尚未开始的操作，不修改当前的另一份工作簿。

#### Scenario: Switch workbook during approval or before undo
- **WHEN** 用户在批准前或撤销前激活另一份含同名工作表的工作簿
- **THEN** 操作被拒绝且两份工作簿均不被误写，返回原工作簿后仍可正确撤销

#### Scenario: Stop during request preparation or approval
- **WHEN** 用户在授权探测或审批完成后、实际执行前停止任务
- **THEN** 当前任务收束，禁止重复轮次和迟到写入

### Requirement: Atomic persistent storage

设置和 DPAPI 密文 SHALL 先完整写入同目录临时文件，再原子替换目标。失败 SHALL 保留可读取的旧版本，不留下半截目标文件。

#### Scenario: Commit is interrupted
- **WHEN** 写入完成前异常或替换被文件锁拒绝
- **THEN** 原目标仍完整可读，后续正常保存可以成功

### Requirement: Trusted panel lifecycle

面板 SHALL 仅允许本地 HTTPS 来源进行顶层导航和调用消息桥；关闭面板 SHALL 停止其后台任务并阻止随后访问宿主对象。

#### Scenario: External navigation or message
- **WHEN** 外部地址尝试占据面板或调用消息桥
- **THEN** 导航或消息被拒绝，本地面板仍可使用

#### Scenario: Close during initialization or queued work
- **WHEN** 面板初始化或 UI 调用尚未完成时控件被销毁
- **THEN** 后续回调不再执行宿主操作，等待任务可正常收束

### Requirement: Account refresh is bounded

账号管理后的自动刷新 SHALL 有界、可取消并合并短时间内重复的焦点事件；授权响应不明确时 SHALL 不调用可能退出当前账号的 authenticate。

#### Scenario: No focus event after account management
- **WHEN** 官方客户端账号管理已打开但 WebView 未收到新的 focus 事件
- **THEN** 面板在有限的刷新窗口内重新检查授权，切换接入方式后停止旧模式刷新

#### Scenario: Malformed user information
- **WHEN** getUserInfo 返回不符合官方格式的值
- **THEN** 明确提示无法确认账号，不当作首次登录执行 authenticate

### Requirement: One command verifies the whole project

The project SHALL provide a single command that builds the solution, runs every .NET test entry point and runs every web test, and that fails as a whole when any step fails.

The command SHALL report a summary stating how many steps passed, failed and were skipped, SHALL list each failure with the location of its log, and SHALL count web test files from the directory rather than from a hard-coded number.

Steps that require a real Office host SHALL be skippable through an explicit switch, and a skipped step SHALL be reported as skipped rather than as passed.

#### Scenario: A step fails

- **WHEN** any build step, test entry point or web test fails
- **THEN** the command exits non-zero
- **AND THEN** the summary lists that step and points to its log

#### Scenario: Host tests skipped

- **WHEN** the command is run with the skip-host switch
- **THEN** host-dependent steps are reported as skipped
- **AND THEN** the remaining steps still run and the exit code reflects only them

### Requirement: Repository-root runtime residue stays out of version control

Runtime residue created at the repository root by child processes — directories named after unexpanded environment-variable literals, lock files written through relative paths, and files created by output redirection to a device name — SHALL be excluded from version control by explicit ignore rules, each accompanied by the reason it appears.

#### Scenario: Residue is created again

- **WHEN** a child process recreates such residue at the repository root
- **THEN** version control does not report it
- **AND THEN** a file that looks like source is still reported normally
