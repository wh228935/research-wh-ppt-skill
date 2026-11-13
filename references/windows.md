# 运行环境

核心：Python 3.10+标准库、PowerShell 7、Windows PowerPoint或WPS。无其他私人skill依赖。先运行preflight.py只探测，不安装。

PPT制作/渲染用本机Office COM；ProgID注册不等于已经实测可用。Word讲稿通过标准库OOXML生成，实际渲染仍需要可用软件。DOCX提取无需第三方库；PDF文字可用PyMuPDF或pypdf，逐页渲染需要PyMuPDF或其他已验证工具。缺依赖先给用途和安装方案，请用户决定，不自动pip或改系统。

COM用命名互斥锁串行，记录创建前PID及任务PID/start time，只关闭任务文稿。仅明确新建且无其他文稿的实例可以Quit，释放引用并核实退出；无法证实所有权不杀进程、不按名称清理。保留用户交互实例。

本机会话用户指定pwsh路径和Word生命周期要求优先。其他Windows机器识别实际PowerShell 7并报告版本，不无声回落5.1；不把本机用户名、D盘、插件缓存路径写成可迁移包硬依赖。字体问题先说明影响并经确认替换，不捆绑字体、Office或凭据。
