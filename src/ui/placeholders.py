from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QScrollArea,
    QStackedWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)


def build_placeholder_tab(message: str) -> QWidget:
    widget = QWidget()
    layout = QVBoxLayout(widget)
    layout.setContentsMargins(40, 40, 40, 40)
    layout.setAlignment(Qt.AlignCenter)
    card = QFrame()
    card.setObjectName("placeholderCard")
    card_layout = QVBoxLayout(card)
    card_layout.setContentsMargins(50, 40, 50, 40)
    card_layout.setSpacing(16)
    card_layout.setAlignment(Qt.AlignCenter)
    icon = QLabel("🛠️")
    icon.setObjectName("placeholderIcon")
    icon.setAlignment(Qt.AlignCenter)
    title = QLabel("功能建设中")
    title.setObjectName("placeholderTitle")
    title.setAlignment(Qt.AlignCenter)
    label = QLabel(message)
    label.setObjectName("placeholderText")
    label.setAlignment(Qt.AlignCenter)
    label.setWordWrap(True)
    card_layout.addWidget(icon)
    card_layout.addWidget(title)
    card_layout.addWidget(label)
    layout.addWidget(card)
    return widget


class AboutTab(QWidget):
    def __init__(self) -> None:
        super().__init__()
        root = QVBoxLayout(self)
        root.setContentsMargins(24, 24, 24, 24)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        root.addWidget(scroll)

        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(12, 8, 12, 16)
        content_layout.setSpacing(16)

        header = QFrame()
        header.setObjectName("aboutHeader")
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(28, 24, 28, 24)
        header_layout.setSpacing(8)
        title = QLabel("基带测试数据统计工具 V1.5")
        title.setObjectName("aboutTitle")
        subtitle = QLabel("统一处理充电、充电保护、分压采集与续航测试数据，支持多模式统计、批量处理与图表导出。")
        subtitle.setObjectName("aboutSubtitle")
        subtitle.setWordWrap(True)
        header_layout.addWidget(title)
        header_layout.addWidget(subtitle)
        content_layout.addWidget(header)

        positioning = self._build_card(
            "🎯 工具定位",
            "本工具用于基带测试阶段的数据整理与结果归档，当前已覆盖“充电测试”“充电保护测试”“分压采集测试”与“续航测试”场景：\n"
            "• 充电测试：支持“单文件模式（Excel）”与“合并模式（Excel+CSV）”\n"
            "• 充电保护测试：支持“单文件模式（Excel）”，输出高温/低温停充与复充温度结果\n"
            "• 分压采集测试：支持“单文件模式（Excel）”，按分压电压与采样电阻推导电流并输出曲线结果\n"
            "• 续航测试：支持“单Excel续航时长统计 / 单log续航时长统计 / 电量指示统计”三类任务\n"
            "• 通用能力：支持拖拽上传、后台线程执行、实时进度/日志、分批策略与分模式输出目录",
        )
        content_layout.addWidget(positioning)

        usage = self._build_card(
            "📝 使用步骤",
            "1️⃣ 在“充电测试”“充电保护测试”“分压采集测试”或“续航测试”页面左侧【文件上传】中添加文件或文件夹（支持拖拽）\n"
            "2️⃣ 在【输出设置】中确认输出目录；如需可开启【分批策略】\n"
            "3️⃣ 如为“分压采集测试”，请先确认采样电阻阻值（默认 20 mΩ）；然后在【数据处理】中点击“统计数据”并选择对应执行项\n"
            "4️⃣ 处理过程中可在右侧【运行日志】查看进度、告警与错误\n"
            "5️⃣ 处理完成后，到输出目录查看按模式生成的 Excel 结果",
        )
        content_layout.addWidget(usage)

        precautions = self._build_card(
            "⚠️ 注意事项",
            "• 命名与配对规则（重点）：\n"
            "  - 充电“合并模式”：按同名文件配对，必须是“同主文件名的 1 个 Excel + 1 个 CSV”（例如 Case01.xlsx + Case01.csv）\n"
            "  - 续航“电量指示统计”：批量时按主文件名配对“1 个 Excel + 1 个 txt/log”；若仅上传 1 个 Excel 和 1 个文本，可自动成对，但仍建议同名\n"
            "  - 单次执行内，同一主文件名不能出现多个候选文件，否则会报“无法唯一配对”\n"
            "• 模式与文件类型需匹配：\n"
            "  - 单log续航时长统计仅支持 .txt/.log，不支持与 Excel 混合输入\n"
            "  - 充电测试单文件模式 / 充电保护测试单文件模式 / 分压采集测试单文件模式 / 单Excel续航时长统计建议仅上传 Excel，避免误触提示\n"
            "• 分压采集测试需额外确认采样电阻阻值（单位 mΩ），系统会按“分压电压 / 采样电阻”推导电流；若阻值填写错误，结果会按比例偏差\n"
            "• 上传文件夹会递归扫描子目录；重复路径会自动去重\n"
            "• 若工具运行时被杀毒软件误识别为异常程序，请设置分批策略（减小短时间内文件处理数量）或添加到白名单，以免影响统计功能正常使用",
        )
        content_layout.addWidget(precautions)

        notes = self._build_card(
            "📊 数据与结果说明",
            "• 支持输入：.xlsx/.xls、.csv、.txt/.log（具体支持范围取决于所选模式）\n"
            "• 输出目录：每次执行会在目标目录下按“模式名_时间戳”创建子目录，避免多次执行互相覆盖\n"
            "• 输出命名：结果默认沿用输入文件的主文件名；若目标目录中重名，会自动追加 (1)、(2) 等序号\n"
            "• 配对模式会做有效性校验：缺失配对、同名多文件、时间点无交集等情况会直接报错\n"
            "• 批处理策略：单个文件/文件组失败不会中断其他任务，最终会输出总计/成功/失败汇总\n"
            "• 温升统计：当检测到“笔壳温度 + 环境温度”列时，会自动追加温升相关结果\n"
            "• 充电测试-单文件模式：若缺少电流或电压，但存在完整的“笔壳温度 + 环境温度”数据，将给出警告并继续输出温升统计结果\n"
            "• 充电保护测试-单文件模式：要求存在“环境温度 + 电芯温度 + 电流 + 电压”数据，用于识别高温/低温停充区间并输出保护温度表与折线图\n"
            "• 分压采集测试-单文件模式：要求存在两列“电压 (V)”相关数据，其中绝对值更大的列作为主电压，较小的一列保留为“分压电压 (V)”并参与电流换算",
        )
        content_layout.addWidget(notes)

        info = QFrame()
        info.setObjectName("aboutCard")
        info_layout = QVBoxLayout(info)
        info_layout.setContentsMargins(28, 22, 28, 22)
        info_layout.setSpacing(10)
        info_title = QLabel("ℹ️ 版本信息")
        info_title.setObjectName("aboutSectionTitle")
        info_body = QLabel(
            "版本：V1.5\n"
            "开发人员：邓景华\n"
            "开发日期：2026-03-27"
        )
        info_body.setObjectName("aboutBody")
        info_layout.addWidget(info_title)
        info_layout.addWidget(info_body)
        content_layout.addWidget(info)
        content_layout.addStretch()

        scroll.setWidget(content)

    def _build_card(self, title_text: str, body_text: str) -> QFrame:
        card = QFrame()
        card.setObjectName("aboutCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(28, 22, 28, 22)
        layout.setSpacing(12)
        title = QLabel(title_text)
        title.setObjectName("aboutSectionTitle")
        body = QLabel(body_text)
        body.setObjectName("aboutBody")
        body.setWordWrap(True)
        layout.addWidget(title)
        layout.addWidget(body)
        return card


class UpdateLogTab(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self._entries = [
            {
                "version": "V1.5",
                "title": "分压采集测试模块上线",
                "time": "2026-03-27",
                "detail": (
                    "【版本目标】\n"
                    "新增“分压采集测试”模块，支持从双电压采样 Excel 中自动判定主电压与分压电压，并基于采样电阻换算电流后输出统计结果。\n\n"
                    "【主要更新】\n"
                    "1. 分压采集测试页面上线：\n"
                    "   - 左侧导航新增“分压采集测试”页，整体交互与现有统计页保持一致。\n"
                    "   - 通过“统计数据”弹窗触发执行，本版本提供“单文件模式（Excel）”。\n"
                    "2. 采样电阻输入新增：\n"
                    "   - 在文件上传区按钮行新增“采样电阻阻值”输入框，默认值为 20 mΩ。\n"
                    "   - 处理时按“分压电压 / 采样电阻”推导电流，并沿用既有电流极性校正规则。\n"
                    "3. 解析与导出能力新增：\n"
                    "   - 自动识别两列“电压 (V)”数据，按绝对值最大值大小区分“电压 (V)”与“分压电压 (V)”。\n"
                    "   - 复用现有充电测试的尾行校验、指标计算、图表导出和批处理能力。\n"
                    "4. 文档与版本同步更新：\n"
                    "   - 关于页补充分压采集测试说明。\n"
                    "   - 主窗口版本提升至 V1.5。\n\n"
                    "【兼容说明】\n"
                    "本版本新增独立模块，不改变既有“充电测试”“充电保护测试”与“续航测试”的处理口径。"
                ),
            },
            {
                "version": "V1.4",
                "title": "充电保护测试模块上线",
                "time": "2026-03-26",
                "detail": (
                    "【版本目标】\n"
                    "新增“充电保护测试”模块，补齐高温/低温停充与复充温度统计能力，并将模块入口接入主导航。\n\n"
                    "【主要更新】\n"
                    "1. 充电保护测试页面上线：\n"
                    "   - 左侧导航新增“充电保护测试”页，布局与现有统计页保持一致。\n"
                    "   - 仍通过“统计数据”弹窗触发执行，本版本提供“单文件模式（Excel）”。\n"
                    "2. 保护测试解析与统计能力新增：\n"
                    "   - 在复用时间/电流/电压解析与尾行校验逻辑的基础上，新增“环境温度 / 电芯温度”识别。\n"
                    "   - 按停充区间识别高温/低温保护温度与复充温度，支持虚假停充过滤与开放区间 warning。\n"
                    "3. 结果输出新增：\n"
                    "   - 输出基础数据、保护温度汇总表，以及“电流 + 环境温度 + 电芯温度 + 电压”折线图。\n"
                    "4. 文档与版本同步更新：\n"
                    "   - 关于页新增模块说明。\n"
                    "   - 主窗口与安装包版本提升至 V1.4。\n\n"
                    "【兼容说明】\n"
                    "本版本新增独立模块，不改变既有“充电测试”与“续航测试”的处理口径。"
                ),
            },
            {
                "version": "V1.3",
                "title": "充电单文件模式兼容仅温升数据输入",
                "time": "2026-03-23",
                "detail": (
                    "【版本目标】\n"
                    "提升“充电测试-单文件模式”对不完整曲线数据的兼容性，避免在仅需温升统计的场景下因缺少电流或电压而直接中断。\n\n"
                    "【主要更新】\n"
                    "1. 单文件模式输入兼容性增强：\n"
                    "   - 当 Excel 中缺少“电流”或“电压”列/有效数据时，不再立即报错终止。\n"
                    "   - 若同时存在完整的“笔壳温度 (°C)”与“环境温度 (°C)”数据，则判定为仅执行“充电温升测试”的场景。\n"
                    "2. 告警与执行策略调整：\n"
                    "   - 对上述场景新增 warning 提示，明确说明已跳过“充电曲线测试”统计，仅继续执行“充电温升测试”统计。\n"
                    "3. 结果输出联动优化：\n"
                    "   - 仅温升场景下，结果页不再生成“充电曲线测试”表格与曲线图。\n"
                    "   - 保留基础数据、温升统计表及温升图，输出结构与实际可统计内容保持一致。\n\n"
                    "【兼容说明】\n"
                    "本版本仅放宽“充电测试-单文件模式”在特定输入场景下的终止条件，不改变既有曲线统计指标口径；当温升数据也不完整时，仍按原有规则报错。"
                ),
            },
            {
                "version": "V1.2",
                "title": "单log续航时长统计增强与交互一致性优化",
                "time": "2026-03-08",
                "detail": (
                    "【版本目标】\n"
                    "增强“单log续航时长统计”在纯文本场景下的处理能力与结果可读性，并统一防误触交互体验。\n\n"
                    "【主要更新】\n"
                    "1. 单log输入规则与批处理能力增强：\n"
                    "   - 单log模式仅接受文本文件（.txt/.log），不支持与 Excel 混合输入。\n"
                    "   - 支持多文本批量处理，并接入分批策略（单批上限 + 批间等待）。\n"
                    "2. 单log结果数据与图表优化：\n"
                    "   - 对普通含时间日志，按秒补齐相邻电量记录之间的时间间隔，便于在图中观察各电量持续时长。\n"
                    "   - 单log结果页隐藏“log-日期”显示列，仅保留“log-时间 (s)/电量(%)（及映射电压）”，不影响续航时长计算与绘图。\n"
                    "3. 特殊文本处理重构：\n"
                    "   - 将特殊文本处理从“电量指示统计”分支移出，统一收敛到“单log续航时长统计”。\n"
                    "   - 特殊文本按既有规则使用 0:00:00 作为起点并按秒递增记录时间，保留全量电压/电量点位。\n"
                    "4. 防误触交互样式统一：\n"
                    "   - 单log模式的防误触提示改为统一使用自定义 SafeConfirmDialog，视觉风格与其它控件一致。\n\n"
                    "【兼容说明】\n"
                    "本版本主要增强续航-单log场景处理与交互一致性，不改变既有充电统计指标口径。"
                ),
            },
            {
                "version": "V1.1",
                "title": "电量指示统计增强与后台线程防卡顿改造",
                "time": "2026-03-07",
                "detail": (
                    "【版本目标】\n"
                    "补齐“电量指示数据统计”在续航时长输出上的需求闭环，并完成统计执行链路的后台线程化改造，解决大数据量场景下主界面卡顿问题。\n\n"
                    "【主要更新】\n"
                    "1. 电量指示统计新增续航时长计算与结果展示：\n"
                    "   - 新增按“log-日期 + log-时间 (s)”结束时间减去 Excel 首个“日期+时间”起点的续航时长计算。\n"
                    "   - 在结果页右侧新增“续航时长统计”表（测试项/数据），输出“续航时长”值。\n"
                    "2. BUG修复：特殊无时间戳日志解析规则修正：\n"
                    "   - 对 `MCU:[APP-I:bat]bat scan handle discharg vol:xxxx(mv) level:yy(%)` 类型文本，已确保按顺序读取全部电压/电量数据。\n"
                    "   - 保留 mV->V 转换，不执行“首次电量去重”规则，分支返回原始全量点位。\n"
                    "3. 统计任务后台线程化（防卡顿）：\n"
                    "   - 充电测试与续航测试页的统计流程均改为 QThread + Worker 异步执行，避免主线程被长耗时任务阻塞。\n"
                    "   - 日志与进度通过 Signal/Slot 回传主线程更新，处理期间主界面可继续进行切页、拖动、查看日志等操作。\n"
                    "   - 新增任务运行态保护：统计进行中禁止重复触发同页任务；程序退出时若仍有任务执行，提示等待完成后再关闭。\n\n"

                    "【兼容说明】\n"
                    "本版本不改变既有统计指标口径与输出结构，主要增强续航-电量指示流程能力，并提升大批量处理时的界面响应性与运行稳定性。"
                ),
            },
            {
                "version": "V1.0",
                "title": "正式版发布：续航测试功能上线",
                "time": "2026-03-05",
                "detail": (
                    "【版本目标】\n"
                    "发布 V1.0 正式版本，补齐续航测试主流程，使充电与续航两大场景均可在同一工具内完成统计。\n\n"
                    "【主要更新】\n"
                    "1. 续航测试功能正式上线：\n"
                    "   - 新增“执行续航时长统计”，用于 Excel（.xlsx/.xls）续航时长数据统计与图表输出。\n"
                    "   - 新增“执行电量指示统计”，用于 Excel（.xlsx/.xls）与文本（.txt/.log）配对数据统计。\n"
                    "2. 续航页数据处理入口与交互对齐充电页：\n"
                    "   - 统一使用“统计数据”按钮，在弹窗中选择执行项，降低学习与操作成本。\n"
                    "   - 统一日志输出格式、状态提示与防误触确认流程，便于批处理排障。\n"
                    "3. 结果输出组织延续既有规则：\n"
                    "   - 按处理模式自动创建时间戳子目录，避免多批次结果互相覆盖。\n"
                    "   - 文件命名冲突时自动追加序号，保持历史结果可追溯。\n\n"
                    "【兼容说明】\n"
                    "V1.0 在保留充电测试既有能力与输出口径的基础上新增续航测试能力，不影响历史数据处理流程。"
                ),
            },
            {
                "version": "V0.11",
                "title": "窗口自适应增强、输出目录分层与图表标题规则修正",
                "time": "2026-03-04",
                "detail": (
                    "【版本目标】\n"
                    "提升窗口自适应能力，规范批处理输出目录结构，并修正图表标题命名回退规则。\n\n"
                    "【主要更新】\n"
                    "1. 高 DPI / 缩放场景适配增强：\n"
                    "   - 首次显示后基于窗口外框真实尺寸（含标题栏/边框）与可用工作区二次校正。\n"
                    "2. 输出目录结构优化：\n"
                    "   - 单文件模式与合并模式执行时，不再直接写入 output 根目录。\n"
                    "   - 改为自动创建“模式名_时间戳（精确到秒）”子目录后再输出结果。\n"
                    "3. 图表标题命名规则修正：\n"
                    "   - 当文件名中既不包含“充电曲线”也不包含“充电温升”时，曲线图与温升图标题均改为 Excel 文件名。\n"
                    "4. 连续执行稳定性修复：\n"
                    "   - 修复首次统计成功、再次统计偶发失败的问题（报错为 openpyxl 临时文件目录不存在）。\n"
                    "   - 临时目录切换由环境变量方式改为显式设置/恢复 tempfile 缓存目录，避免残留旧路径。\n"
                    "【兼容说明】\n"
                    "本版本不改变统计指标计算口径，主要增强界面适配与输出组织方式，并修正标题命名细则。"
                ),
            },
            {
                "version": "V0.10",
                "title": "处理入口重构、误触防护增强与进度可视化",
                "time": "2026-03-04",
                "detail": (
                    "【版本目标】\n"
                    "统一“数据处理”入口交互，提升模式选择清晰度与防误操作能力，并增强处理过程可视化反馈。\n\n"
                    "【主要更新】\n"
                    "1. 数据处理入口重构：\n"
                    "   - 原“统计数据 + 合并后统计数据”双按钮改为单一“统计数据”入口。\n"
                    "   - 点击后弹出模式选择弹窗，提供“单文件模式 / 合并模式”二选一，并附简短功能说明。\n"
                    "2. 弹窗界面美化：\n"
                    "   - 新增模式选择弹窗视觉样式，强化标题区、模式卡片区与操作区层次。\n"
                    "   - 防误触弹窗改为统一风格的自定义提醒框，优化提示可读性与按钮层级。\n"
                    "3. 合并模式防误触逻辑增强：\n"
                    "   - 保留“未检测到 .csv 文件”提醒。\n"
                    "   - 新增“Excel 与 CSV 数量不一致”提醒（例如混入单文件模式输入），需二次确认后继续执行。\n"
                    "4. 实时进度可视化：\n"
                    "   - 在“统计数据”按钮右侧新增处理进度条，实时展示百分比与已完成/总量。\n"
                    "   - 处理过程中按单项成功/失败实时推进，完成后自动显示 100% 并切换完成态样式。\n\n"
                    "【兼容说明】\n"
                    "本版本不改变统计指标口径与输出结构，主要更新 UI 交互、防误触策略与处理进度展示。"
                ),
            },
            {
                "version": "V0.9",
                "title": "分批策略上线与日志去重优化",
                "time": "2026-03-03",
                "detail": (
                    "【版本目标】\n"
                    "提升在部分终端安全软件环境下的批处理稳定性，并减少运行日志重复信息，提升可读性。\n\n"
                    "【主要更新】\n"
                    "1. 新增“分批策略”能力（输出设置区）：\n"
                    "   - 新增“分批策略...”入口，可配置“单批上限 + 批间等待秒数”。\n"
                    "   - 默认关闭，开启后按批次执行：处理 N 个文件（或文件组）后等待 X 秒再继续。\n"
                    "   - 支持设置持久化，重启软件后自动保留上次配置。\n"
                    "2. 分批策略弹窗交互优化：\n"
                    "   - 开关升级为滑块样式，状态展示统一为“已开启/已关闭”。\n"
                    "   - 参数区改为同一行展示：单批上限 [输入框] 个文件，批间等待 [输入框] 秒。\n"
                    "   - 数值输入框移除上下箭头按钮，尺寸与字体按易读性优化。\n"
                    "3. 批处理执行链路接入节奏控制：\n"
                    "   - “统计数据”与“合并后统计数据”两条流程均支持分批等待执行。\n"
                    "   - 批次切换时输出进度日志，便于追踪当前处理状态。\n"
                    "4. 日志去重与标识统一：\n"
                    "   - 成功日志统一为“[成功] 输出成功：...”，失败日志统一为“[失败] ...处理失败：...”。\n"
                    "   - 汇总区移除逐条成功/失败明细，避免与过程日志重复。\n\n"
                    "【兼容说明】\n"
                    "本版本不改变统计指标计算口径，主要增强批处理稳定性与日志/界面体验。"
                ),
            },
            {
                "version": "V0.8",
                "title": "O.L 自动替换与 Unit 驱动的电流/电压补齐",
                "time": "2026-03-02",
                "detail": (
                    "【版本目标】\n"
                    "补齐“相反数据来源”场景下的细化规则，解决 O.L 异常值与单位识别问题，确保合并后结果稳定。\n\n"
                    "【主要更新】\n"
                    "1. 新增 O.L 自动替换规则：\n"
                    "   - Excel 电流列、电压列：若检测到 O.L，使用后一个值替换后再参与计算（已删除）。\n"
                    "   - CSV Value 列：时间交集筛选后，若检测到 O.L，同样使用后一个值替换。\n"
                    "2. CSV Unit 驱动映射策略：\n"
                    "   - Unit=V：CSV 作为电压列补入，Excel 需提供电流列。\n"
                    "   - Unit=mA：CSV 作为电流列补入，Excel 需提供电压列。\n"
                    "   - Unit=A：CSV 先按 A->mA（乘以1000）后作为电流列补入，Excel 需提供电压列。\n"
                    "3. 冲突校验与错误提示完善：\n"
                    "   - 新增来源冲突校验（如 Unit 与 Excel 已有列角色不匹配）并给出明确错误信息。\n"
                    "4. 日志增强：\n"
                    "   - 新增 CSV/Excel O.L 替换数量提示，便于追溯数据修复情况。\n\n"
                    "【兼容说明】\n"
                    "本版本在 V0.7 基础上增强异常值处理与单位识别，不改变核心统计指标定义。"
                ),
            },
            {
                "version": "V0.7",
                "title": "重复秒级时间点修正与交集匹配",
                "time": "2026-03-02",
                "detail": (
                    "【版本目标】\n"
                    "提升“合并后统计数据”在时间异常场景下的健壮性，减少因时间点微小偏差导致的失败。\n\n"
                    "【主要更新】\n"
                    "1. 重复秒级时间点处理增强：\n"
                    "   - 当重复秒前后存在缺失 1 秒时，自动修正前/后重复点为缺失秒。\n"
                    "   - 无缺失秒场景下，保持“保留首次出现数据”的原有策略。\n"
                    "2. 匹配规则调整：\n"
                    "   - Excel 与 CSV 时间匹配由“Excel 全量必须匹配”改为“取两者时间点交集”。\n"
                    "3. 合并来源灵活化：\n"
                    "   - 支持 Excel 与 CSV 的电流/电压来源不固定，可由任一侧提供对应列数据。\n"
                    "4. 日志与文档同步：\n"
                    "   - 增加重复秒修正与交集过滤相关提示，需求文档 1.2.2 同步更新。\n\n"
                    "【兼容说明】\n"
                    "本版本主要调整时间匹配与来源判定规则，不改变统计结果计算口径。"
                ),
            },
            {
                "version": "V0.6",
                "title": "防误触提醒与 .xls/.xlsx 同等支持",
                "time": "2026-03-01",
                "detail": (
                    "【版本目标】\n"
                    "降低误操作风险，并统一 Excel 输入兼容性，确保 .xls 与 .xlsx 在充电测试流程中同等可用。\n\n"
                    "【主要更新】\n"
                    "1. 新增“防误触”二次确认弹窗：\n"
                    "   - 点击“统计数据”时，若检测到上传内容包含 .csv，将提示可能误操作并要求确认。\n"
                    "   - 点击“合并后统计数据”时，若未检测到 .csv，将提示可能误操作并要求确认。\n"
                    "2. Excel 输入扩展为 .xlsx/.xls 同等处理：\n"
                    "   - 统计数据与合并后统计数据流程均支持读取 .xlsx 与 .xls。\n"
                    "   - 合并流程支持“同名 Excel（.xlsx/.xls）+ .csv”配对。\n"
                    "3. 兼容异常后缀场景：\n"
                    "   - 针对“文件后缀是 .xls，但实际内容是 xlsx”的情况，新增自动识别并按 xlsx 解析，避免报错中断。\n"
                    "4. 界面与文案同步：\n"
                    "   - 文件上传提示、选择器过滤条件、操作提示语同步更新为 .xlsx/.xls。\n"
                    "5. 文档同步：\n"
                    "   - 需求文档中 1.1 与 1.2.2 相关描述已同步更新，保持与实现一致。\n\n"
                    "【兼容说明】\n"
                    "本版本不改变统计指标计算口径，仅增强输入兼容性与操作防呆提示。"
                ),
            },
            {
                "version": "V0.5",
                "title": "图表兼容性修复与展示优化",
                "time": "2026-02-28",
                "detail": (
                    "【版本目标】\n"
                    "修复同一 Excel 在 WPS 与 Office 中图表显示不一致问题，提升图表稳定性与可读性。\n\n"
                    "【主要更新】\n"
                    "1. 充电曲线图兼容性修复：\n"
                    "   - 时间列改为原生 datetime 写入，并统一设置 hh:mm:ss 显示格式。\n"
                    "   - 双坐标轴图显式绑定轴 ID 与交叉轴，修复 Office 中轴映射异常。\n"
                    "2. 绘图区布局优化：\n"
                    "   - 引入手动绘图区布局参数，仅缩放绘图区，不改变图表外框尺寸。\n"
                    "   - 调整绘图区位置与留白，避免标题、图例与坐标轴标签相互遮挡。\n"
                    "3. 温升图坐标轴修复：\n"
                    "   - 显式设置温升图 X/Y 轴位置与标签显示，并补齐时间类目轴绑定。\n"
                    "   - 修复温升图在 Office 中偶发缺少坐标轴显示的问题。\n"
                    "4. 图表标题样式增强：\n"
                    "   - 新增标题字号配置，默认统一提升并加粗。\n\n"
                    "【兼容性说明】\n"
                    "本版本主要涉及 Excel 图表生成与样式配置，不影响现有统计计算逻辑与批处理入口。"
                ),
            },
            {
                "version": "V0.4",
                "title": "规则修正与输出格式统一",
                "time": "2026-02-27",
                "detail": (
                    "【版本目标】\n"
                    "修正需求理解偏差导致的解析逻辑问题，并统一输出格式与日志提示，保证结果可追溯。\n\n"
                    "【主要更新】\n"
                    "1. 温升列识别规则修正：\n"
                    "   - “笔壳温度”列改为同时包含“笔壳”和“温度”两个关键词。\n"
                    "   - “环境温度”列改为同时包含“环境”和“温度”两个关键词。\n"
                    "   - 移除“未匹配环境关键词时自动使用另一列温度”的兜底逻辑。\n"
                    "2. 温度列保留策略优化：\n"
                    "   - 仅当笔壳温度与环境温度同时匹配成功时，才按温升列输出。\n"
                    "   - 若仅匹配到其中一列，该列将保留在“其它数据”中，不再丢失。\n"
                    "3. 电流方向修正规则修正：\n"
                    "   - 由“最大值判定符号”调整为“绝对值最大值对应符号判定”。\n"
                    "   - 当触发整列取相反数时，新增 warning 日志提示。\n"
                    "4. 结果展示与单位格式统一：\n"
                    "   - 汇总表中 mA、V 单位前增加空格。\n"
                    "   - 温升相关三项统一增加 °C 单位。\n"
                    "   - 最终输出 Excel 中所有有数据单元格统一设置为居中显示。\n\n"
                    "5. 图表显示规则优化：\n"
                    "   - 双坐标轴图表仅保留左侧坐标轴网格线。\n"
                    "   - 图例统一调整到底部显示。\n\n"
                    "6. 图表尺寸与布局优化：\n"
                    "   - 默认图表高度调高，提升可读性。\n"
                    "   - 温升图与上方图表的垂直间距同步增大，避免重叠。\n\n"
                    "【兼容说明】\n"
                    "以上调整不改变批处理入口与文件组织方式，仅修正解析与展示规则。"
                ),
            },
            {
                "version": "V0.3",
                "title": "UI组件优化",
                "time": "2026-02-26",
                "detail": (
                    "【版本目标】\n"
                    "优化文件上传组件与运行日志显示，提升界面美观度与用户体验。\n\n"
                    "【主要更新】\n"
                    "1. 文件上传组件重构：\n"
                    "   - 将拖拽区域与文件列表合并为统一组件。\n"
                    "   - 无文件时显示图标与提示文字，有文件后自动切换显示列表。\n"
                    "   - 整个区域均支持拖拽上传功能。\n"
                    "2. 运行日志显示优化：\n"
                    "   - 日志级别标签增加背景色区分（INFO/WARN/ERROR）。\n"
                    "   - 时间戳显示为蓝色，[成功]显示为绿色，[失败]显示为红色。\n"
                    "   - 正文内容使用深灰色，提高可读性。\n"
                    "   - “清空日志”按钮移至组件底部。\n"
                    "3. 组件标题样式优化：\n"
                    "   - 标题增加渐变背景色与边框，视觉更突出。\n"
                    "4. 全局样式调整：\n"
                    "   - 统一控件尺寸与字体大小，界面更紧凑。\n"
                    "   - 优化按钮、输入框、列表项的内边距与圆角。\n\n"
                    "【兼容说明】\n"
                    "本版本仅修改 UI 层；统计流程、文件解析与输出逻辑保持不变。"
                ),
            },
            {
                "version": "V0.2",
                "title": "界面优化升级",
                "time": "2026-02-25",
                "detail": (
                    "【版本目标】\n"
                    "聚焦界面可用性与信息组织优化，在不修改统计处理核心逻辑的前提下，提升日常操作效率与可读性。\n\n"
                    "【主要更新】\n"
                    "1. 充电测试界面重构为左右双栏：\n"
                    "   - 左侧固定为“文件上传、数据处理、输出设置”三段流程化区域。\n"
                    "   - 右侧为独立运行日志区，支持清空和实时滚动显示。\n"
                    "2. 统一标题层级：\n"
                    "   - “运行日志”改为与其它分组一致的分组标题样式与对齐方式。\n"
                    "3. 字体与间距优化：\n"
                    "   - 全局字体调整为更紧凑的浅色桌面风格字号，减轻视觉拥挤。\n"
                    "   - 输入框、按钮、分组说明文字的字号与留白重新标定。\n"
                    "4. 页面职责拆分：\n"
                    "   - 关于页与更新日志页移除运行日志显示，专注于文档信息展示。\n"
                    "5. 视觉主题升级：\n"
                    "   - 全局改为浅色主基调，增强表单可读性与长时间使用舒适度。\n\n"
                    "【兼容说明】\n"
                    "本版本仅修改 UI 层；统计流程、文件解析与输出逻辑保持不变。"
                ),
            },
            {
                "version": "V0.1",
                "title": "首版发布",
                "time": "2026-02-23",
                "detail": (
                    "【版本目标】\n"
                    "完成基带测试数据统计工具首个可用版本，打通“充电测试”核心流程。\n\n"
                    "【主要功能】\n"
                    "1. 建立桌面端主框架：\n"
                    "   - 左侧导航页签：充电测试、续航测试（占位）、待开发功能、关于、更新日志。\n"
                    "2. 实现“统计数据”功能：\n"
                    "   - 解析 Excel 关键列（时间、电流、电压、温度）。\n"
                    "   - 计算预充电流、恒流电流、截充电流、满充电压、充电时长。\n"
                    "   - 在满足条件时计算温升指标并输出图表。\n"
                    "3. 实现“合并后统计数据”功能：\n"
                    "   - 按同名文件配对 Excel（.xlsx/.xls）+ csv。\n"
                    "   - 依据秒级时间戳匹配电压数据后统一输出统计结果。\n"
                    "4. 批处理与输出：\n"
                    "   - 支持多文件/文件夹输入。\n"
                    "   - 支持输出目录配置与重名自动追加序号。\n"
                    "5. 基础日志能力：\n"
                    "   - 记录处理开始、完成、告警与错误，便于排障。\n\n"
                    "【限制说明】\n"
                    "续航测试功能在本版本仍为占位，后续版本按需求逐步开放。"
                ),
            }
        ]
        self._build_ui()

    def _build_ui(self) -> None:
        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(24, 24, 24, 24)
        root_layout.setSpacing(0)
        self.stack = QStackedWidget()
        root_layout.addWidget(self.stack)

        list_page = QWidget()
        list_layout = QVBoxLayout(list_page)
        list_layout.setSpacing(16)
        list_layout.setContentsMargins(0, 0, 0, 0)
        hint = QLabel("📝 更新记录（点击查看详情）")
        hint.setObjectName("sectionTitle")
        self.list_widget = QListWidget()
        self.list_widget.setObjectName("changeList")
        for entry in self._entries:
            item = QListWidgetItem(
                f"🌟 {entry['version']}  |  {entry['title']}  |  📅 {entry['time']}"
            )
            self.list_widget.addItem(item)
        self.list_widget.itemClicked.connect(self._on_item_clicked)
        list_layout.addWidget(hint)
        list_layout.addWidget(self.list_widget)
        self.stack.addWidget(list_page)

        detail_page = QWidget()
        detail_layout = QVBoxLayout(detail_page)
        detail_layout.setSpacing(14)
        detail_layout.setContentsMargins(0, 0, 0, 0)
        self.detail_title = QLabel("")
        self.detail_title.setObjectName("sectionTitle")
        self.detail_body = QTextEdit()
        self.detail_body.setObjectName("detailBody")
        self.detail_body.setReadOnly(True)
        back_row = QHBoxLayout()
        back_row.setContentsMargins(0, 8, 0, 0)
        back_row.addStretch()
        back_btn = QPushButton("⬅ 返回列表")
        back_btn.setProperty("accent", "subtle")
        back_btn.clicked.connect(self._back_to_list)
        back_row.addWidget(back_btn)
        detail_layout.addWidget(self.detail_title)
        detail_layout.addWidget(self.detail_body)
        detail_layout.addLayout(back_row)
        self.stack.addWidget(detail_page)

    def _on_item_clicked(self, item: QListWidgetItem) -> None:
        index = self.list_widget.row(item)
        entry = self._entries[index]
        self.detail_title.setText(f"📌 {entry['version']} - {entry['title']} ({entry['time']})")
        self.detail_body.setPlainText(entry["detail"])
        self.stack.setCurrentIndex(1)

    def _back_to_list(self) -> None:
        self.stack.setCurrentIndex(0)
