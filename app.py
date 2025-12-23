
# app.py
import os
import sys
from dataclasses import dataclass
from typing import List

import pythoncom  # ✅ 解决 CoInitialize 报错（线程内初始化 COM）

from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSettings
from PyQt5.QtGui import QIcon, QPixmap, QFont, QPainter, QPainterPath
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QVBoxLayout, QHBoxLayout,
    QListWidget, QListWidgetItem, QFileDialog, QMessageBox, QGroupBox,
    QRadioButton, QLineEdit, QCheckBox, QSpinBox, QTextEdit, QProgressBar,
    QFrame
)

from word_processor import process_document


# ========= 资源路径（兼容开发环境 & PyInstaller） =========
def resource_path(relative_path: str) -> str:
    """
    PyInstaller onefile 会把资源解压到 sys._MEIPASS
    开发环境则使用当前文件目录
    """
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


# ========= 你的绝对路径（兜底） =========
ABS_LOGO = r"C:\Users\MY43DN\Desktop\ing-logo.png"
ABS_ICON = r"C:\Users\MY43DN\Desktop\app.ico"

# ========= 优先用同目录资源（推荐），找不到再用绝对路径 =========
DEFAULT_LOGO = resource_path("ing-logo.png")
DEFAULT_ICON = resource_path("app.ico")


def is_word_file(path: str) -> bool:
    return os.path.splitext(path)[1].lower() in (".doc", ".docx")


def neon_stylesheet() -> str:
    """黑科技风 QSS：深色 + 霓虹高亮 + 圆角卡片 + 金色脚注"""
    return r"""
    QWidget {
        background-color: #0B0F14;
        color: #D9E2EF;
        font-family: "Microsoft YaHei UI", "Microsoft YaHei", "Segoe UI";
        font-size: 11pt;
    }

    QLabel#Title {
        font-size: 18pt;
        font-weight: 700;
        color: #EAF2FF;
        letter-spacing: 0.5px;
    }
    QLabel#SubTitle {
        font-size: 10pt;
        color: rgba(217,226,239,0.70);
    }

    QFrame#Card {
        background-color: #0E141C;
        border: 1px solid rgba(120, 170, 255, 0.18);
        border-radius: 14px;
    }
    QFrame#Card:hover {
        border: 1px solid rgba(120, 170, 255, 0.35);
    }

    QGroupBox {
        border: 1px solid rgba(120, 170, 255, 0.18);
        border-radius: 12px;
        margin-top: 12px;
        padding: 12px;
        background-color: rgba(14,20,28,0.55);
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        left: 12px;
        padding: 0 6px;
        color: rgba(180, 210, 255, 0.90);
        font-weight: 600;
    }

    QLineEdit {
        background-color: #0B1119;
        border: 1px solid rgba(120, 170, 255, 0.18);
        border-radius: 10px;
        padding: 8px 10px;
        selection-background-color: #00E5FF;
        selection-color: #001018;
    }
    QLineEdit:focus {
        border: 1px solid rgba(0, 229, 255, 0.65);
        background-color: #0A1017;
    }

    QPushButton {
        background-color: rgba(0, 229, 255, 0.10);
        border: 1px solid rgba(0, 229, 255, 0.25);
        color: #CFFBFF;
        padding: 10px 12px;
        border-radius: 12px;
        font-weight: 600;
    }
    QPushButton:hover {
        background-color: rgba(0, 229, 255, 0.18);
        border: 1px solid rgba(0, 229, 255, 0.40);
    }
    QPushButton:pressed {
        background-color: rgba(0, 229, 255, 0.08);
    }
    QPushButton:disabled {
        background-color: rgba(255,255,255,0.05);
        border: 1px solid rgba(255,255,255,0.10);
        color: rgba(217,226,239,0.35);
    }
    QPushButton#Primary {
        background-color: rgba(0, 229, 255, 0.18);
        border: 1px solid rgba(0, 229, 255, 0.55);
        color: #EAFDFF;
        font-size: 12pt;
        padding: 12px 14px;
    }

    QListWidget {
        background-color: #0B1119;
        border: 1px solid rgba(120, 170, 255, 0.18);
        border-radius: 12px;
        padding: 8px;
    }
    QListWidget::item {
        padding: 10px 10px;
        margin: 4px;
        border-radius: 10px;
        background-color: rgba(255,255,255,0.03);
        border: 1px solid rgba(255,255,255,0.05);
    }
    QListWidget::item:selected {
        background-color: rgba(0, 229, 255, 0.12);
        border: 1px solid rgba(0, 229, 255, 0.30);
    }

    QCheckBox, QRadioButton {
        spacing: 8px;
        color: rgba(217,226,239,0.92);
    }

    QTextEdit {
        background-color: #070B10;
        border: 1px solid rgba(120, 170, 255, 0.18);
        border-radius: 12px;
        padding: 10px;
        font-family: "Cascadia Mono", "Consolas";
        font-size: 10pt;
        color: rgba(220, 245, 255, 0.92);
    }

    QProgressBar {
        background-color: #0B1119;
        border: 1px solid rgba(120,170,255,0.18);
        border-radius: 10px;
        text-align: center;
        color: rgba(217,226,239,0.80);
        height: 18px;
    }
    QProgressBar::chunk {
        border-radius: 10px;
        background-color: rgba(0, 229, 255, 0.55);
    }

    QLabel#FooterGold {
        color: #D4AF37; /* 金色 */
        font-size: 9pt;
        letter-spacing: 0.6px;
    }
    """


def safe_exists(path) -> bool:
    """防止 os.path.exists(None)"""
    return isinstance(path, (str, bytes, os.PathLike)) and bool(path) and os.path.exists(path)


def pick_resource(preferred: str, fallback: str) -> str:
    """优先用 preferred（一般是 resource_path），不存在就用 fallback（绝对路径）"""
    if safe_exists(preferred):
        return preferred
    if safe_exists(fallback):
        return fallback
    return ""


def rounded_square_pixmap(image_path: str, size: int = 64, radius: int = 16) -> QPixmap:
    """Logo：居中裁成正方形 + 缩放 + 圆角蒙版（不变形）"""
    pix = QPixmap(image_path)
    if pix.isNull():
        return QPixmap()

    w, h = pix.width(), pix.height()
    side = min(w, h)
    x = (w - side) // 2
    y = (h - side) // 2
    pix = pix.copy(x, y, side, side)
    pix = pix.scaled(size, size, Qt.KeepAspectRatio, Qt.SmoothTransformation)

    out = QPixmap(size, size)
    out.fill(Qt.transparent)

    painter = QPainter(out)
    painter.setRenderHint(QPainter.Antialiasing, True)
    path = QPainterPath()
    path.addRoundedRect(0, 0, size, size, radius, radius)
    painter.setClipPath(path)
    painter.drawPixmap(0, 0, pix)
    painter.end()

    return out


class DropListWidget(QListWidget):
    filesDropped = pyqtSignal(list)

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setSelectionMode(self.ExtendedSelection)
        self.setToolTip("把 .doc/.docx 文件拖进来（也支持拖文件夹：自动读取文件夹内 doc/docx）")

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        event.acceptProposedAction()

    def dropEvent(self, event):
        paths = []
        for url in event.mimeData().urls():
            p = url.toLocalFile()
            if os.path.isdir(p):
                for name in os.listdir(p):
                    fp = os.path.join(p, name)
                    if os.path.isfile(fp) and is_word_file(fp):
                        paths.append(fp)
            else:
                if os.path.isfile(p) and is_word_file(p):
                    paths.append(p)

        if paths:
            self.filesDropped.emit(paths)
        event.acceptProposedAction()


@dataclass
class JobConfig:
    naming_mode: str      # "overwrite" | "suffix" | "custom"
    suffix: str
    custom_name: str
    output_dir: str
    use_same_dir: bool
    output_ext: str       # ".docx" or ".doc"
    keep_blank_lines: int
    tab_to_space: bool
    compress_spaces: bool
    process_headers_footers: bool


class Worker(QThread):
    log = pyqtSignal(str)
    progress = pyqtSignal(int, int)
    finished_ok = pyqtSignal()
    failed = pyqtSignal(str)

    def __init__(self, files: List[str], cfg: JobConfig):
        super().__init__()
        self.files = files
        self.cfg = cfg

    def build_output_path(self, in_path: str) -> str:
        base_dir = os.path.dirname(in_path)
        in_name = os.path.splitext(os.path.basename(in_path))[0]

        out_dir = base_dir if self.cfg.use_same_dir or not self.cfg.output_dir else self.cfg.output_dir
        os.makedirs(out_dir, exist_ok=True)

        # ✅ 严格按用户选择
        if self.cfg.naming_mode == "overwrite":
            out_name = in_name
        elif self.cfg.naming_mode == "custom":
            out_name = self.cfg.custom_name if self.cfg.custom_name else (in_name + "_cleaned")
        else:
            suf = self.cfg.suffix if self.cfg.suffix else "_cleaned"
            out_name = in_name + suf

        return os.path.join(out_dir, out_name + self.cfg.output_ext)

    def run(self):
        # ✅ 关键：线程内初始化 COM，避免 CoInitialize 报错
        pythoncom.CoInitialize()

        total = len(self.files)
        try:
            for i, f in enumerate(self.files, start=1):
                outp = self.build_output_path(f)
                self.log.emit(f"🚀 开始处理：{f}")
                self.log.emit(f"📦 输出位置：{outp}")

                process_document(
                    f, outp,
                    keep_max_blank_lines=self.cfg.keep_blank_lines,
                    tab_to_space=self.cfg.tab_to_space,
                    compress_spaces=self.cfg.compress_spaces,
                    process_headers_footers=self.cfg.process_headers_footers
                )

                self.log.emit("✅ 完成\n")
                self.progress.emit(i, total)

            self.finished_ok.emit()

        except Exception as e:
            self.failed.emit(str(e))

        finally:
            pythoncom.CoUninitialize()


class Card(QFrame):
    def __init__(self):
        super().__init__()
        self.setObjectName("Card")
        self.setFrameShape(QFrame.NoFrame)


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("MY43DN", "WordCleanerUI_Neon")

        self.setWindowTitle(" Word 格式炼化器")

        # ✅ 选择 icon（优先同目录，找不到用绝对路径）
        icon_path = pick_resource(DEFAULT_ICON, ABS_ICON)
        if safe_exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.setFont(QFont("Microsoft YaHei UI", 10))

        root = QVBoxLayout(self)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(12)

        # ===== 顶部品牌区 =====
        header = Card()
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(14, 12, 14, 12)
        header_layout.setSpacing(12)

        self.logo_label = QLabel()
        self.logo_label.setFixedSize(64, 64)

        # ✅ 选择 logo（优先同目录，找不到用绝对路径）
        logo_path = pick_resource(DEFAULT_LOGO, ABS_LOGO)
        if safe_exists(logo_path):
            self.logo_label.setPixmap(rounded_square_pixmap(logo_path, size=64, radius=16))
        else:
            self.logo_label.setText("Logo 未找到")

        title_box = QVBoxLayout()
        self.title = QLabel(" Word 格式炼化器")
        self.title.setObjectName("Title")
        self.sub = QLabel("拖入文件 → 一键清理空格/Tab → 假列表转真列表 → 规则化输出")
        self.sub.setObjectName("SubTitle")
        title_box.addWidget(self.title)
        title_box.addWidget(self.sub)

        header_layout.addWidget(self.logo_label)
        header_layout.addLayout(title_box)
        header_layout.addStretch(1)
        root.addWidget(header)

        # ===== 中间：左文件 / 右配置 =====
        mid = QHBoxLayout()
        mid.setSpacing(12)

        # --- 左：文件区 ---
        left_card = Card()
        left_layout = QVBoxLayout(left_card)
        left_layout.setContentsMargins(14, 14, 14, 14)
        left_layout.setSpacing(10)

        hint = QLabel("📥 将 .doc / .docx 拖到下面；也可点“添加文件”。（支持拖文件夹）")
        hint.setStyleSheet("color: rgba(217,226,239,0.72);")
        left_layout.addWidget(hint)

        self.listw = DropListWidget()
        self.listw.filesDropped.connect(self.add_files)
        left_layout.addWidget(self.listw, 1)

        btn_row = QHBoxLayout()
        self.btn_add = QPushButton("➕ 添加文件")
        self.btn_remove = QPushButton("🗑️ 移除选中")
        self.btn_clear = QPushButton("🧹 清空列表")
        btn_row.addWidget(self.btn_add)
        btn_row.addWidget(self.btn_remove)
        btn_row.addWidget(self.btn_clear)
        left_layout.addLayout(btn_row)

        self.btn_add.clicked.connect(self.pick_files)
        self.btn_remove.clicked.connect(self.remove_selected)
        self.btn_clear.clicked.connect(self.listw.clear)

        mid.addWidget(left_card, 2)

        # --- 右：配置区 ---
        right_card = Card()
        right_layout = QVBoxLayout(right_card)
        right_layout.setContentsMargins(14, 14, 14, 14)
        right_layout.setSpacing(10)

        # 输出策略
        g_out = QGroupBox("输出策略（命名）")
        v = QVBoxLayout(g_out)

        self.rb_overwrite = QRadioButton("覆盖模式：输出文件名与原文件一致（按输出目录落盘）")
        self.rb_suffix = QRadioButton("后缀模式：原文件名 + 后缀")
        self.rb_custom = QRadioButton("自定义模式：仅单文件可用")
        self.rb_suffix.setChecked(True)

        self.ed_suffix = QLineEdit(self.settings.value("suffix", "_cleaned"))
        self.ed_custom = QLineEdit(self.settings.value("custom_name", "炼化输出"))

        v.addWidget(self.rb_overwrite)
        v.addWidget(self.rb_suffix)
        v.addWidget(self.rb_custom)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("后缀："))
        row1.addWidget(self.ed_suffix)
        v.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("自定义名："))
        row2.addWidget(self.ed_custom)
        v.addLayout(row2)

        right_layout.addWidget(g_out)

        # 输出目录
        g_dir = QGroupBox("输出位置（目录）")
        v2 = QVBoxLayout(g_dir)

        self.cb_same_dir = QCheckBox("输出到原文件所在目录（推荐）")
        self.cb_same_dir.setChecked(True)

        rowd = QHBoxLayout()
        self.ed_outdir = QLineEdit(self.settings.value("out_dir", ""))
        self.btn_outdir = QPushButton("📁 选择目录")
        rowd.addWidget(self.ed_outdir)
        rowd.addWidget(self.btn_outdir)

        v2.addWidget(self.cb_same_dir)
        v2.addLayout(rowd)

        self.btn_outdir.clicked.connect(self.pick_outdir)
        right_layout.addWidget(g_dir)

        # 清理选项
        g_cfg = QGroupBox("炼化参数（清理规则）")
        v3 = QVBoxLayout(g_cfg)

        self.cb_tab2space = QCheckBox("Tab → 空格（统一制表符）")
        self.cb_tab2space.setChecked(True)

        self.cb_compress = QCheckBox("压缩连续空格（多空格→1个）")
        self.cb_compress.setChecked(True)

        self.cb_hf = QCheckBox("处理页眉/页脚")
        self.cb_hf.setChecked(True)

        rowb = QHBoxLayout()
        rowb.addWidget(QLabel("连续空行最多保留："))
        self.sp_blank = QSpinBox()
        self.sp_blank.setRange(0, 10)
        self.sp_blank.setValue(1)
        rowb.addWidget(self.sp_blank)
        rowb.addStretch(1)

        rowe = QHBoxLayout()
        rowe.addWidget(QLabel("输出格式："))
        self.rb_docx = QRadioButton(".docx（推荐）")
        self.rb_doc = QRadioButton(".doc")
        self.rb_docx.setChecked(True)
        rowe.addWidget(self.rb_docx)
        rowe.addWidget(self.rb_doc)
        rowe.addStretch(1)

        v3.addWidget(self.cb_tab2space)
        v3.addWidget(self.cb_compress)
        v3.addWidget(self.cb_hf)
        v3.addLayout(rowb)
        v3.addLayout(rowe)

        right_layout.addWidget(g_cfg)

        # 开始按钮
        self.btn_run = QPushButton("⚡ 一键炼化 / 开始处理")
        self.btn_run.setObjectName("Primary")
        right_layout.addWidget(self.btn_run)

        right_layout.addStretch(1)

        mid.addWidget(right_card, 1)
        root.addLayout(mid, 1)

        # ===== 底部：进度 + 状态 =====
        bottom = Card()
        bottom_layout = QHBoxLayout(bottom)
        bottom_layout.setContentsMargins(14, 10, 14, 10)
        bottom_layout.setSpacing(10)

        self.status_label = QLabel("状态：待命")
        self.status_label.setStyleSheet("color: rgba(217,226,239,0.75);")

        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setTextVisible(True)

        bottom_layout.addWidget(self.status_label)
        bottom_layout.addWidget(self.progress, 1)

        root.addWidget(bottom)

        # ===== 日志 =====
        log_card = Card()
        log_layout = QVBoxLayout(log_card)
        log_layout.setContentsMargins(14, 14, 14, 14)
        log_layout.setSpacing(10)

        log_title = QLabel("🧾 运行日志（工程模式）")
        log_title.setStyleSheet("color: rgba(180,210,255,0.90); font-weight: 600;")
        log_layout.addWidget(log_title)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setPlaceholderText("这里会输出处理过程日志…")
        log_layout.addWidget(self.log)

        root.addWidget(log_card)

        # ✅ 你要的：软件底部金色小字（Designed by ...）
        self.footer = QLabel("Designed by 余智秋 in Shanghai.")
        self.footer.setObjectName("FooterGold")
        self.footer.setAlignment(Qt.AlignCenter)
        root.addWidget(self.footer)

        # ===== 绑定 =====
        self.btn_run.clicked.connect(self.run_job)

        self.rb_overwrite.toggled.connect(self.sync_mode_ui)
        self.rb_suffix.toggled.connect(self.sync_mode_ui)
        self.rb_custom.toggled.connect(self.sync_mode_ui)
        self.sync_mode_ui()

        self.worker = None
        self.resize(1200, 800)

    def sync_mode_ui(self):
        """根据输出策略启用/禁用输入框，避免误用"""
        if self.rb_suffix.isChecked():
            self.ed_suffix.setEnabled(True)
            self.ed_custom.setEnabled(False)
        elif self.rb_custom.isChecked():
            self.ed_suffix.setEnabled(False)
            self.ed_custom.setEnabled(True)
        else:
            self.ed_suffix.setEnabled(False)
            self.ed_custom.setEnabled(False)

    def append_log(self, s: str):
        self.log.append(s)

    def add_files(self, files: List[str]):
        existing = set(self.get_all_files())
        for f in files:
            if f not in existing and is_word_file(f):
                item = QListWidgetItem(f)
                item.setToolTip(f)
                self.listw.addItem(item)
        self.status_label.setText(f"状态：已加载 {self.listw.count()} 个文件")

    def get_all_files(self) -> List[str]:
        return [self.listw.item(i).text() for i in range(self.listw.count())]

    def pick_files(self):
        last = self.settings.value("last_open_dir", os.path.expanduser("~"))
        paths, _ = QFileDialog.getOpenFileNames(self, "选择 Word 文件", last, "Word 文件 (*.doc *.docx)")
        if paths:
            self.settings.setValue("last_open_dir", os.path.dirname(paths[0]))
            self.add_files(paths)

    def pick_outdir(self):
        last = self.settings.value("last_out_dir", os.path.expanduser("~"))
        d = QFileDialog.getExistingDirectory(self, "选择输出目录", last)
        if d:
            self.settings.setValue("last_out_dir", d)
            self.ed_outdir.setText(d)
            self.cb_same_dir.setChecked(False)

    def remove_selected(self):
        for item in self.listw.selectedItems():
            self.listw.takeItem(self.listw.row(item))
        self.status_label.setText(f"状态：已加载 {self.listw.count()} 个文件")

    def build_config(self) -> JobConfig:
        if self.rb_overwrite.isChecked():
            mode = "overwrite"
        elif self.rb_custom.isChecked():
            mode = "custom"
        else:
            mode = "suffix"

        cfg = JobConfig(
            naming_mode=mode,
            suffix=self.ed_suffix.text().strip(),
            custom_name=self.ed_custom.text().strip(),
            output_dir=self.ed_outdir.text().strip(),
            use_same_dir=self.cb_same_dir.isChecked(),
            output_ext=".docx" if self.rb_docx.isChecked() else ".doc",
            keep_blank_lines=int(self.sp_blank.value()),
            tab_to_space=self.cb_tab2space.isChecked(),
            compress_spaces=self.cb_compress.isChecked(),
            process_headers_footers=self.cb_hf.isChecked(),
        )

        self.settings.setValue("suffix", cfg.suffix)
        self.settings.setValue("custom_name", cfg.custom_name)
        self.settings.setValue("out_dir", cfg.output_dir)
        return cfg

    def run_job(self):
        files = self.get_all_files()
        if not files:
            QMessageBox.warning(self, "未检测到文件", "请先拖入或添加 .doc/.docx 文件。")
            return

        cfg = self.build_config()

        if cfg.naming_mode == "custom" and len(files) != 1:
            QMessageBox.warning(self, "自定义模式限制", "自定义输出名仅支持单文件处理。")
            return

        if (not cfg.use_same_dir) and (not cfg.output_dir):
            QMessageBox.warning(self, "输出目录为空", "请选择输出目录，或勾选“输出到原目录”。")
            return

        self.btn_run.setEnabled(False)
        self.progress.setValue(0)
        self.status_label.setText("状态：炼化启动中…")

        self.append_log("========== 🚀 任务启动 ==========")
        self.append_log(f"文件数量：{len(files)}")
        self.append_log(f"输出策略：{cfg.naming_mode}")
        self.append_log(f"输出格式：{cfg.output_ext}")
        self.append_log("================================\n")

        self.worker = Worker(files, cfg)
        self.worker.log.connect(self.append_log)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished_ok.connect(self.on_done)
        self.worker.failed.connect(self.on_fail)
        self.worker.start()

    def on_progress(self, done: int, total: int):
        pct = int(done * 100 / total)
        self.progress.setValue(pct)
        self.status_label.setText(f"状态：处理中 {done}/{total}（{pct}%）")

    def on_done(self):
        self.append_log("========== ✅ 全部完成 ==========")
        self.status_label.setText("状态：完成 ✅")
        self.progress.setValue(100)
        self.btn_run.setEnabled(True)
        QMessageBox.information(self, "完成", "所有文件处理完成！")

    def on_fail(self, err: str):
        self.append_log("========== ❌ 发生错误 ==========")
        self.append_log(err)
        self.status_label.setText("状态：失败 ❌")
        self.btn_run.setEnabled(True)
        QMessageBox.critical(self, "错误", f"处理失败：\n{err}")


def main():
    app = QApplication(sys.argv)
    app.setStyleSheet(neon_stylesheet())
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
