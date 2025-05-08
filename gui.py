# gui.py
import sys
import os
from PyQt5 import QtWidgets
from PyQt5.QtCore import QDate, QTime
from storage_excel import ExcelLogger

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("工作记录器")
        self.resize(480, 620)

        # --- 日期 & 时间 ---
        self.date_edit = QtWidgets.QDateEdit(self)
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(QDate.currentDate())
        self.date_edit.setGeometry(30, 20, 200, 30)

        self.start_time = QtWidgets.QTimeEdit(self)
        self.start_time.setTime(QTime.currentTime())
        self.start_time.setGeometry(30, 70, 120, 30)

        self.end_time = QtWidgets.QTimeEdit(self)
        self.end_time.setTime(QTime.currentTime())
        self.end_time.setGeometry(180, 70, 120, 30)

        # --- 工作内容 ---
        lbl1 = QtWidgets.QLabel("工作内容：", self)
        lbl1.setGeometry(30, 120, 80, 25)
        self.work_edit = QtWidgets.QPlainTextEdit(self)
        self.work_edit.setPlaceholderText("填写今天的主要工作内容…")
        self.work_edit.setGeometry(30, 150, 420, 80)

        # --- 效果 ---
        lbl2 = QtWidgets.QLabel("效果：", self)
        lbl2.setGeometry(30, 250, 80, 25)
        self.effect_edit = QtWidgets.QPlainTextEdit(self)
        self.effect_edit.setPlaceholderText("工作效果简述…")
        self.effect_edit.setGeometry(30, 280, 420, 80)

        # --- 备注 ---
        lbl3 = QtWidgets.QLabel("备注：", self)
        lbl3.setGeometry(30, 380, 80, 25)
        self.remark_edit = QtWidgets.QPlainTextEdit(self)
        self.remark_edit.setPlaceholderText("其它补充说明…")
        self.remark_edit.setGeometry(30, 410, 420, 80)

        # --- 保存按钮 ---
        self.save_btn = QtWidgets.QPushButton("保存到 Excel", self)
        self.save_btn.setGeometry(30, 520, 120, 30)
        self.save_btn.clicked.connect(self.save_to_excel)

        # --- 日志存储 ---
        # # 文件会输出到脚本同级目录，可改成其他路径
        # out_dir = os.path.dirname(os.path.abspath(__file__))
        # self.logger = ExcelLogger(out_dir=out_dir)
        # --- 日志存储 ---
        # 输出目录：如果被 PyInstaller 冻结成 exe，用 exe 所在目录；否则用脚本目录
        if getattr(sys, "frozen", False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        self.logger = ExcelLogger(out_dir=base_dir)

    def save_to_excel(self):
        date = self.date_edit.date().toString("yyyy-MM-dd")
        t0 = self.start_time.time().toString("HH:mm")
        t1 = self.end_time.time().toString("HH:mm")
        work   = self.work_edit.toPlainText().strip()
        effect = self.effect_edit.toPlainText().strip()
        remark = self.remark_edit.toPlainText().strip()

        # 校验
        if not work:
            QtWidgets.QMessageBox.warning(self, "警告", "请填写工作内容。")
            return

        try:
            self.logger.log(date, t0, t1, work, effect, remark)
            QtWidgets.QMessageBox.information(self, "完成", "已保存到 Excel！")
            # 清空输入
            self.work_edit.clear()
            self.effect_edit.clear()
            self.remark_edit.clear()
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "错误", f"保存失败：{e}")

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())
