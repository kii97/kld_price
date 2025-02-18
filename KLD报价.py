import sys
import json
from PyQt5.QtWidgets import QMainWindow, QApplication,QMessageBox
from PyQt5.QtCore import Qt
from UI.main_ui import Ui_MainWindow
from UI.settings_ui import Ui_SettingsWindow
from main import price_object


class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.settings_dic=self.load_json_datas()
        self.setupUi(self)
        self.toolButton.clicked.connect(self.show_settings_window)
        self.kvar_rbt.clicked.connect(self.kvar_clicked)
        self.fire_rbt.clicked.connect(self.fae_clicked)
        self.arc_rbt.clicked.connect(self.ap_clicked)
        self.drag_newedit.textChanged.connect(self.autocheck)
        self.start_btn.clicked.connect(self.start)

    def show_settings_window(self):
        self.settings_dic=self.load_json_datas()
        self.settings_window = SettingsWindow(self.settings_dic)
        self.settings_window.show()

    def kvar_clicked(self):
        self.checkBox_kvar.setVisible(True)
        self.checkBox_BKSTSC.setVisible(True)
        self.checkBox_AP.setVisible(False)
        self.checkBox_FAE.setVisible(False)

    def ap_clicked(self):
        self.checkBox_AP.setVisible(True)
        self.checkBox_FAE.setVisible(False)
        self.checkBox_kvar.setVisible(False)
        self.checkBox_BKSTSC.setVisible(False)

    def fae_clicked(self):
        self.checkBox_FAE.setVisible(True)
        self.checkBox_AP.setVisible(False)
        self.checkBox_kvar.setVisible(False)
        self.checkBox_BKSTSC.setVisible(False)

    def autocheck(self):
        file=self.drag_newedit.displayText().split("/")[-1]
        if "灭火" in file:
            self.fire_rbt.setChecked(True)
            self.fae_clicked()
        elif "弧光" in file:
            self.arc_rbt.setChecked(True)
            self.ap_clicked()
        else:
            self.kvar_rbt.setChecked(True)
            self.kvar_clicked()

    def load_json_datas(self):
        self.settings_dic={}
        try:
            with open('settings.json', 'r') as jsfile:
                return json.loads(jsfile.read())
        except:
            return {}

    def keyPressEvent(self,QKeyEvent):
        if QKeyEvent.key() == Qt.Key_Return:
            self.start()
        elif QKeyEvent.key() == Qt.Key_Enter:
            self.start()
        elif QKeyEvent.key() == Qt.Key_Escape:
            self.close()

    def start(self):
        self.settings_dic=self.load_json_datas()
        self.config_dic={
            "kvar":[self.kvar_rbt.isChecked(),self.checkBox_kvar.isChecked(),self.checkBox_BKSTSC.isChecked()],
            "fae":[self.fire_rbt.isChecked(),self.checkBox_FAE.isChecked()],
            "ap":[self.arc_rbt.isChecked(),self.checkBox_AP.isChecked()]}
        if self.drag_newedit.displayText()=="":
            QMessageBox.information(self,"错误","还没选择文件！")
        else:
            if self.settings_dic["coord_edit"]=="":
                self.settings_dic["coord_edit"]="J3"
            if self.settings_dic["suffix_edit"]=="":
                self.settings_dic["suffix_edit"] ="(1)"
            obt=price_object(self.drag_newedit.displayText(),self.settings_dic,self.config_dic)
            if obt.operate():
                QMessageBox.information(self,"OK", "已成功生成报价单！")
                if not self.checkBox_mode.isChecked():
                    self.close()
            else:
                QMessageBox.information(self,"错误", "出错了！\n可能原因：\n1.配置单内的项目名称有斜杠\n2.开着同名报价单")

class SettingsWindow(QMainWindow, Ui_SettingsWindow):
    def __init__(self,settings_dic):
        super(SettingsWindow, self).__init__()
        self.setupUi(self)
        self.coord_edit.setPlaceholderText("J3")
        self.suffix_edit.setPlaceholderText("(1)")
        self.settings_dic=settings_dic
        self.load()
        self.SaveButton.clicked.connect(self.save)
        self.CancelButton.clicked.connect(self.cancel)

    def save(self):
        dic={
            "checkBox_date":self.checkBox_date.isChecked(),
            "name_edit":self.name_edit.displayText(),
            "code_edit":self.code_edit.displayText(),
            "coord_edit":self.coord_edit.displayText(),
            "groupBox_null":self.groupBox_null.isChecked(),
            "spinBox":self.spinBox.value(),
            "groupBox_suffix":self.groupBox_suffix.isChecked(),
            "suffix_edit":self.suffix_edit.displayText()
        }
        with open('settings.json','w')as jsfile:
            jsfile.write(json.dumps(dic))
        self.close()

    def load(self):
        try:
            self.checkBox_date.setChecked(self.settings_dic["checkBox_date"])
            self.name_edit.setText(self.settings_dic["name_edit"])
            self.code_edit.setText(self.settings_dic["code_edit"])
            self.coord_edit.setText(self.settings_dic["coord_edit"])
            self.groupBox_null.setChecked(self.settings_dic["groupBox_null"])
            self.spinBox.setValue(self.settings_dic["spinBox"])
            self.groupBox_suffix.setChecked(self.settings_dic["groupBox_suffix"])
            self.suffix_edit.setText(self.settings_dic["suffix_edit"])
        except:
            pass

    def cancel(self):
        self.close()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())