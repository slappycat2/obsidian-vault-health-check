# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'v_chk_setupscreen.ui'
##
## Created by: Qt User Interface Compiler version 6.9.1
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PyQt5.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect,
    QSize, QTime, QUrl, Qt)
from PyQt5.QtGui import (QBrush, QColor, QConicalGradient, QCursor,
    QFont, QFontDatabase, QGradient, QIcon,
    QImage, QKeySequence, QLinearGradient, QPainter,
    QPalette, QPixmap, QRadialGradient, QTransform)
from PyQt5.QtWidgets import (QAbstractButton, QApplication, QCheckBox, QComboBox,
    QDialog, QDialogButtonBox, QLabel, QLineEdit,
    QPushButton, QSizePolicy, QVBoxLayout, QWidget)

class Ui_v_chk_setup_dlg(object):
    def setupUi(self, v_chk_setup_dlg):
        if not v_chk_setup_dlg.objectName():
            v_chk_setup_dlg.setObjectName(u"v_chk_setup_dlg")
        v_chk_setup_dlg.resize(494, 302)
        icon = QIcon()
        icon.addFile(u"../img/swenlogo.ico", QSize(), QIcon.Mode.Normal, QIcon.State.Off)
        v_chk_setup_dlg.setWindowIcon(icon)
        self.buttonBox = QDialogButtonBox(v_chk_setup_dlg)
        self.buttonBox.setObjectName(u"buttonBox")
        self.buttonBox.setGeometry(QRect(110, 240, 341, 32))
        self.buttonBox.setOrientation(Qt.Horizontal)
        self.buttonBox.setStandardButtons(QDialogButtonBox.Cancel|QDialogButtonBox.Ok)
        self.logo = QLabel(v_chk_setup_dlg)
        self.logo.setObjectName(u"logo")
        self.logo.setGeometry(QRect(380, 10, 101, 101))
        self.logo.setPixmap(QPixmap(u"../img/SwenLogo125.png"))
        self.verticalLayoutWidget = QWidget(v_chk_setup_dlg)
        self.verticalLayoutWidget.setObjectName(u"verticalLayoutWidget")
        self.verticalLayoutWidget.setGeometry(QRect(20, 70, 131, 81))
        self.vl_lbls_dirs = QVBoxLayout(self.verticalLayoutWidget)
        self.vl_lbls_dirs.setObjectName(u"vl_lbls_dirs")
        self.vl_lbls_dirs.setContentsMargins(0, 0, 0, 0)
        self.lbl_dir_vault_ = QLabel(self.verticalLayoutWidget)
        self.lbl_dir_vault_.setObjectName(u"lbl_dir_vault_")
        self.lbl_dir_vault_.setTextFormat(Qt.PlainText)

        self.vl_lbls_dirs.addWidget(self.lbl_dir_vault_)

        self.lbl_dirs_ignore = QLabel(self.verticalLayoutWidget)
        self.lbl_dirs_ignore.setObjectName(u"lbl_dirs_ignore")
        self.lbl_dirs_ignore.setTextFormat(Qt.PlainText)

        self.vl_lbls_dirs.addWidget(self.lbl_dirs_ignore)

        self.lbl_pn_wb_exec = QLabel(self.verticalLayoutWidget)
        self.lbl_pn_wb_exec.setObjectName(u"lbl_pn_wb_exec")
        self.lbl_pn_wb_exec.setTextFormat(Qt.PlainText)

        self.vl_lbls_dirs.addWidget(self.lbl_pn_wb_exec)

        self.verticalLayoutWidget_2 = QWidget(v_chk_setup_dlg)
        self.verticalLayoutWidget_2.setObjectName(u"verticalLayoutWidget_2")
        self.verticalLayoutWidget_2.setGeometry(QRect(160, 70, 191, 81))
        self.vl_dirs = QVBoxLayout(self.verticalLayoutWidget_2)
        self.vl_dirs.setObjectName(u"vl_dirs")
        self.vl_dirs.setContentsMargins(0, 0, 0, 0)
        self.lov_vault_path = QComboBox(self.verticalLayoutWidget_2)
        self.lov_vault_path.setObjectName(u"lov_vault_path")

        self.vl_dirs.addWidget(self.lov_vault_path)

        self.csl_dirs_ignore = QLineEdit(self.verticalLayoutWidget_2)
        self.csl_dirs_ignore.setObjectName(u"csl_dirs_ignore")

        self.vl_dirs.addWidget(self.csl_dirs_ignore)

        self.lineEdit_2 = QLineEdit(self.verticalLayoutWidget_2)
        self.lineEdit_2.setObjectName(u"lineEdit_2")

        self.vl_dirs.addWidget(self.lineEdit_2)

        self.verticalLayoutWidget_3 = QWidget(v_chk_setup_dlg)
        self.verticalLayoutWidget_3.setObjectName(u"verticalLayoutWidget_3")
        self.verticalLayoutWidget_3.setGeometry(QRect(20, 180, 131, 41))
        self.vl_options = QVBoxLayout(self.verticalLayoutWidget_3)
        self.vl_options.setObjectName(u"vl_options")
        self.vl_options.setContentsMargins(0, 0, 0, 0)
        self.chk_show_wb_notes = QCheckBox(self.verticalLayoutWidget_3)
        self.chk_show_wb_notes.setObjectName(u"chk_show_wb_notes")

        self.vl_options.addWidget(self.chk_show_wb_notes)

        self.chk_use_full_paths = QCheckBox(self.verticalLayoutWidget_3)
        self.chk_use_full_paths.setObjectName(u"chk_use_full_paths")

        self.vl_options.addWidget(self.chk_use_full_paths)

        self.btn_browse_wb_exec = QPushButton(v_chk_setup_dlg)
        self.btn_browse_wb_exec.setObjectName(u"btn_browse_wb_exec")
        self.btn_browse_wb_exec.setGeometry(QRect(358, 126, 56, 17))
        self.lbl_hdr = QLabel(v_chk_setup_dlg)
        self.lbl_hdr.setObjectName(u"lbl_hdr")
        self.lbl_hdr.setGeometry(QRect(20, 30, 221, 21))
        font = QFont()
        font.setPointSize(12)
        self.lbl_hdr.setFont(font)

        self.retranslateUi(v_chk_setup_dlg)
        self.buttonBox.accepted.connect(v_chk_setup_dlg.accept)
        self.buttonBox.rejected.connect(v_chk_setup_dlg.reject)

        QMetaObject.connectSlotsByName(v_chk_setup_dlg)
    # setupUi

    def retranslateUi(self, v_chk_setup_dlg):
        v_chk_setup_dlg.setWindowTitle(QCoreApplication.translate("v_chk_setup_dlg", u"Obsidian Vault Health Check", None))
        self.logo.setText("")
        self.lbl_dir_vault_.setText(QCoreApplication.translate("v_chk_setup_dlg", u"Obsidian Vault Path:", None))
#if QT_CONFIG(tooltip)
        self.lbl_dirs_ignore.setToolTip(QCoreApplication.translate("v_chk_setup_dlg", u"Enter a comma separated list of folders to exclude from the analysis", None))
#endif // QT_CONFIG(tooltip)
#if QT_CONFIG(statustip)
        self.lbl_dirs_ignore.setStatusTip(QCoreApplication.translate("v_chk_setup_dlg", u"Enter a comma separated list of folders to exclude from the analysis", None))
#endif // QT_CONFIG(statustip)
        self.lbl_dirs_ignore.setText(QCoreApplication.translate("v_chk_setup_dlg", u"Vault Directories To Ignore:", None))
#if QT_CONFIG(tooltip)
        self.lbl_pn_wb_exec.setToolTip(QCoreApplication.translate("v_chk_setup_dlg", u"Enter the full path of your workbook program executale file.", None))
#endif // QT_CONFIG(tooltip)
#if QT_CONFIG(statustip)
        self.lbl_pn_wb_exec.setStatusTip(QCoreApplication.translate("v_chk_setup_dlg", u"Tell us where to find your spreadsheet program", None))
#endif // QT_CONFIG(statustip)
        self.lbl_pn_wb_exec.setText(QCoreApplication.translate("v_chk_setup_dlg", u"Full Path to Workbook Executable:", None))
#if QT_CONFIG(tooltip)
        self.csl_dirs_ignore.setToolTip(QCoreApplication.translate("v_chk_setup_dlg", u"Enter a comma separated list of folders to exclude from the analysis", None))
#endif // QT_CONFIG(tooltip)
#if QT_CONFIG(statustip)
        self.csl_dirs_ignore.setStatusTip(QCoreApplication.translate("v_chk_setup_dlg", u"Enter a comma separated list of folders to exclude from the analysis", None))
#endif // QT_CONFIG(statustip)
#if QT_CONFIG(tooltip)
        self.chk_show_wb_notes.setToolTip(QCoreApplication.translate("v_chk_setup_dlg", u"Include How-To Notes on each tab, detailing usage and tips", None))
#endif // QT_CONFIG(tooltip)
        self.chk_show_wb_notes.setText(QCoreApplication.translate("v_chk_setup_dlg", u"Show Help Notes", None))
#if QT_CONFIG(tooltip)
        self.chk_use_full_paths.setToolTip(QCoreApplication.translate("v_chk_setup_dlg", u"Use  Standard Note Names or Relative Vault Path Names in Links", None))
#endif // QT_CONFIG(tooltip)
        self.chk_use_full_paths.setText(QCoreApplication.translate("v_chk_setup_dlg", u"Show Relative Paths in Links", None))
        self.btn_browse_wb_exec.setText(QCoreApplication.translate("v_chk_setup_dlg", u"Browse", None))
        self.lbl_hdr.setText(QCoreApplication.translate("v_chk_setup_dlg", u"Obsidian Vault Health Check Seup", None))
    # retranslateUi

ab = Ui_v_chk_setup_dlg()