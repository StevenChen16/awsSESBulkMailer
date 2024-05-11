import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit, QTextEdit, 
                             QLabel, QDialog, QTableWidget, QTableWidgetItem, QMessageBox, QMenuBar, QHeaderView)
from PyQt6.QtGui import QAction
from PyQt6.QtWidgets import QComboBox
import boto3
from botocore.exceptions import ClientError

class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("SMTP 设置")
        layout = QVBoxLayout()
        self.aws_access_key_id_input = QLineEdit()
        self.aws_secret_access_key_input = QLineEdit()
        self.aws_region_input = QComboBox()
        self.aws_region_input.setEditable(True)  # 允许用户输入值
        # 填充一些常见的AWS区域
        common_regions = [
            'us-east-1', 'us-west-1', 'us-west-2',
            'eu-west-1', 'eu-central-1', 'ap-southeast-1',
            'ap-southeast-2', 'ap-northeast-1', 'sa-east-1'
        ]
        self.aws_region_input.addItems(common_regions)

        layout.addWidget(QLabel("AWS Access Key ID:"))
        layout.addWidget(self.aws_access_key_id_input)
        layout.addWidget(QLabel("AWS Secret Access Key:"))
        layout.addWidget(self.aws_secret_access_key_input)
        layout.addWidget(QLabel("AWS Region:"))
        layout.addWidget(self.aws_region_input)
        save_button = QPushButton("保存")
        save_button.clicked.connect(self.accept)
        layout.addWidget(save_button)
        self.setLayout(layout)

    def get_settings(self):
        return {
            "aws_access_key_id": self.aws_access_key_id_input.text(),
            "aws_secret_access_key": self.aws_secret_access_key_input.text(),
            "aws_region": self.aws_region_input.currentText(),  # 获取用户选择或输入的区域
        }

class RecipientsDialog(QDialog):
    def __init__(self, parent=None, initial_data=[]):
        super().__init__(parent)
        self.setWindowTitle("编辑收件人")
        layout = QVBoxLayout()
        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(["显示名称", "邮箱地址"])
        self.load_initial_data(initial_data)
        layout.addWidget(self.table)
        add_button = QPushButton("添加收件人")
        add_button.clicked.connect(self.add_row)
        layout.addWidget(add_button)
        save_button = QPushButton("保存")
        save_button.clicked.connect(self.accept)
        layout.addWidget(save_button)
        self.setLayout(layout)

    def load_initial_data(self, initial_data):
        for name, email in initial_data:
            self.add_row(name, email)

    def add_row(self, name="", email=""):
        row_count = self.table.rowCount()
        self.table.insertRow(row_count)
        self.table.setItem(row_count, 0, QTableWidgetItem(name))
        self.table.setItem(row_count, 1, QTableWidgetItem(email))

    def get_recipients(self):
        recipients = []
        for row in range(self.table.rowCount()):
            display_name = self.table.item(row, 0)
            email_address = self.table.item(row, 1)
            if display_name and email_address:  # Ensure both items exist
                display_name = display_name.text() if display_name.text() else ""
                email_address = email_address.text()
                recipients.append((display_name, email_address))
        return recipients

class EmailSender(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = {}
        self.recipients = []
        self.initUI()

    def initUI(self):
        self.setWindowTitle('邮件发送器')
        self.setGeometry(100, 100, 800, 600)
        main_layout = QVBoxLayout()
        
        menubar = self.menuBar()
        about_action = QAction('关于', self)
        about_action.triggered.connect(lambda: QMessageBox.information(self, "关于", "邮件发送软件。"))
        settings_action = QAction('设置', self)
        settings_action.triggered.connect(self.show_settings_dialog)
        menubar.addAction(about_action)
        menubar.addAction(settings_action)
        
        sender_layout = QHBoxLayout()
        self.sender_name_input = QLineEdit()
        self.sender_email_input = QLineEdit()
        sender_layout.addWidget(QLabel("显示名称:"))
        sender_layout.addWidget(self.sender_name_input)
        sender_layout.addWidget(QLabel("发件人地址:"))
        sender_layout.addWidget(self.sender_email_input)
        main_layout.addLayout(sender_layout)

        main_layout.addWidget(QLabel("-" * 100))

        recipients_layout = QHBoxLayout()
        self.recipients_table = QTableWidget()
        self.recipients_table.setHorizontalHeaderLabels(["显示名称", "邮箱地址"])
        self.recipients_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        edit_recipients_btn = QPushButton('编辑收件人')
        edit_recipients_btn.clicked.connect(self.edit_recipients)
        recipients_layout.addWidget(self.recipients_table)
        recipients_layout.addWidget(edit_recipients_btn)
        main_layout.addLayout(recipients_layout)

        main_layout.addWidget(QLabel("-" * 100))

        self.subject_input = QLineEdit()
        self.content_input = QTextEdit()
        main_layout.addWidget(QLabel("邮件主题:"))
        main_layout.addWidget(self.subject_input)
        main_layout.addWidget(QLabel("邮件内容:"))
        main_layout.addWidget(self.content_input)
        send_btn = QPushButton('发送邮件')
        send_btn.clicked.connect(self.on_send_click)
        main_layout.addWidget(send_btn)
        
        central_widget = QWidget(self)
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

    def edit_recipients(self):
        dialog = RecipientsDialog(self, self.recipients)
        if dialog.exec():
            self.recipients = dialog.get_recipients()
            self.update_recipients_table()

    def update_recipients_table(self):
        self.recipients_table.setRowCount(0)
        for display_name, email_address in self.recipients:
            row_count = self.recipients_table.rowCount()
            self.recipients_table.insertRow(row_count)
            self.recipients_table.setItem(row_count, 0, QTableWidgetItem(display_name))
            self.recipients_table.setItem(row_count, 1, QTableWidgetItem(email_address))

    def show_settings_dialog(self):
        dialog = SettingsDialog(self)
        if dialog.exec():
            self.settings = dialog.get_settings()

    def send_personalized_email(self, subject, content, recipient):
        # 检查是否已设置所有必要的配置
        if not all(key in self.settings for key in ['aws_access_key_id', 'aws_secret_access_key', 'aws_region']):
            QMessageBox.warning(self, "警告", "AWS密钥和区域信息未配置，请先设置。点击确定以打开设置对话框。",
                                QMessageBox.StandardButton.Ok)
            self.show_settings_dialog()  # 直接打开设置对话框
            return

        aws_region = self.settings['aws_region']
        aws_access_key_id = self.settings['aws_access_key_id']
        aws_secret_access_key = self.settings['aws_secret_access_key']
        sender = self.sender_email_input.text()

        charset = "UTF-8"
        html_body = f"<html><body><h1>{subject}</h1><p>{content}</p></body></html>"
        text_body = content

        ses_client = boto3.client('ses', region_name=aws_region,
                                aws_access_key_id=aws_access_key_id,
                                aws_secret_access_key=aws_secret_access_key)

        try:
            response = ses_client.send_email(
                Destination={'ToAddresses': [recipient]},
                Message={
                    'Body': {'Html': {'Charset': charset, 'Data': html_body}, 'Text': {'Charset': charset, 'Data': text_body}},
                    'Subject': {'Charset': charset, 'Data': subject},
                },
                Source=sender,
            )
        except ClientError as e:
            QMessageBox.critical(self, "邮件发送失败", e.response['Error']['Message'])
            return
        QMessageBox.information(self, "邮件发送", f"邮件已发送！消息ID: {response['MessageId']}")


    def on_send_click(self):
        subject = self.subject_input.text()
        content = self.content_input.toPlainText()
        if not self.recipients:
            QMessageBox.warning(self, "错误", "请添加至少一个收件人")
            return

        for display_name, email_address in self.recipients:
            recipient = f"{display_name} <{email_address}>" if display_name else email_address
            self.send_personalized_email(subject, content, recipient)

app = QApplication(sys.argv)
ex = EmailSender()
ex.show()
sys.exit(app.exec())