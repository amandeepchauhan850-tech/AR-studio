import sys
import os
from PyQt5.QtWidgets import (QApplication, QDialog, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QFrame, QWidget)
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt

# .exe banane ke baad images ka path dhoondhne ke liye function
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class CustomPasswordDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Enter Password")
        
        # Dialog Size
        self.setFixedSize(400, 380)
        
        # Custom Styling
        self.setStyleSheet("""
            QDialog {
                background-color: #ffffff;
                border-radius: 16px;
            }
            QLineEdit {
                border: 2px solid #a855f7;
                border-radius: 10px;
                padding: 10px 12px;
                background-color: #ffffff;
                font-size: 15px;
                color: #1e293b;
            }
            QLineEdit:focus {
                border: 2px solid #9333ea;
            }
            QPushButton#SubmitBtn {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #a855f7, stop:1 #d946ef);
                color: white;
                border: none;
                border-radius: 10px;
                padding: 12px;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton#SubmitBtn:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #9333ea, stop:1 #c026d3);
            }
            QPushButton#CancelBtn {
                background-color: #e2e8f0;
                color: #475569;
                border: none;
                border-radius: 10px;
                padding: 12px;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton#CancelBtn:hover {
                background-color: #cbd5e1;
            }
        """)

        # Image Paths
        self.img_idle = resource_path("POP 1.png")    # Photo 1: Empty Box
        self.img_closed = resource_path("POP 2.png")  # Photo 2: Typing / Eyes Closed
        self.img_peek = resource_path("POP 3.png")    # Photo 3: Password Visible

        self.is_password_visible = False
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(25, 20, 25, 20)
        layout.setAlignment(Qt.AlignCenter)

        # 1. Monkey Image Display Frame
        self.image_label = QLabel(self)
        self.image_label.setAlignment(Qt.AlignCenter)
        self.image_label.setFixedHeight(140)
        self.update_monkey_image(self.img_idle)
        layout.addWidget(self.image_label)

        # 2. Text Labels
        title = QLabel("Enter Password", self)
        title.setStyleSheet("font-size: 20px; font-weight: bold; color: #0f172a; margin-top: 5px;")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        subtitle = QLabel("Please enter your password to continue", self)
        subtitle.setStyleSheet("font-size: 13px; color: #64748b; margin-bottom: 12px;")
        subtitle.setAlignment(Qt.AlignCenter)
        layout.addWidget(subtitle)

        # 3. Password Box + Colored Eye Button Box
        input_container = QWidget(self)
        input_layout = QHBoxLayout(input_container)
        input_layout.setContentsMargins(0, 0, 0, 0)
        input_layout.setSpacing(5)

        self.password_entry = QLineEdit(input_container)
        self.password_entry.setPlaceholderText("••••••••")
        self.password_entry.setEchoMode(QLineEdit.Password)
        self.password_entry.textChanged.connect(self.on_text_changed)

        # Eye Icon Button (Purple Colored)
        self.eye_button = QPushButton("👁", input_container)
        self.eye_button.setCursor(Qt.PointingHandCursor)
        self.eye_button.setFixedWidth(40)
        self.eye_button.setFixedHeight(40)
        self.eye_button.setStyleSheet("""
            QPushButton {
                border: 2px solid #a855f7;
                border-radius: 10px;
                background-color: #f3e8ff;
                color: #9333ea;
                font-size: 18px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #e9d5ff;
                color: #c026d3;
            }
        """)
        self.eye_button.clicked.connect(self.toggle_password_visibility)

        input_layout.addWidget(self.password_entry)
        input_layout.addWidget(self.eye_button)

        layout.addWidget(input_container)

        # 4. Action Buttons (Cancel / Submit)
        btn_layout = QHBoxLayout()
        btn_layout.setContentsMargins(0, 15, 0, 0)
        btn_layout.setSpacing(12)

        cancel_btn = QPushButton("Cancel", self)
        cancel_btn.setObjectName("CancelBtn")
        cancel_btn.setCursor(Qt.PointingHandCursor)
        cancel_btn.clicked.connect(self.reject)

        submit_btn = QPushButton("Submit", self)
        submit_btn.setObjectName("SubmitBtn")
        submit_btn.setCursor(Qt.PointingHandCursor)
        submit_btn.clicked.connect(self.accept)

        btn_layout.addWidget(cancel_btn)
        btn_layout.addWidget(submit_btn)
        layout.addLayout(btn_layout)

        self.setLayout(layout)

    def update_monkey_image(self, image_path):
        pixmap = QPixmap(image_path)
        if not pixmap.isNull():
            scaled_pixmap = pixmap.scaled(130, 130, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.image_label.setPixmap(scaled_pixmap)

    def on_text_changed(self, text):
        if not text:
            self.update_monkey_image(self.img_idle)
        else:
            if not self.is_password_visible:
                self.update_monkey_image(self.img_closed)

    def toggle_password_visibility(self):
        if self.is_password_visible:
            self.password_entry.setEchoMode(QLineEdit.Password)
            self.is_password_visible = False
            if self.password_entry.text():
                self.update_monkey_image(self.img_closed)  # POP 2.png
            else:
                self.update_monkey_image(self.img_idle)    # POP 1.png
        else:
            self.password_entry.setEchoMode(QLineEdit.Normal)
            self.is_password_visible = True
            self.update_monkey_image(self.img_peek)        # POP 3.png

# Main Execution Block
if __name__ == "__main__":
    app = QApplication(sys.argv)
    dialog = CustomPasswordDialog()
    if dialog.exec_() == QDialog.Accepted:
        print("Entered Password:", dialog.password_entry.text())
    sys.exit()