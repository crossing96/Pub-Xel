# Pub-Xel - A Biomedical Reference Management Tool
# Copyright (C) 2024  Jongyeob Kim
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version...
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program. If not, see <https://www.gnu.org/licenses/>.


import sys
from PyQt6.QtWidgets import QApplication, QDialog, QVBoxLayout, QLabel, QPushButton, QMainWindow, QStackedLayout, QHBoxLayout, QSpacerItem, QSizePolicy, QWidget
from PyQt6.QtGui import QPixmap, QFont
import os

script_dir = os.path.dirname(os.path.abspath(__file__))
assets_dir = os.path.join(script_dir, 'assets')
src_dir = os.path.join(script_dir, 'src')
ui_dir = os.path.join(script_dir, 'ui')
data_dir = os.path.join(script_dir, 'data')

welcome1_path = os.path.join(assets_dir, 'welcome1.png')

class WelcomeDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Welcome")
        self.setFixedSize(500, 400)  # Set fixed size for the dialog
        
        self.stacked_layout = QStackedLayout()
        
        # First page
        self.page1 = QVBoxLayout()
        self.label1 = QLabel("""Welcome to Pub-Xel!

Pub-Xel is a support tool designed to help you efficiently manage biomedical articles directly within Microsoft Excel. 

Lightweight and user-friendly, Pub-Xel offers powerful features to enhance your study workflow.

Please ensure you have Microsoft Excel installed for the best experience.

Press Next to continue. """)
        self.label1.setWordWrap(True)  # Enable word wrap
        font1 = QFont()
        font1.setPointSize(12)  # Increase font size
        self.label1.setFont(font1)
        self.page1.addWidget(self.label1)

        self.page1.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding))
        
        self.next_button = QPushButton("Next")
        self.next_button.setFixedSize(100, 30)  # Set fixed size for the button
        hbox = QHBoxLayout()
        hbox.addSpacerItem(QSpacerItem(40, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum))
        hbox.addWidget(self.next_button)
        self.page1.addLayout(hbox)
        
        self.next_button.clicked.connect(self.show_next_page)
        
        self.page1_widget = QWidget()
        self.page1_widget.setLayout(self.page1)
        self.stacked_layout.addWidget(self.page1_widget)
        
        # Second page
        self.page2 = QVBoxLayout()
        
        self.label2_before = QLabel("To get started, save a new Pub-Xel Worksheet to begin managing your articles. After pressing OK button below, press the button in the main window, as shown below:\n")
        self.label2_before.setWordWrap(True)  # Enable word wrap
        font2 = QFont()
        font2.setPointSize(12)
        self.label2_before.setFont(font2)
        self.page2.addWidget(self.label2_before)
        
        self.image_label = QLabel()
        pixmap = QPixmap(welcome1_path)  # Replace with the path to your image
        if pixmap.height() > 200:
            pixmap = pixmap.scaledToHeight(200)  # Scale the image to fit within the dialog, maintaining aspect ratio
        self.image_label.setPixmap(pixmap)
        self.page2.addWidget(self.image_label)
        
        self.label2_after = QLabel("Thank you for using Pub-Xel!")
        self.label2_after.setWordWrap(True)  # Enable word wrap
        font2.setPointSize(12)
        self.label2_after.setFont(font2)
        self.page2.addWidget(self.label2_after)

        self.page2.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding))
        
        self.ok_button = QPushButton("OK")
        self.ok_button.setFixedSize(100, 30)  # Set fixed size for the button
        hbox2 = QHBoxLayout()
        hbox2.addSpacerItem(QSpacerItem(40, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum))
        hbox2.addWidget(self.ok_button)
        self.page2.addLayout(hbox2)
        
        self.ok_button.clicked.connect(self.accept)
        
        self.page2_widget = QWidget()
        self.page2_widget.setLayout(self.page2)
        self.stacked_layout.addWidget(self.page2_widget)
        
        self.setLayout(self.stacked_layout)
    
    def show_next_page(self):
        self.stacked_layout.setCurrentIndex(1)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Main Window")
        self.setGeometry(100, 100, 800, 600)
        
        self.show_welcome_dialog()
    
    def show_welcome_dialog(self):
        self.welcome_dialog = WelcomeDialog()
        self.welcome_dialog.exec()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec())