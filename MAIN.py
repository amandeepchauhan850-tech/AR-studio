import os
import sys
import tempfile
import ctypes

# Fix High DPI scaling issue so window size doesn't shrink when Qt popup opens
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from docx import Document
from docx.shared import Mm
from PIL import Image, ImageDraw, ImageTk, ImageWin
from reportlab.lib.pagesizes import A4, A5, landscape, legal, letter
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

# PDF Handling using PyMuPDF (No Poppler required)
import fitz

# PyQt5 for Custom Password Dialog
from PyQt5.QtWidgets import (QApplication, QDialog, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QWidget)
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt

# Enable Qt High DPI scaling matching system
os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"

# Windows Printers win32 API
try:
    import win32api
    import win32print
    import win32ui
except ImportError:
    win32print = None
    win32api = None
    win32ui = None


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath('.')
    return os.path.join(base_path, relative_path)


class CustomPasswordDialog(QDialog):
    def __init__(self, doc_ref, parent=None):
        super().__init__(parent)
        self.doc_ref = doc_ref
        self.setWindowTitle("Enter Password")
        self.setFixedSize(400, 410)
        self.is_authenticated = False

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

        self.img_idle = resource_path("POP 1.png")
        self.img_closed = resource_path("POP 2.png")
        self.img_peek = resource_path("POP 3.png")
        self.img_error = resource_path("POP 4.png")

        self.is_password_visible = False
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(25, 15, 25, 20)
        layout.setAlignment(Qt.AlignCenter)

        self.image_label = QLabel(self)
        self.image_label.setAlignment(Qt.AlignCenter)
        self.image_label.setFixedHeight(130)
        self.update_monkey_image(self.img_idle)
        layout.addWidget(self.image_label)

        title = QLabel("Enter Password", self)
        title.setStyleSheet("font-size: 19px; font-weight: bold; color: #0f172a; margin-top: 2px;")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        subtitle = QLabel("Please enter your password to continue", self)
        subtitle.setStyleSheet("font-size: 12px; color: #64748b; margin-bottom: 4px;")
        subtitle.setAlignment(Qt.AlignCenter)
        layout.addWidget(subtitle)

        self.error_label = QLabel("", self)
        self.error_label.setStyleSheet("font-size: 12px; font-weight: bold; color: #ef4444; margin-bottom: 6px;")
        self.error_label.setAlignment(Qt.AlignCenter)
        self.error_label.setWordWrap(True)
        self.error_label.hide()
        layout.addWidget(self.error_label)

        input_container = QWidget(self)
        input_layout = QHBoxLayout(input_container)
        input_layout.setContentsMargins(0, 0, 0, 0)
        input_layout.setSpacing(5)

        self.password_entry = QLineEdit(input_container)
        self.password_entry.setPlaceholderText("••••••••")
        self.password_entry.setEchoMode(QLineEdit.Password)
        self.password_entry.textChanged.connect(self.on_text_changed)

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

        btn_layout = QHBoxLayout()
        btn_layout.setContentsMargins(0, 12, 0, 0)
        btn_layout.setSpacing(12)

        cancel_btn = QPushButton("Cancel", self)
        cancel_btn.setObjectName("CancelBtn")
        cancel_btn.setCursor(Qt.PointingHandCursor)
        cancel_btn.clicked.connect(self.reject)

        submit_btn = QPushButton("Submit", self)
        submit_btn.setObjectName("SubmitBtn")
        submit_btn.setCursor(Qt.PointingHandCursor)
        submit_btn.clicked.connect(self.validate_and_submit)

        btn_layout.addWidget(cancel_btn)
        btn_layout.addWidget(submit_btn)
        layout.addLayout(btn_layout)

        self.setLayout(layout)

    def update_monkey_image(self, image_path):
        if os.path.exists(image_path):
            pixmap = QPixmap(image_path)
            if not pixmap.isNull():
                scaled_pixmap = pixmap.scaled(125, 125, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                self.image_label.setPixmap(scaled_pixmap)

    def on_text_changed(self, text):
        if self.error_label.isVisible():
            self.error_label.hide()

        if not text:
            self.update_monkey_image(self.img_idle)
        else:
            if not self.is_password_visible:
                self.update_monkey_image(self.img_closed)

    def toggle_password_visibility(self):
        if self.is_password_visible:
            self.password_entry.setEchoMode(QLineEdit.Password)
            self.is_password_visible = False
            if self.error_label.isVisible():
                self.update_monkey_image(self.img_error)
            elif self.password_entry.text():
                self.update_monkey_image(self.img_closed)
            else:
                self.update_monkey_image(self.img_idle)
        else:
            self.password_entry.setEchoMode(QLineEdit.Normal)
            self.is_password_visible = True
            if self.error_label.isVisible():
                self.update_monkey_image(self.img_error)
            else:
                self.update_monkey_image(self.img_peek)

    def validate_and_submit(self):
        password = self.password_entry.text()
        if self.doc_ref and self.doc_ref.authenticate(password):
            self.is_authenticated = True
            self.accept()
        else:
            self.error_label.setText("Incorrect password, please try again!")
            self.error_label.show()
            self.update_monkey_image(self.img_error)
            self.password_entry.selectAll()
            self.password_entry.setFocus()


class RoundedButton(tk.Canvas):
    def __init__(
        self,
        parent,
        text,
        gradient_colors=('#00A2FF', '#0052D4'),
        fg_color='white',
        command=None,
        font=('Segoe UI', 8, 'bold'),
        height=28,
        **kwargs,
    ):
        super().__init__(
            parent,
            height=height,
            bg=parent['bg'],
            highlightthickness=0,
            bd=0,
            cursor='hand2',
            **kwargs,
        )
        self.command = command
        self.color1, self.color2 = gradient_colors
        self.base_color1 = self.color1
        self.base_color2 = self.color2
        self.fg_color = fg_color
        self.text = text
        self.font = font
        self.height = height

        self.bind('<Configure>', self._draw)
        self.bind('<Button-1>', self._on_click)
        self.bind('<Enter>', self._on_enter)
        self.bind('<Leave>', self._on_leave)

    def _hex_to_rgb(self, hex_str):
        hex_str = hex_str.lstrip('#')
        return tuple(int(hex_str[i : i + 2], 16) for i in (0, 2, 4))

    def _rgb_to_hex(self, rgb):
        return f'#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}'

    def _draw(self, event=None):
        self.delete('all')
        w = self.winfo_width()
        h = self.height

        if w <= 0 or h <= 0:
            return

        r = h // 2
        scale = 2
        img_w, img_h = w * scale, h * scale
        r_scaled = r * scale

        pil_img = Image.new('RGBA', (img_w, img_h), (0, 0, 0, 0))
        c1 = self._hex_to_rgb(self.color1)
        c2 = self._hex_to_rgb(self.color2)

        grad_img = Image.new('RGBA', (img_w, img_h))
        grad_draw = ImageDraw.Draw(grad_img)
        for x in range(img_w):
            ratio = x / max(1, img_w - 1)
            nr = int(c1[0] + (c2[0] - c1[0]) * ratio)
            ng = int(c1[1] + (c2[1] - c1[1]) * ratio)
            nb = int(c1[2] + (c2[2] - c1[2]) * ratio)
            grad_draw.line([(x, 0), (x, img_h)], fill=(nr, ng, nb, 255))

        mask = Image.new('L', (img_w, img_h), 0)
        mask_draw = ImageDraw.Draw(mask)
        mask_draw.rounded_rectangle([0, 0, img_w, img_h], radius=r_scaled, fill=255)

        pil_img.paste(grad_img, (0, 0), mask)
        pil_img = pil_img.resize((w, h), Image.Resampling.LANCZOS)

        self.tk_gradient = ImageTk.PhotoImage(pil_img)
        self.create_image(0, 0, image=self.tk_gradient, anchor='nw')

        self.create_text(
            w / 2,
            h / 2,
            text=self.text,
            fill=self.fg_color,
            font=self.font,
            anchor='center',
        )

    def _on_enter(self, event):
        def brighten(hex_str):
            rgb = self._hex_to_rgb(hex_str)
            return self._rgb_to_hex(
                tuple(min(255, int(c * 1.15 + 10)) for c in rgb)
            )

        self.color1 = brighten(self.base_color1)
        self.color2 = brighten(self.base_color2)
        self._draw()

    def _on_leave(self, event):
        self.color1 = self.base_color1
        self.color2 = self.base_color2
        self._draw()

    def _on_click(self, event):
        if self.command:
            self.command()


class CustomToggleSwitch(tk.Canvas):
    def __init__(
        self,
        parent,
        text='',
        initial=False,
        command=None,
        height=20,
        bg_color=None,
        **kwargs,
    ):
        bg = bg_color if bg_color else parent['bg']
        super().__init__(
            parent,
            height=height,
            bg=bg,
            highlightthickness=0,
            bd=0,
            cursor='hand2',
            **kwargs,
        )
        self.text = text
        self.is_on = initial
        self.command = command
        self.height_px = height

        self.active_color = '#10B981'
        self.inactive_color = '#CBD5E1'

        self.bind('<Configure>', self._draw)
        self.bind('<Button-1>', self._on_click)

    def get(self):
        return self.is_on

    def set(self, val):
        self.is_on = bool(val)
        self._draw()

    def _draw(self, event=None):
        self.delete('all')
        w = self.winfo_width()
        h = self.height_px

        if w <= 0 or h <= 0:
            return

        switch_w = 36
        switch_h = 18
        y_pos = (h - switch_h) // 2

        scale = 2
        sw_sc, sh_sc = switch_w * scale, switch_h * scale
        radius_sc = sh_sc // 2

        pil_img = Image.new('RGBA', (sw_sc, sh_sc), (0, 0, 0, 0))
        draw = ImageDraw.Draw(pil_img)

        bg_hex = self.active_color if self.is_on else self.inactive_color
        draw.rounded_rectangle(
            [0, 0, sw_sc, sh_sc], radius=radius_sc, fill=bg_hex
        )

        thumb_radius = radius_sc - (2 * scale)
        cx = sw_sc - radius_sc if self.is_on else radius_sc
        cy = radius_sc
        draw.ellipse(
            [cx - thumb_radius, cy - thumb_radius, cx + thumb_radius, cy + thumb_radius],
            fill='white',
        )

        pil_img = pil_img.resize((switch_w, switch_h), Image.Resampling.LANCZOS)
        self.tk_switch = ImageTk.PhotoImage(pil_img)
        self.create_image(0, y_pos, image=self.tk_switch, anchor='nw')

        self.create_text(
            switch_w + 8,
            h / 2,
            text=self.text,
            fill='#1F2937',
            font=('Segoe UI', 8, 'bold'),
            anchor='w',
        )

    def _on_click(self, event):
        self.is_on = not self.is_on
        self._draw()
        if self.command:
            self.command(self.is_on)


class CustomGradientDropdown(tk.Canvas):
    def __init__(
        self,
        parent,
        values=None,
        default=None,
        command=None,
        height=28,
        **kwargs,
    ):
        super().__init__(
            parent,
            height=height,
            bg=parent['bg'],
            highlightthickness=0,
            bd=0,
            cursor='hand2',
            **kwargs,
        )

        self.values = values or []
        self.current_val = tk.StringVar(
            value=default if default is not None else (self.values[0] if self.values else '')
        )
        self.command = command
        self.height = height

        self.gradient_colors = ('#1E90FF', '#0052D4')
        self.hover_gradient_colors = ('#249BFF', '#0A63E5')

        self.popup_bg = '#FFFFFF'
        self.popup_hover = '#EEF6FF'
        self.popup_selected = '#1677F2'
        self.popup_text = '#243B64'
        self.popup_border = '#D9E5F4'
        self.popup_shadow = '#C9D7E8'

        self.popup = None
        self.popup_canvas = None
        self._popup_rows = []
        self._button_hovered = False

        self.bind('<Configure>', self._draw)
        self.bind('<Button-1>', self._toggle_menu)
        self.bind('<Enter>', self._on_button_enter)
        self.bind('<Leave>', self._on_button_leave)

        self.bind_all('<Escape>', self._escape_close, add='+')
        
        top_level = self.winfo_toplevel()
        top_level.bind('<Unmap>', lambda e: self._close_menu(), add='+')

    def get(self):
        return self.current_val.get()

    def set(self, val):
        if val in self.values or not self.values:
            self.current_val.set(val)
            self._draw()

    def _hex_to_rgb(self, hex_str):
        hex_str = hex_str.lstrip('#')
        return tuple(int(hex_str[i:i + 2], 16) for i in (0, 2, 4))

    def _draw(self, event=None):
        self.delete('all')

        w = self.winfo_width()
        h = self.height

        if w <= 0 or h <= 0:
            return

        scale = 2
        img_w, img_h = w * scale, h * scale
        radius = max(8, int(h * 0.42))
        radius_sc = radius * scale

        c1_hex, c2_hex = (
            self.hover_gradient_colors
            if self._button_hovered
            else self.gradient_colors
        )
        c1 = self._hex_to_rgb(c1_hex)
        c2 = self._hex_to_rgb(c2_hex)

        grad_img = Image.new('RGBA', (img_w, img_h), (0, 0, 0, 0))
        grad_draw = ImageDraw.Draw(grad_img)

        for x in range(img_w):
            ratio = x / max(1, img_w - 1)
            rgb = tuple(
                int(c1[i] + (c2[i] - c1[i]) * ratio)
                for i in range(3)
            )
            grad_draw.line(
                [(x, 0), (x, img_h)],
                fill=(*rgb, 255)
            )

        mask = Image.new('L', (img_w, img_h), 0)
        mask_draw = ImageDraw.Draw(mask)
        mask_draw.rounded_rectangle(
            [0, 0, img_w, img_h],
            radius=radius_sc,
            fill=255
        )

        button_img = Image.new('RGBA', (img_w, img_h), (0, 0, 0, 0))
        button_img.paste(grad_img, (0, 0), mask)
        button_img = button_img.resize((w, h), Image.Resampling.LANCZOS)

        self.tk_gradient = ImageTk.PhotoImage(button_img)
        self.create_image(0, 0, image=self.tk_gradient, anchor='nw')

        self.create_text(
            10,
            h / 2,
            text=self.current_val.get(),
            fill='white',
            font=('Segoe UI', 8, 'bold'),
            anchor='w',
        )

        arrow_x = w - 14
        arrow_y = h / 2

        if self.popup and self.popup.winfo_exists():
            self.create_line(
                arrow_x - 3, arrow_y + 2,
                arrow_x, arrow_y - 2,
                fill='white',
                width=2,
                capstyle='round'
            )
            self.create_line(
                arrow_x, arrow_y - 2,
                arrow_x + 3, arrow_y + 2,
                fill='white',
                width=2,
                capstyle='round'
            )
        else:
            self.create_line(
                arrow_x - 3, arrow_y - 2,
                arrow_x, arrow_y + 2,
                fill='white',
                width=2,
                capstyle='round'
            )
            self.create_line(
                arrow_x, arrow_y + 2,
                arrow_x + 3, arrow_y - 2,
                fill='white',
                width=2,
                capstyle='round'
            )

    def _on_button_enter(self, event=None):
        self._button_hovered = True
        self._draw()

    def _on_button_leave(self, event=None):
        self._button_hovered = False
        self._draw()

    def _toggle_menu(self, event=None):
        if self.popup and self.popup.winfo_exists():
            self._close_menu()
        else:
            self._open_menu()

    def _open_menu(self):
        if not self.values:
            return

        self._close_menu()

        popup_width = max(self.winfo_width(), 160)
        row_height = max(34, self.height + 6)
        padding = 4
        popup_height = len(self.values) * row_height + padding * 2

        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.height + 4

        screen_bottom = self.winfo_screenheight()
        if y + popup_height > screen_bottom - 10:
            y = self.winfo_rooty() - popup_height - 4
            if y < 5:
                y = 5

        self.popup = tk.Toplevel(self)
        self.popup.overrideredirect(True)
        self.popup.configure(bg=self.popup_shadow)
        self.popup.geometry(f'{popup_width}x{popup_height}+{x}+{y}')

        try:
            self.popup.attributes('-topmost', True)
        except tk.TclError:
            pass

        self.popup_canvas = tk.Canvas(
            self.popup,
            width=popup_width,
            height=popup_height,
            bg=self.popup_shadow,
            highlightthickness=0,
            bd=0
        )
        self.popup_canvas.pack(fill='both', expand=True)

        self._rounded_rectangle(
            self.popup_canvas,
            2, 2,
            popup_width - 2,
            popup_height - 2,
            radius=10,
            fill=self.popup_bg,
            outline=self.popup_border,
            width=1
        )

        self._popup_rows = []

        for index, value in enumerate(self.values):
            y1 = padding + index * row_height
            y2 = y1 + row_height

            selected = (value == self.current_val.get())

            if selected:
                row_fill = self.popup_selected
                text_fill = 'white'
            else:
                row_fill = self.popup_bg
                text_fill = self.popup_text

            self._rounded_rectangle(
                self.popup_canvas,
                5, y1,
                popup_width - 5, y2,
                radius=6,
                fill=row_fill,
                outline=''
            )

            self.popup_canvas.create_text(
                18,
                (y1 + y2) / 2,
                text='▣',
                fill=text_fill,
                font=('Segoe UI Symbol', 9),
                tags=('row', f'row_{index}')
            )

            self.popup_canvas.create_text(
                36,
                (y1 + y2) / 2,
                text=value,
                fill=text_fill,
                font=('Segoe UI', 8, 'bold' if selected else 'normal'),
                anchor='w',
                tags=('row', f'row_{index}')
            )

            if selected:
                self.popup_canvas.create_text(
                    popup_width - 18,
                    (y1 + y2) / 2,
                    text='✓',
                    fill='white',
                    font=('Segoe UI', 10, 'bold'),
                    tags=('row', f'row_{index}')
                )

            self._popup_rows.append((index, y1, y2, value))

        self.popup_canvas.bind('<Motion>', self._popup_motion)
        self.popup_canvas.bind('<Leave>', self._popup_leave)
        self.popup_canvas.bind('<Button-1>', self._popup_click)
        self.popup.bind('<Escape>', lambda e: self._close_menu())
        self.popup.bind('<FocusOut>', lambda e: self._close_menu())

        self._draw()

        try:
            self.popup.focus_force()
        except tk.TclError:
            pass

    def _popup_motion(self, event):
        row_index = self._row_at(event.y)
        self._redraw_popup_rows(row_index)

    def _popup_leave(self, event=None):
        self._redraw_popup_rows(-1)

    def _popup_click(self, event):
        row_index = self._row_at(event.y)
        if row_index >= 0:
            value = self._popup_rows[row_index][3]
            self._select_value(value)

    def _row_at(self, y):
        for index, y1, y2, value in self._popup_rows:
            if y1 <= y <= y2:
                return index
        return -1

    def _redraw_popup_rows(self, hover_index):
        if not self.popup_canvas:
            return

        popup_width = self.popup_canvas.winfo_width()

        self.popup_canvas.delete('row')

        for index, y1, y2, value in self._popup_rows:
            selected = (value == self.current_val.get())
            hovered = (index == hover_index)

            if selected:
                fill = self.popup_selected
                text_fill = 'white'
            elif hovered:
                fill = self.popup_hover
                text_fill = self.popup_text
            else:
                fill = self.popup_bg
                text_fill = self.popup_text

            self._rounded_rectangle(
                self.popup_canvas,
                5, y1,
                popup_width - 5, y2,
                radius=6,
                fill=fill,
                outline='',
                tags=('row', f'row_{index}')
            )

            self.popup_canvas.create_text(
                18,
                (y1 + y2) / 2,
                text='▣',
                fill=text_fill,
                font=('Segoe UI Symbol', 9),
                tags=('row', f'row_{index}')
            )

            self.popup_canvas.create_text(
                36,
                (y1 + y2) / 2,
                text=value,
                fill=text_fill,
                font=('Segoe UI', 8, 'bold' if selected else 'normal'),
                anchor='w',
                tags=('row', f'row_{index}')
            )

            if selected:
                self.popup_canvas.create_text(
                    popup_width - 18,
                    (y1 + y2) / 2,
                    text='✓',
                    fill='white',
                    font=('Segoe UI', 10, 'bold'),
                    tags=('row', f'row_{index}')
                )

    def _select_value(self, value):
        self.current_val.set(value)
        self._close_menu()
        self._draw()

        if self.command:
            self.command(value)

    def _close_menu(self):
        if self.popup is not None:
            try:
                self.popup.destroy()
            except tk.TclError:
                pass

        self.popup = None
        self.popup_canvas = None
        self._popup_rows = []
        self._draw()

    def _escape_close(self, event=None):
        if self.popup and self.popup.winfo_exists():
            self._close_menu()

    def _rounded_rectangle(
        self,
        canvas,
        x1, y1, x2, y2,
        radius=8,
        fill='white',
        outline='',
        width=1,
        tags=None
    ):
        canvas.create_rectangle(
            x1 + radius,
            y1,
            x2 - radius,
            y2,
            fill=fill,
            outline=outline,
            width=width,
            tags=tags
        )
        canvas.create_rectangle(
            x1,
            y1 + radius,
            x2,
            y2 - radius,
            fill=fill,
            outline=outline,
            width=width,
            tags=tags
        )

        for box, start in [
            ((x1, y1, x1 + 2 * radius, y1 + 2 * radius), 90),
            ((x2 - 2 * radius, y1, x2, y1 + 2 * radius), 0),
            ((x1, y2 - 2 * radius, x1 + 2 * radius, y2), 180),
            ((x2 - 2 * radius, y2 - 2 * radius, x2, y2), 270),
        ]:
            canvas.create_arc(
                *box,
                start=start,
                extent=90,
                fill=fill,
                outline=outline,
                width=width,
                tags=tags
            )


class CustomSlider(tk.Canvas):
    def __init__(
        self,
        parent,
        from_=0,
        to=100,
        initial=50,
        command=None,
        height=20,
        width=None,
        bg_color=None,
        **kwargs,
    ):
        bg = bg_color if bg_color else parent['bg']
        super().__init__(
            parent,
            height=height,
            width=width if width else 100,
            bg=bg,
            highlightthickness=0,
            bd=0,
            cursor='hand2',
            **kwargs,
        )
        self.from_ = from_
        self.to = to
        self.value = initial
        self.command = command
        self.height_px = height

        self.grad_start = '#7B2CBF'
        self.grad_end = '#00F5D4'
        self.track_bg = '#CBD5E1'

        self.bind('<Configure>', self._draw)
        self.bind('<Button-1>', self._on_click)
        self.bind('<B1-Motion>', self._on_drag)

    def set(self, val):
        self.value = max(self.from_, min(self.to, val))
        self._draw()

    def get(self):
        return self.value

    def _hex_to_rgb(self, hex_str):
        hex_str = hex_str.lstrip('#')
        return tuple(int(hex_str[i : i + 2], 16) for i in (0, 2, 4))

    def _draw(self, event=None):
        self.delete('all')
        w = self.winfo_width()
        h = self.height_px

        if w <= 0 or h <= 0:
            return

        padding = 10
        track_h = 8
        track_y = (h - track_h) // 2
        track_w = w - (padding * 2)

        if track_w <= 0:
            return

        ratio = (self.value - self.from_) / float(self.to - self.from_)
        filled_w = int(track_w * ratio)

        scale = 2
        img_w, img_h = w * scale, h * scale
        p_sc = padding * scale
        th_sc = track_h * scale
        ty_sc = track_y * scale
        tw_sc = track_w * scale
        fw_sc = filled_w * scale

        pil_img = Image.new('RGBA', (img_w, img_h), (0, 0, 0, 0))

        track_bg_img = Image.new('RGBA', (img_w, img_h), (0, 0, 0, 0))
        t_draw = ImageDraw.Draw(track_bg_img)
        t_draw.rounded_rectangle(
            [p_sc, ty_sc, p_sc + tw_sc, ty_sc + th_sc],
            radius=th_sc // 2,
            fill=self.track_bg,
        )

        grad_img = Image.new('RGBA', (img_w, img_h), (0, 0, 0, 0))
        g_draw = ImageDraw.Draw(grad_img)

        c1 = self._hex_to_rgb(self.grad_start)
        c2 = self._hex_to_rgb(self.grad_end)

        for x in range(p_sc, p_sc + tw_sc):
            r_local = (x - p_sc) / float(tw_sc)
            nr = int(c1[0] + (c2[0] - c1[0]) * r_local)
            ng = int(c1[1] + (c2[1] - c1[1]) * r_local)
            nb = int(c1[2] + (c2[2] - c1[2]) * r_local)
            g_draw.line([(x, ty_sc), (x, ty_sc + th_sc)], fill=(nr, ng, nb, 255))

        mask = Image.new('L', (img_w, img_h), 0)
        m_draw = ImageDraw.Draw(mask)
        if fw_sc > 0:
            m_draw.rounded_rectangle(
                [p_sc, ty_sc, p_sc + max(th_sc, fw_sc), ty_sc + th_sc],
                radius=th_sc // 2,
                fill=255,
            )

        pil_img.paste(track_bg_img, (0, 0), track_bg_img)
        pil_img.paste(grad_img, (0, 0), mask)

        pil_img = pil_img.resize((w, h), Image.Resampling.LANCZOS)
        self.tk_track = ImageTk.PhotoImage(pil_img)
        self.create_image(0, 0, image=self.tk_track, anchor='nw')

        thumb_x = padding + filled_w
        thumb_y = h // 2
        r_outer = 7
        r_inner = 4

        self.create_oval(
            thumb_x - r_outer,
            thumb_y - r_outer,
            thumb_x + r_outer,
            thumb_y + r_outer,
            fill='#00F5D4',
            outline='',
        )
        self.create_oval(
            thumb_x - r_inner,
            thumb_y - r_inner,
            thumb_x + r_inner,
            thumb_y + r_inner,
            fill='white',
            outline='',
        )

    def _update_val(self, x):
        w = self.winfo_width()
        padding = 10
        track_w = w - (padding * 2)

        if track_w <= 0:
            return

        rel_x = max(0, min(track_w, x - padding))
        val_ratio = rel_x / float(track_w)
        self.value = self.from_ + (self.to - self.from_) * val_ratio
        self._draw()

        if self.command:
            self.command(self.value)

    def _on_click(self, event):
        self._update_val(event.x)

    def _on_drag(self, event):
        self._update_val(event.x)


class ARCardStudioApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AR Studio — ID & Passport Photo Master Pro")
        
        self.root.state('zoomed')
        self.root.geometry(f"{self.root.winfo_screenwidth()}x{self.root.winfo_screenheight()}+0+0")

        self.PALETTE = {
            'bg_app': '#F3F4F6',
            'bg_sidebar': '#FFFFFF',
            'bg_card': '#FAFAFA',
            'text_primary': '#1F2937',
            'text_secondary': '#6B7280',
            'border': '#E5E7EB',
            'canvas_bg': '#E5E7EB',
        }

        self.root.configure(bg=self.PALETTE['bg_app'])

        ico_path = resource_path('logo.ico')
        png_path = resource_path('logo.png')

        if os.path.exists(ico_path):
            try:
                self.root.iconbitmap(ico_path)
            except Exception:
                pass

        self.PAGE_SIZES = {
            'A4': (210, 297),
            'A5': (148, 210),
            'Letter': (215.9, 279.4),
            'Legal': (215.9, 355.6),
        }

        self.PHOTO_SIZES = {
            'Passport size photo (33x43 mm)': (33.0, 43.0),
            'Stamp size photo (20x25 mm)': (20.0, 25.0),
        }

        self.id_slots = {}
        self.id_slot_rects = []

        self.app_mode = 'ID_CARD'
        self.pp_image_path = None
        self.pp_image_obj = None

        self.setup_ui(png_path)

    def play_click_sound(self):
        pass

    def get_current_page_dim(self):
        page_name = self.combo_page.get()
        w, h = self.PAGE_SIZES[page_name]
        if self.combo_orient.get() == 'Landscape':
            return h, w
        return w, h

    def get_current_photo_dim(self):
        size_name = self.combo_photo_size.get()
        return self.PHOTO_SIZES.get(size_name, (33.0, 43.0))

    def switch_mode(self, mode):
        self.play_click_sound()
        self.app_mode = mode
        if mode == 'ID_CARD':
            self.btn_mode_id._draw()
            self.pp_tools_frame.pack_forget()
            self.id_tools_frame.pack(fill='x')
            self.photo_type_frame.pack_forget()
            
            self.slider_id_rows.set(1)
            self.slider_id_cols.set(2)
            self.slider_margin_left.set(11)
            self.slider_margin_top.set(5)
            self.slider_gap_y.set(6)
            self.slider_gap_x.set(12)
            self.on_id_slider_change(None)
            self.on_slider_change(None)
        else:
            self.btn_mode_pp._draw()
            self.id_tools_frame.pack_forget()
            self.pp_tools_frame.pack(fill='x')
            self.photo_type_frame.pack(fill='x', pady=(2, 0))

            self.apply_photo_type_defaults(self.combo_photo_size.get())

        self.draw_preview()

    def apply_photo_type_defaults(self, photo_type):
        if photo_type == 'Passport size photo (33x43 mm)':
            self.slider_pp_rows.set(7)
            self.slider_pp_cols.set(6)
            self.slider_margin_left.set(4)
            self.slider_margin_top.set(4)
            self.slider_gap_y.set(3)
            self.slider_gap_x.set(1)
        elif photo_type == 'Stamp size photo (20x25 mm)':
            self.slider_pp_rows.set(11)
            self.slider_pp_cols.set(9)
            self.slider_margin_left.set(4)
            self.slider_margin_top.set(4)
            self.slider_gap_y.set(1)
            self.slider_gap_x.set(3)

        self.on_pp_slider_change(None)
        self.on_slider_change(None)

    def on_photo_type_change(self, val):
        self.play_click_sound()
        if self.app_mode == 'PASSPORT':
            self.apply_photo_type_defaults(val)
        self.draw_preview()

    def show_password_popup(self, doc_ref):
        qapp = QApplication.instance()
        if qapp is None:
            qapp = QApplication(sys.argv)

        dialog = CustomPasswordDialog(doc_ref=doc_ref)
        result = dialog.exec_()

        self.root.state('zoomed')
        self.root.update_idletasks()

        return dialog.is_authenticated

    def upload_aadhaar_card(self):
        file = filedialog.askopenfilename(
            title="Select Aadhaar Card File (PDF or Image)",
            filetypes=[('Supported Files', '*.pdf *.png *.jpg *.jpeg *.bmp'), ('PDF Files', '*.pdf'), ('Image Files', '*.png *.jpg *.jpeg *.bmp')]
        )
        if not file:
            return

        try:
            img = None
            if file.lower().endswith('.pdf'):
                doc = fitz.open(file)
                if doc.is_encrypted:
                    authenticated = self.show_password_popup(doc)
                    if not authenticated:
                        doc.close()
                        return
                
                page = doc[0]
                pix = page.get_pixmap(dpi=300)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                doc.close()
            else:
                img = Image.open(file)

            w, h = img.size

            if file.lower().endswith('.pdf'):
                crop_y1 = int(h * 0.725)
                crop_y2 = int(h * 0.926)
                crop_x1 = int(w * 0.08)
                crop_x2 = int(w * 0.93)
                bottom_section = img.crop((crop_x1, crop_y1, crop_x2, crop_y2))
                bw, bh = bottom_section.size
                
                front_crop = bottom_section.crop((0, 0, int(bw * 0.48), bh))
                back_crop = bottom_section.crop((int(bw * 0.505), 0, int(bw*0.99), bh))
            else:
                if h > w:
                    front_crop = img.crop((0, int(h * 0.5), w, int(h * 0.75)))
                    back_crop = img.crop((0, int(h * 0.75), w, h))
                else:
                    front_crop = img.crop((0, 0, int(w * 0.48), h))
                    back_crop = img.crop((int(w * 0.505), 0, w, h))

            temp_dir = tempfile.gettempdir()
            front_path = os.path.join(temp_dir, "aadhaar_front_temp.png")
            back_path = os.path.join(temp_dir, "aadhaar_back_temp.png")

            front_crop.save(front_path)
            back_crop.save(back_path)

            self.id_slots[1] = front_path
            self.id_slots[2] = back_path
            self.draw_preview()

        except Exception as e:
            messagebox.showerror("Error", f"Could not process Aadhaar Card File:\n{str(e)}")

    def upload_voter_card(self):
        file = filedialog.askopenfilename(
            title="Select Voter Card File (PDF or Image)",
            filetypes=[('Supported Files', '*.pdf *.png *.jpg *.jpeg *.bmp'), ('PDF Files', '*.pdf'), ('Image Files', '*.png *.jpg *.jpeg *.bmp')]
        )
        if not file:
            return

        try:
            img = None
            if file.lower().endswith('.pdf'):
                doc = fitz.open(file)
                if doc.is_encrypted:
                    authenticated = self.show_password_popup(doc)
                    if not authenticated:
                        doc.close()
                        return
                
                page = doc[0]
                pix = page.get_pixmap(dpi=300)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                doc.close()
            else:
                img = Image.open(file)

            w, h = img.size

            if file.lower().endswith('.pdf'):
                crop_y1 = int(h * 0.113)
                crop_y2 = int(h * 0.295)
            
                front_x1 = int(w * 0.055)
                front_x2 = int(w * 0.468)
                front_crop = img.crop((front_x1, crop_y1, front_x2, crop_y2))

                back_x1 = int(w * 0.550)
                back_x2 = int(w * 0.958)
                back_crop = img.crop((back_x1, crop_y1, back_x2, crop_y2))
            else:
                if h > w:
                    front_crop = img.crop((0, 0, w, int(h * 0.5)))
                    back_crop = img.crop((0, int(h * 0.5), w, h))
                else:
                    front_crop = img.crop((0, 0, int(w * 0.50), h))
                    back_crop = img.crop((int(w * 0.50), 0, w, h))

            temp_dir = tempfile.gettempdir()
            front_path = os.path.join(temp_dir, "voter_front_temp.png")
            back_path = os.path.join(temp_dir, "voter_back_temp.png")

            front_crop.save(front_path)
            back_crop.save(back_path)

            self.id_slots[1] = front_path
            self.id_slots[2] = back_path
            self.draw_preview()

        except Exception as e:
            messagebox.showerror("Error", f"Could not process Voter Card File:\n{str(e)}")

    def setup_ui(self, logo_path):
        main_frame = tk.Frame(self.root, bg=self.PALETTE['bg_app'])
        main_frame.pack(fill='both', expand=True, padx=4, pady=4)

        sidebar = tk.Frame(
            main_frame,
            bg=self.PALETTE['bg_sidebar'],
            width=380,
            highlightbackground=self.PALETTE['border'],
            highlightthickness=1,
        )
        sidebar.pack(side='left', fill='y', padx=(0, 4))
        sidebar.pack_propagate(False)

        toggle_bar = tk.Frame(sidebar, bg=self.PALETTE['bg_sidebar'], pady=1)
        toggle_bar.pack(fill='x', padx=6, pady=(4, 2))

        toggle_bar.columnconfigure(0, weight=1, uniform='toggle_group')
        toggle_bar.columnconfigure(1, weight=1, uniform='toggle_group')

        self.btn_mode_id = RoundedButton(
            toggle_bar,
            text='ID Card Mode',
            gradient_colors=('#1E90FF', '#0052D4'),
            height=30,
            font=('Segoe UI', 8, 'bold'),
            command=lambda: self.switch_mode('ID_CARD'),
        )
        self.btn_mode_id.grid(row=0, column=0, padx=2, sticky='ew')

        self.btn_mode_pp = RoundedButton(
            toggle_bar,
            text='Photo Mode',
            gradient_colors=('#7B2CBF', '#F72585'),
            height=30,
            font=('Segoe UI', 8, 'bold'),
            command=lambda: self.switch_mode('PASSPORT'),
        )
        self.btn_mode_pp.grid(row=0, column=1, padx=2, sticky='ew')

        sb_content = tk.Frame(sidebar, bg=self.PALETTE['bg_sidebar'])
        sb_content.pack(fill='both', expand=True, padx=6, pady=1)

        self.tools_container = tk.Frame(sb_content, bg=self.PALETTE['bg_sidebar'])
        self.tools_container.pack(fill='x')

        self.id_tools_frame = tk.Frame(
            self.tools_container, bg=self.PALETTE['bg_sidebar']
        )
        self.id_tools_frame.pack(fill='x')

        tk.Label(
            self.id_tools_frame,
            text='1. ID CARD LAYOUT GRID',
            font=('Segoe UI', 8, 'bold'),
            fg='#0052D4',
            bg=self.PALETTE['bg_sidebar'],
        ).pack(anchor='w', pady=(1, 0))

        id_grid_box = tk.Frame(
            self.id_tools_frame,
            bg=self.PALETTE['bg_card'],
            bd=1,
            relief='solid',
            highlightbackground=self.PALETTE['border'],
            padx=4,
            pady=2,
        )
        id_grid_box.pack(fill='x', pady=1)

        btn_aadhaar_upload = RoundedButton(
            id_grid_box,
            text='Load Aadhaar Card (Auto Split Front/Back)',
            gradient_colors=('#FF512F', '#DD2476'),
            height=28,
            command=self.upload_aadhaar_card,
        )
        btn_aadhaar_upload.pack(fill='x', pady=(2, 2))

        btn_voter_upload = RoundedButton(
            id_grid_box,
            text='Load Voter Card (Auto Split Front/Back)',
            gradient_colors=('#8E2DE2', '#4A00E0'),
            height=28,
            command=self.upload_voter_card,
        )
        btn_voter_upload.pack(fill='x', pady=(2, 4))

        self.lbl_id_rows = tk.Label(
            id_grid_box,
            text='ID Card Rows (1 - 10): 1',
            font=('Segoe UI', 8, 'bold'),
            fg=self.PALETTE['text_primary'],
            bg=self.PALETTE['bg_card'],
        )
        self.lbl_id_rows.pack(anchor='w')
        self.slider_id_rows = CustomSlider(
            id_grid_box,
            from_=1,
            to=10,
            initial=1,
            bg_color=self.PALETTE['bg_card'],
            command=self.on_id_slider_change,
        )
        self.slider_id_rows.pack(fill='x', pady=1)

        self.lbl_id_cols = tk.Label(
            id_grid_box,
            text='ID Card Columns (1 - 10): 2',
            font=('Segoe UI', 8, 'bold'),
            fg=self.PALETTE['text_primary'],
            bg=self.PALETTE['bg_card'],
        )
        self.lbl_id_cols.pack(anchor='w')
        self.slider_id_cols = CustomSlider(
            id_grid_box,
            from_=1,
            to=10,
            initial=2,
            bg_color=self.PALETTE['bg_card'],
            command=self.on_id_slider_change,
        )
        self.slider_id_cols.pack(fill='x', pady=1)

        self.pp_tools_frame = tk.Frame(
            self.tools_container, bg=self.PALETTE['bg_sidebar']
        )

        tk.Label(
            self.pp_tools_frame,
            text='1. PHOTO SETTINGS',
            font=('Segoe UI', 8, 'bold'),
            fg='#0052D4',
            bg=self.PALETTE['bg_sidebar'],
        ).pack(anchor='w', pady=(1, 0))

        pp_box = tk.Frame(
            self.pp_tools_frame,
            bg=self.PALETTE['bg_card'],
            bd=1,
            relief='solid',
            highlightbackground=self.PALETTE['border'],
            padx=4,
            pady=2,
        )
        pp_box.pack(fill='x', pady=1)

        btn_pp_upload = RoundedButton(
            pp_box,
            text='Load Photo (Auto-Fill Sheet)',
            gradient_colors=('#00C9FF', '#92FE9D'),
            fg_color='#1F2937',
            height=28,
            command=self.upload_pp_photo,
        )
        btn_pp_upload.pack(fill='x', pady=2)

        self.lbl_pp_rows = tk.Label(
            pp_box,
            text='Photo Rows (1 - 15): 7',
            font=('Segoe UI', 8, 'bold'),
            fg=self.PALETTE['text_primary'],
            bg=self.PALETTE['bg_card'],
        )
        self.lbl_pp_rows.pack(anchor='w')
        self.slider_pp_rows = CustomSlider(
            pp_box,
            from_=1,
            to=15,
            initial=7,
            bg_color=self.PALETTE['bg_card'],
            command=self.on_pp_slider_change,
        )
        self.slider_pp_rows.pack(fill='x', pady=1)

        self.lbl_pp_cols = tk.Label(
            pp_box,
            text='Photo Columns (1 - 15): 6',
            font=('Segoe UI', 8, 'bold'),
            fg=self.PALETTE['text_primary'],
            bg=self.PALETTE['bg_card'],
        )
        self.lbl_pp_cols.pack(anchor='w')
        self.slider_pp_cols = CustomSlider(
            pp_box,
            from_=1,
            to=15,
            initial=6,
            bg_color=self.PALETTE['bg_card'],
            command=self.on_pp_slider_change,
        )
        self.slider_pp_cols.pack(fill='x', pady=1)

        tk.Label(
            sb_content,
            text='2. PAGE & ORIENTATION',
            font=('Segoe UI', 8, 'bold'),
            fg='#0052D4',
            bg=self.PALETTE['bg_sidebar'],
        ).pack(anchor='w', pady=(1, 0))

        config_box = tk.Frame(
            sb_content,
            bg=self.PALETTE['bg_card'],
            bd=1,
            relief='solid',
            highlightbackground=self.PALETTE['border'],
            padx=4,
            pady=2,
        )
        config_box.pack(fill='x', pady=1)

        dropdown_grid = tk.Frame(config_box, bg=self.PALETTE['bg_card'])
        dropdown_grid.pack(fill='x')

        tk.Label(
            dropdown_grid,
            text='Page Size:',
            font=('Segoe UI', 8, 'bold'),
            fg=self.PALETTE['text_secondary'],
            bg=self.PALETTE['bg_card'],
        ).grid(row=0, column=0, sticky='w', padx=(0, 2))

        self.combo_page = CustomGradientDropdown(
            dropdown_grid,
            values=list(self.PAGE_SIZES.keys()),
            default='A4',
            height=26,
            command=lambda val: [self.play_click_sound(), self.draw_preview()],
        )
        self.combo_page.grid(row=0, column=1, sticky='ew', padx=(0, 4))

        tk.Label(
            dropdown_grid,
            text='Orient:',
            font=('Segoe UI', 8, 'bold'),
            fg=self.PALETTE['text_secondary'],
            bg=self.PALETTE['bg_card'],
        ).grid(row=0, column=2, sticky='w', padx=(0, 2))

        self.combo_orient = CustomGradientDropdown(
            dropdown_grid,
            values=['Portrait', 'Landscape'],
            default='Portrait',
            height=26,
            command=lambda val: [self.play_click_sound(), self.draw_preview()],
        )
        self.combo_orient.grid(row=0, column=3, sticky='ew')

        dropdown_grid.columnconfigure(1, weight=1)
        dropdown_grid.columnconfigure(3, weight=1)

        self.photo_type_frame = tk.Frame(config_box, bg=self.PALETTE['bg_card'])
        
        tk.Label(
            self.photo_type_frame,
            text='Photo Type:',
            font=('Segoe UI', 8, 'bold'),
            fg=self.PALETTE['text_secondary'],
            bg=self.PALETTE['bg_card'],
        ).pack(anchor='w', pady=(2, 0))

        self.combo_photo_size = CustomGradientDropdown(
            self.photo_type_frame,
            values=list(self.PHOTO_SIZES.keys()),
            default='Passport size photo (33x43 mm)',
            height=26,
            command=self.on_photo_type_change,
        )
        self.combo_photo_size.pack(fill='x', pady=(0, 2))

        tk.Label(
            sb_content,
            text='3. PAGE MARGINS',
            font=('Segoe UI', 8, 'bold'),
            fg='#0052D4',
            bg=self.PALETTE['bg_sidebar'],
        ).pack(anchor='w', pady=(1, 0))

        margin_box = tk.Frame(
            sb_content,
            bg=self.PALETTE['bg_card'],
            bd=1,
            relief='solid',
            highlightbackground=self.PALETTE['border'],
            padx=4,
            pady=2,
        )
        margin_box.pack(fill='x', pady=1)

        self.lbl_margin_left = tk.Label(
            margin_box,
            text='Left Margin: 11 mm',
            font=('Segoe UI', 8, 'bold'),
            fg=self.PALETTE['text_primary'],
            bg=self.PALETTE['bg_card'],
        )
        self.lbl_margin_left.pack(anchor='w')
        self.slider_margin_left = CustomSlider(
            margin_box,
            from_=0,
            to=50,
            initial=11,
            bg_color=self.PALETTE['bg_card'],
            command=self.on_slider_change,
        )
        self.slider_margin_left.pack(fill='x', pady=1)

        self.lbl_margin_top = tk.Label(
            margin_box,
            text='Top Margin: 5 mm',
            font=('Segoe UI', 8, 'bold'),
            fg=self.PALETTE['text_primary'],
            bg=self.PALETTE['bg_card'],
        )
        self.lbl_margin_top.pack(anchor='w')
        self.slider_margin_top = CustomSlider(
            margin_box,
            from_=0,
            to=50,
            initial=5,
            bg_color=self.PALETTE['bg_card'],
            command=self.on_slider_change,
        )
        self.slider_margin_top.pack(fill='x', pady=1)

        tk.Label(
            sb_content,
            text='4. GAP BETWEEN CARDS/PHOTOS',
            font=('Segoe UI', 8, 'bold'),
            fg='#0052D4',
            bg=self.PALETTE['bg_sidebar'],
        ).pack(anchor='w', pady=(1, 0))

        gap_box = tk.Frame(
            sb_content,
            bg=self.PALETTE['bg_card'],
            bd=1,
            relief='solid',
            highlightbackground=self.PALETTE['border'],
            padx=4,
            pady=2,
        )
        gap_box.pack(fill='x', pady=1)

        self.lbl_gap_y = tk.Label(
            gap_box,
            text='Row Gap Y: 6 mm',
            font=('Segoe UI', 8, 'bold'),
            fg=self.PALETTE['text_primary'],
            bg=self.PALETTE['bg_card'],
        )
        self.lbl_gap_y.pack(anchor='w')
        self.slider_gap_y = CustomSlider(
            gap_box,
            from_=0,
            to=20,
            initial=6,
            bg_color=self.PALETTE['bg_card'],
            command=self.on_slider_change,
        )
        self.slider_gap_y.pack(fill='x', pady=1)

        self.lbl_gap_x = tk.Label(
            gap_box,
            text='Column Gap X: 12 mm',
            font=('Segoe UI', 8, 'bold'),
            fg=self.PALETTE['text_primary'],
            bg=self.PALETTE['bg_card'],
        )
        self.lbl_gap_x.pack(anchor='w')
        self.slider_gap_x = CustomSlider(
            gap_box,
            from_=0,
            to=20,
            initial=12,
            bg_color=self.PALETTE['bg_card'],
            command=self.on_slider_change,
        )
        self.slider_gap_x.pack(fill='x', pady=1)

        self.toggle_fold = CustomToggleSwitch(
            gap_box,
            text='Cutting Lines',
            initial=False,
            bg_color=self.PALETTE['bg_card'],
            command=lambda val: [self.play_click_sound(), self.draw_preview()],
        )
        self.toggle_fold.pack(fill='x', pady=(2, 1))

        self.toggle_border = CustomToggleSwitch(
            gap_box,
            text='Border Lines',
            initial=True,
            bg_color=self.PALETTE['bg_card'],
            command=lambda val: [self.play_click_sound(), self.draw_preview()],
        )
        self.toggle_border.pack(fill='x', pady=(0, 1))

        tk.Label(
            sb_content,
            text='5. SAVE & PRINT OPTIONS',
            font=('Segoe UI', 8, 'bold'),
            fg='#0052D4',
            bg=self.PALETTE['bg_sidebar'],
        ).pack(anchor='w', pady=(1, 0))

        export_grid = tk.Frame(sb_content, bg=self.PALETTE['bg_sidebar'])
        export_grid.pack(fill='x', pady=1)

        btn_pdf = RoundedButton(
            export_grid,
            text='Save PDF',
            gradient_colors=('#00B4DB', '#0083B0'),
            height=26,
            command=lambda: self.export_sheet('pdf'),
        )
        btn_pdf.grid(row=0, column=0, padx=2, pady=1, sticky='ew')

        btn_jpg = RoundedButton(
            export_grid,
            text='Save JPG',
            gradient_colors=('#FF8008', '#FFC837'),
            height=26,
            command=lambda: self.export_sheet('jpg'),
        )
        btn_jpg.grid(row=0, column=1, padx=2, pady=1, sticky='ew')

        btn_png = RoundedButton(
            export_grid,
            text='Save PNG',
            gradient_colors=('#00A2FF', '#0052D4'),
            height=26,
            command=lambda: self.export_sheet('png'),
        )
        btn_png.grid(row=1, column=0, padx=2, pady=1, sticky='ew')

        btn_word = RoundedButton(
            export_grid,
            text='Word (.docx)',
            gradient_colors=('#8A2387', '#E94057'),
            height=26,
            command=self.save_word,
        )
        btn_word.grid(row=1, column=1, padx=2, pady=1, sticky='ew')

        btn_print = RoundedButton(
            export_grid,
            text='Print Sheet (Select Printer)',
            gradient_colors=('#11998e', '#38ef7d'),
            height=28,
            command=self.open_printer_dialog,
        )
        btn_print.grid(row=2, column=0, columnspan=2, padx=2, pady=1, sticky='ew')

        export_grid.columnconfigure(0, weight=1)
        export_grid.columnconfigure(1, weight=1)

        btn_clear = RoundedButton(
            sb_content,
            text='Clear All',
            gradient_colors=('#FF416C', '#FF4B2B'),
            height=26,
            command=self.clear_all,
        )
        btn_clear.pack(fill='x', pady=(2, 0))

        preview_box = tk.Frame(
            main_frame,
            bg=self.PALETTE['bg_sidebar'],
            highlightbackground=self.PALETTE['border'],
            highlightthickness=1,
        )
        preview_box.pack(side='right', fill='both', expand=True)

        zoom_bar = tk.Frame(preview_box, bg=self.PALETTE['bg_sidebar'], pady=2)
        zoom_bar.pack(fill='x', padx=6)

        self.preview_header = tk.Label(
            zoom_bar,
            text='Live Print Sheet Preview',
            font=('Segoe UI', 9, 'bold'),
            fg=self.PALETTE['text_primary'],
            bg=self.PALETTE['bg_sidebar'],
        )
        self.preview_header.pack(side='left')

        zoom_controls = tk.Frame(zoom_bar, bg=self.PALETTE['bg_sidebar'])
        zoom_controls.pack(side='right')

        self.lbl_zoom = tk.Label(
            zoom_controls,
            text='Zoom: 100%',
            font=('Segoe UI', 8, 'bold'),
            fg=self.PALETTE['text_primary'],
            bg=self.PALETTE['bg_sidebar'],
        )
        self.lbl_zoom.pack(side='left', padx=(0, 2))

        self.zoom_slider = CustomSlider(
            zoom_controls,
            from_=50,
            to=500,
            initial=100,
            width=180,
            bg_color=self.PALETTE['bg_sidebar'],
            command=self.on_zoom_change,
        )
        self.zoom_slider.pack(side='left', padx=4)

        btn_reset_zoom = RoundedButton(
            zoom_controls,
            text='Reset',
            gradient_colors=('#BDC3C7', '#2C3E50'),
            height=20,
            width=55,
            command=self.reset_zoom,
        )
        btn_reset_zoom.pack(side='left', padx=(2, 0))

        self.canvas_frame = tk.Frame(preview_box, bg=self.PALETTE['canvas_bg'])
        self.canvas_frame.pack(fill='both', expand=True, padx=6, pady=(0, 2))

        self.hbar = tk.Scrollbar(
            self.canvas_frame, orient='horizontal', cursor='fleur'
        )
        self.hbar.pack(side='bottom', fill='x')

        self.vbar = tk.Scrollbar(
            self.canvas_frame, orient='vertical', cursor='fleur'
        )
        self.vbar.pack(side='right', fill='y')

        self.canvas = tk.Canvas(
            self.canvas_frame,
            bg=self.PALETTE['canvas_bg'],
            bd=0,
            highlightthickness=0,
            xscrollcommand=self.hbar.set,
            yscrollcommand=self.vbar.set,
            cursor='fleur',
        )
        self.canvas.pack(fill='both', expand=True)

        self.hbar.config(command=self.canvas.xview)
        self.vbar.config(command=self.canvas.yview)

        self.canvas.bind('<ButtonPress-1>', self.on_pan_start)
        self.canvas.bind('<B1-Motion>', self.on_pan_move)
        self.canvas.bind('<Motion>', self.on_mouse_move)
        self.canvas.bind('<MouseWheel>', self.on_mouse_wheel)
        self.canvas.bind('<Configure>', lambda e: self.draw_preview())

        footer_right = tk.Frame(preview_box, bg=self.PALETTE['bg_sidebar'], pady=2)
        footer_right.pack(side='bottom', fill='x', padx=6)

        tk.Label(
            footer_right,
            text=(
                'Powered by Aman Chauhan  |  © 2026 AR Studio Pro. All Rights'
                ' Reserved. Version 26.009.21'
            ),
            font=('Segoe UI', 8, 'bold'),
            fg='#1D4ED8',
            bg=self.PALETTE['bg_sidebar'],
            anchor='e',
        ).pack(side='top', fill='x')

        self.switch_mode('ID_CARD')

    def upload_pp_photo(self):
        file = filedialog.askopenfilename(
            filetypes=[('Image Files', '*.png *.jpg *.jpeg *.bmp')]
        )
        if file:
            self.pp_image_path = file
            self.pp_image_obj = Image.open(file)
            self.draw_preview()

    def on_mouse_wheel(self, event):
        current_zoom = self.zoom_slider.get()
        if event.delta > 0:
            new_zoom = min(500, current_zoom + 10)
        else:
            new_zoom = max(50, current_zoom - 10)
        
        self.zoom_slider.set(new_zoom)
        self.on_zoom_change(new_zoom)

    def on_mouse_move(self, event):
        if self.app_mode == 'ID_CARD':
            mx = self.canvas.canvasx(event.x)
            my = self.canvas.canvasy(event.y)
            for (
                slot_idx,
                btn_x1,
                btn_y1,
                btn_x2,
                btn_y2,
                slot_x1,
                slot_y1,
                slot_x2,
                slot_y2,
            ) in self.id_slot_rects:
                if (
                    slot_idx in self.id_slots
                    and btn_x1 <= mx <= btn_x2
                    and btn_y1 <= my <= btn_y2
                ):
                    self.canvas.config(cursor='hand2')
                    return
        self.canvas.config(cursor='fleur')

    def on_pan_start(self, event):
        click_x = self.canvas.canvasx(event.x)
        click_y = self.canvas.canvasy(event.y)

        if self.app_mode == 'ID_CARD':
            for (
                slot_idx,
                btn_x1,
                btn_y1,
                btn_x2,
                btn_y2,
                slot_x1,
                slot_y1,
                slot_x2,
                slot_y2,
            ) in self.id_slot_rects:
                if (
                    slot_idx in self.id_slots
                    and btn_x1 <= click_x <= btn_x2
                    and btn_y1 <= click_y <= btn_y2
                ):
                    del self.id_slots[slot_idx]
                    self.draw_preview()
                    return

                if slot_x1 <= click_x <= slot_x2 and slot_y1 <= click_y <= slot_y2:
                    file = filedialog.askopenfilename(
                        title=f'Select Image for Slot {slot_idx}',
                        filetypes=[('Image Files', '*.png *.jpg *.jpeg *.bmp')],
                    )
                    if file:
                        self.id_slots[slot_idx] = file
                        self.draw_preview()
                    return

        self.canvas.scan_mark(event.x, event.y)

    def on_pan_move(self, event):
        self.canvas.scan_dragto(event.x, event.y, gain=1)

    def on_id_slider_change(self, val):
        self.lbl_id_rows.config(
            text=f'ID Card Rows (1 - 10): {int(self.slider_id_rows.get())}'
        )
        self.lbl_id_cols.config(
            text=f'ID Card Columns (1 - 10): {int(self.slider_id_cols.get())}'
        )
        self.draw_preview()

    def on_pp_slider_change(self, val):
        self.lbl_pp_rows.config(
            text=f'Photo Rows (1 - 15): {int(self.slider_pp_rows.get())}'
        )
        self.lbl_pp_cols.config(
            text=f'Photo Columns (1 - 15): {int(self.slider_pp_cols.get())}'
        )
        self.draw_preview()

    def on_zoom_change(self, val):
        zoom_pct = int(val)
        self.lbl_zoom.config(text=f'Zoom: {zoom_pct}%')
        self.draw_preview()

    def reset_zoom(self):
        self.zoom_slider.set(100)
        self.lbl_zoom.config(text='Zoom: 100%')
        self.draw_preview()

    def on_slider_change(self, val):
        self.lbl_margin_left.config(
            text=f'Left Margin: {int(self.slider_margin_left.get())} mm'
        )
        self.lbl_margin_top.config(
            text=f'Top Margin: {int(self.slider_margin_top.get())} mm'
        )
        self.lbl_gap_y.config(text=f'Row Gap Y: {int(self.slider_gap_y.get())} mm')
        self.lbl_gap_x.config(
            text=f'Column Gap X: {int(self.slider_gap_x.get())} mm'
        )
        self.draw_preview()

    def clear_all(self):
        if self.app_mode == 'ID_CARD':
            self.id_slots.clear()
        else:
            self.pp_image_path = None
            self.pp_image_obj = None
        self.draw_preview()

    def draw_preview(self):
        self.canvas.delete('all')
        self.id_slot_rects.clear()

        p_w_mm, p_h_mm = self.get_current_page_dim()
        page_name = self.combo_page.get()
        orient = self.combo_orient.get()

        self.preview_header.config(
            text=(
                f'Live Preview [{self.app_mode.replace("_", " ")}] — {page_name}'
                f' [{orient}] ({p_w_mm:.1f} × {p_h_mm:.1f} mm)'
            )
        )

        zoom_factor = (float(self.zoom_slider.get()) / 100.0) * 1.6

        if orient == 'Portrait':
            base_h = 510
            scale_ratio = (base_h / p_h_mm) * zoom_factor
        else:
            base_w = 620
            scale_ratio = (base_w / p_w_mm) * zoom_factor

        page_w = int(p_w_mm * scale_ratio)
        page_h = int(p_h_mm * scale_ratio)

        canvas_w = max(self.canvas.winfo_width(), page_w + 100)
        canvas_h = max(self.canvas.winfo_height(), page_h + 100)

        self.canvas.config(scrollregion=(0, 0, canvas_w, canvas_h))

        page_x = max(20, (canvas_w - page_w) // 2)
        page_y = max(20, (canvas_h - page_h) // 2)

        self.canvas.create_rectangle(
            page_x + 5,
            page_y + 5,
            page_x + page_w + 5,
            page_y + page_h + 5,
            fill='#D1D5DB',
            outline='',
        )
        self.canvas.create_rectangle(
            page_x,
            page_y,
            page_x + page_w,
            page_y + page_h,
            fill='white',
            outline='#9CA3AF',
            width=1,
        )

        margin_left = int(self.slider_margin_left.get() * scale_ratio)
        margin_top = int(self.slider_margin_top.get() * scale_ratio)
        gap_y = int(self.slider_gap_y.get() * scale_ratio)
        gap_x = int(self.slider_gap_x.get() * scale_ratio)

        if self.app_mode == 'ID_CARD':
            card_w = int(88.9 * scale_ratio)
            card_h = int(50.8 * scale_ratio)

            rows = int(self.slider_id_rows.get())
            cols = int(self.slider_id_cols.get())

            start_x = page_x + margin_left
            start_y = page_y + margin_top

            if self.toggle_fold.get() and cols > 1:
                for c in range(cols - 1):
                    fold_x = start_x + (c + 1) * card_w + c * gap_x + (gap_x // 2)
                    self.canvas.create_line(
                        fold_x,
                        start_y,
                        fold_x,
                        start_y + rows * card_h + (rows - 1) * gap_y,
                        fill='#EF4444',
                        dash=(4, 4),
                        width=max(1, int(1.5 * zoom_factor)),
                    )

            slot_count = 1
            for r in range(rows):
                for c in range(cols):
                    x = start_x + c * (card_w + gap_x)
                    y = start_y + r * (card_h + gap_y)

                    if x + card_w > page_x + page_w or y + card_h > page_y + page_h:
                        continue

                    slot_idx = slot_count
                    slot_count += 1

                    img_path = self.id_slots.get(slot_idx)

                    btn_size = 18
                    btn_x1 = x + card_w - btn_size - 3
                    btn_y1 = y + 3
                    btn_x2 = x + card_w - 3
                    btn_y2 = y + btn_size + 3

                    self.id_slot_rects.append((
                        slot_idx,
                        btn_x1,
                        btn_y1,
                        btn_x2,
                        btn_y2,
                        x,
                        y,
                        x + card_w,
                        y + card_h,
                    ))

                    if img_path and os.path.exists(img_path):
                        img_p = Image.open(img_path)
                        img_p = img_p.resize(
                            (max(1, card_w), max(1, card_h)), Image.Resampling.LANCZOS
                        )
                        tk_img = ImageTk.PhotoImage(img_p)
                        setattr(self, f'id_img_cache_{slot_idx}', tk_img)
                        self.canvas.create_image(
                            x + card_w / 2, y + card_h / 2, image=tk_img
                        )
                        if self.toggle_border.get():
                            self.canvas.create_rectangle(
                                x, y, x + card_w, y + card_h, outline='#10B981', width=1.5
                            )

                        self.canvas.create_rectangle(
                            btn_x1,
                            btn_y1,
                            btn_x2,
                            btn_y2,
                            fill='#EF4444',
                            outline='white',
                            width=1,
                        )
                        self.canvas.create_text(
                            (btn_x1 + btn_x2) / 2,
                            (btn_y1 + btn_y2) / 2,
                            text='X',
                            fill='white',
                            font=('Segoe UI', 9, 'bold'),
                        )
                    else:
                        self.canvas.create_rectangle(
                            x,
                            y,
                            x + card_w,
                            y + card_h,
                            outline='#9CA3AF',
                            dash=(3, 3),
                            fill='#F9FAFB',
                        )
                        center_x = x + card_w / 2
                        center_y = y + card_h / 2
                        
                        font_slot_size = max(8, int(10 * zoom_factor))
                        font_click_size = max(7, int(8 * zoom_factor))

                        self.canvas.create_text(
                            center_x,
                            center_y - 9,
                            text=f'Slot {slot_idx}',
                            fill='#374151',
                            font=('Segoe UI', font_slot_size, 'bold'),
                        )
                        self.canvas.create_text(
                            center_x,
                            center_y + 9,
                            text='Click to Upload',
                            fill='#2563EB',
                            font=('Segoe UI', font_click_size),
                        )

        else:
            pw_mm, ph_mm = self.get_current_photo_dim()
            pp_w = int(pw_mm * scale_ratio)
            pp_h = int(ph_mm * scale_ratio)

            rows = int(self.slider_pp_rows.get())
            cols = int(self.slider_pp_cols.get())

            start_x = page_x + margin_left
            start_y = page_y + margin_top

            if self.pp_image_path and os.path.exists(self.pp_image_path):
                img_p = Image.open(self.pp_image_path)
                img_p = img_p.resize(
                    (max(1, pp_w), max(1, pp_h)), Image.Resampling.LANCZOS
                )
                tk_pp = ImageTk.PhotoImage(img_p)
                self.cache_pp_photo = tk_pp

            if self.toggle_fold.get():
                for c in range(1, cols):
                    fx = start_x + c * pp_w + (c - 1) * gap_x + (gap_x // 2)
                    self.canvas.create_line(
                        fx,
                        start_y,
                        fx,
                        start_y + rows * pp_h + (rows - 1) * gap_y,
                        fill='#EF4444',
                        dash=(4, 4),
                        width=1,
                    )
                for r in range(1, rows):
                    fy = start_y + r * pp_h + (r - 1) * gap_y + (gap_y // 2)
                    self.canvas.create_line(
                        start_x,
                        fy,
                        start_x + cols * pp_w + (cols - 1) * gap_x,
                        fy,
                        fill='#EF4444',
                        dash=(4, 4),
                        width=1,
                    )

            photo_num = 1
            is_stamp_size = (self.combo_photo_size.get() == 'Stamp size photo (20x25 mm)')
            
            for r in range(rows):
                for c in range(cols):
                    x = start_x + c * (pp_w + gap_x)
                    y = start_y + r * (pp_h + gap_y)

                    if x + pp_w > page_x + page_w or y + pp_h > page_y + page_h:
                        continue

                    if self.pp_image_path and os.path.exists(self.pp_image_path):
                        self.canvas.create_image(
                            x + pp_w / 2, y + pp_h / 2, image=self.cache_pp_photo
                        )
                        if self.toggle_border.get():
                            self.canvas.create_rectangle(
                                x, y, x + pp_w, y + pp_h, outline='#2563EB', width=1
                            )
                    else:
                        self.canvas.create_rectangle(
                            x, y, x + pp_w, y + pp_h, outline='#D1D5DB', dash=(2, 2)
                        )
                        
                        if is_stamp_size:
                            font_size = max(4, int(5 * zoom_factor))
                            text_str = f'P-{photo_num}'
                        else:
                            font_size = max(6, int(8 * zoom_factor))
                            text_str = f'Photo {photo_num}'

                        self.canvas.create_text(
                            x + pp_w / 2,
                            y + pp_h / 2,
                            text=text_str,
                            fill='#6B7280',
                            font=('Segoe UI', font_size, 'bold'),
                        )

                    photo_num += 1

    def generate_full_res_image(self, transparent=False):
        p_w_mm, p_h_mm = self.get_current_page_dim()

        scale = 300 / 25.4
        img_w, img_h = int(p_w_mm * scale), int(p_h_mm * scale)
        bg_mode = (0, 0, 0, 0) if transparent else (255, 255, 255, 255)
        sheet_img = Image.new('RGBA', (img_w, img_h), bg_mode)
        draw = ImageDraw.Draw(sheet_img)

        margin_left = int(self.slider_margin_left.get() * scale)
        margin_top = int(self.slider_margin_top.get() * scale)
        gap_y = int(self.slider_gap_y.get() * scale)
        gap_x = int(self.slider_gap_x.get() * scale)

        if self.app_mode == 'ID_CARD':
            card_w, card_h = int(88.9 * scale), int(50.8 * scale)
            rows = int(self.slider_id_rows.get())
            cols = int(self.slider_id_cols.get())

            slot_count = 1
            for r in range(rows):
                for c in range(cols):
                    x = margin_left + c * (card_w + gap_x)
                    y = margin_top + r * (card_h + gap_y)

                    if x + card_w > img_w or y + card_h > img_h:
                        continue

                    slot_idx = slot_count
                    slot_count += 1

                    img_path = self.id_slots.get(slot_idx)

                    if img_path and os.path.exists(img_path):
                        card_img = (
                            Image.open(img_path)
                            .convert('RGBA')
                            .resize((card_w, card_h), Image.Resampling.LANCZOS)
                        )
                        sheet_img.paste(card_img, (x, y), card_img)
                        if self.toggle_border.get():
                            draw.rectangle(
                                [x, y, x + card_w, y + card_h], outline='black', width=3
                            )

            if self.toggle_fold.get() and cols > 1:
                for c in range(cols - 1):
                    fold_x = margin_left + (c + 1) * card_w + c * gap_x + (gap_x // 2)
                    fold_y1 = margin_top
                    fold_y2 = margin_top + rows * card_h + (rows - 1) * gap_y
                    for step_y in range(fold_y1, fold_y2, 40):
                        draw.line(
                            [(fold_x, step_y), (fold_x, min(step_y + 20, fold_y2))],
                            fill='red',
                            width=3,
                        )

        else:
            pw_mm, ph_mm = self.get_current_photo_dim()
            pp_w, pp_h = int(pw_mm * scale), int(ph_mm * scale)
            rows = int(self.slider_pp_rows.get())
            cols = int(self.slider_pp_cols.get())

            if self.pp_image_path and os.path.exists(self.pp_image_path):
                photo_img = (
                    Image.open(self.pp_image_path)
                    .convert('RGBA')
                    .resize((pp_w, pp_h), Image.Resampling.LANCZOS)
                )

                for r in range(rows):
                    for c in range(cols):
                        x = margin_left + c * (pp_w + gap_x)
                        y = margin_top + r * (pp_h + gap_y)

                        if x + pp_w > img_w or y + pp_h > img_h:
                            continue

                        sheet_img.paste(photo_img, (x, y), photo_img)
                        if self.toggle_border.get():
                            draw.rectangle(
                                [x, y, x + pp_w, y + pp_h], outline='black', width=2
                            )

                if self.toggle_fold.get():
                    for c in range(1, cols):
                        fx = margin_left + c * pp_w + (c - 1) * gap_x + (gap_x // 2)
                        fy1 = margin_top
                        fy2 = margin_top + rows * pp_h + (rows - 1) * gap_y
                        for sy in range(fy1, fy2, 30):
                            draw.line(
                                [(fx, sy), (fx, min(sy + 15, fy2))], fill='red', width=2
                            )
                    for r in range(1, rows):
                        fy = margin_top + r * pp_h + (r - 1) * gap_y + (gap_y // 2)
                        fx1 = margin_left
                        fx2 = margin_left + cols * pp_w + (cols - 1) * gap_x
                        for sx in range(fx1, fx2, 30):
                            draw.line(
                                [(sx, fy), (min(sx + 15, fx2), fy)], fill='red', width=2
                            )

        return sheet_img

    def export_sheet(self, fmt):
        if self.app_mode == 'ID_CARD' and not self.id_slots:
            messagebox.showerror('Error', 'Please select at least one card image!')
            return
        elif self.app_mode == 'PASSPORT' and not self.pp_image_path:
            messagebox.showerror('Error', 'Please load a passport photo first!')
            return

        ext = f'.{fmt}'
        file_type = (
            [(f'{fmt.upper()} Image', f'*{ext}')]
            if fmt != 'pdf'
            else [('PDF Document', '*.pdf')]
        )

        save_path = filedialog.asksaveasfilename(
            initialfile='AR Studio Export',
            defaultextension=ext,
            filetypes=file_type,
            title=f'Save Sheet As {fmt.upper()}',
        )
        if not save_path:
            return

        if fmt == 'jpg':
            full_img = self.generate_full_res_image(transparent=False).convert('RGB')
            full_img.save(save_path, 'JPEG', quality=100)
        elif fmt == 'png':
            full_img = self.generate_full_res_image(transparent=True)
            full_img.save(save_path, 'PNG')
        elif fmt == 'pdf':
            self.save_pdf(save_path)

        messagebox.showinfo(
            'AR Studio Pro', f'Sheet saved successfully!\nSaved at: {save_path}'
        )

    def save_pdf(self, save_path):
        page_name = self.combo_page.get()
        page_sizes_map = {'A4': A4, 'A5': A5, 'Letter': letter, 'Legal': legal}

        base_size = page_sizes_map[page_name]
        pdf_size = (
            landscape(base_size)
            if self.combo_orient.get() == 'Landscape'
            else base_size
        )

        c = canvas.Canvas(save_path, pagesize=pdf_size)
        p_w_mm, p_h_mm = self.get_current_page_dim()

        margin_left = self.slider_margin_left.get() * mm
        margin_top = self.slider_margin_top.get() * mm
        gap_y = self.slider_gap_y.get() * mm
        gap_x = self.slider_gap_x.get() * mm

        if self.app_mode == 'ID_CARD':
            card_w, card_h = 88.9 * mm, 50.8 * mm
            rows = int(self.slider_id_rows.get())
            cols = int(self.slider_id_cols.get())

            slot_count = 1
            for r in range(rows):
                for col in range(cols):
                    x = margin_left + col * (card_w + gap_x)
                    y = (p_h_mm * mm) - margin_top - (r + 1) * card_h - r * gap_y

                    if x + card_w > (p_w_mm * mm) or y < 0:
                        continue

                    slot_idx = slot_count
                    slot_count += 1

                    img_path = self.id_slots.get(slot_idx)

                    if img_path and os.path.exists(img_path):
                        c.drawImage(img_path, x, y, width=card_w, height=card_h)
                        if self.toggle_border.get():
                            c.rect(x, y, card_w, card_h, stroke=1, fill=0)

            if self.toggle_fold.get() and cols > 1:
                c.setDash(3, 3)
                c.setStrokeColorRGB(0.8, 0, 0)
                for col in range(cols - 1):
                    fold_x = (
                        margin_left + (col + 1) * card_w + col * gap_x + (gap_x / 2.0)
                    )
                    fold_y1 = (p_h_mm * mm) - margin_top
                    fold_y2 = (
                        (p_h_mm * mm) - margin_top - rows * card_h - (rows - 1) * gap_y
                    )
                    c.line(fold_x, fold_y1, fold_x, fold_y2)

        else:
            pw_mm, ph_mm = self.get_current_photo_dim()
            pp_w, pp_h = pw_mm * mm, ph_mm * mm
            rows = int(self.slider_pp_rows.get())
            cols = int(self.slider_pp_cols.get())

            if self.pp_image_path and os.path.exists(self.pp_image_path):
                for r in range(rows):
                    for col in range(cols):
                        x = margin_left + col * (pp_w + gap_x)
                        y = (p_h_mm * mm) - margin_top - (r + 1) * pp_h - r * gap_y

                        if x + pp_w > (p_w_mm * mm) or y < 0:
                            continue

                        c.drawImage(self.pp_image_path, x, y, width=pp_w, height=pp_h)
                        if self.toggle_border.get():
                            c.rect(x, y, pp_w, pp_h, stroke=1, fill=0)

                if self.toggle_fold.get():
                    c.setDash(3, 3)
                    c.setStrokeColorRGB(0.8, 0, 0)
                    for col in range(1, cols):
                        fx = margin_left + col * pp_w + (col - 1) * gap_x + (gap_x / 2.0)
                        fy1 = (p_h_mm * mm) - margin_top
                        fy2 = (
                            (p_h_mm * mm) - margin_top - rows * pp_h - (rows - 1) * gap_y
                        )
                        c.line(fx, fy1, fx, fy2)
                    for r in range(1, rows):
                        fy = (
                            (p_h_mm * mm)
                            - margin_top
                            - r * pp_h
                            - (r - 1) * gap_y
                            - (gap_y / 2.0)
                        )
                        fx1 = margin_left
                        fx2 = margin_left + cols * pp_w + (cols - 1) * gap_x
                        c.line(fx1, fy, fx2, fy)

        c.save()

    def save_word(self):
        if self.app_mode == 'ID_CARD' and not self.id_slots:
            messagebox.showerror('Error', 'Please select at least one card image!')
            return
        elif self.app_mode == 'PASSPORT' and not self.pp_image_path:
            messagebox.showerror('Error', 'Please load a passport photo first!')
            return

        save_path = filedialog.asksaveasfilename(
            initialfile='AR Studio Export',
            defaultextension='.docx',
            filetypes=[('MS Word Document', '*.docx')],
            title='Save MS Word Document As',
        )
        if not save_path:
            return

        p_w_mm, p_h_mm = self.get_current_page_dim()

        temp_img_path = 'temp_ar_sheet.png'
        full_img = self.generate_full_res_image(transparent=True)
        full_img.save(temp_img_path, 'PNG')

        doc = Document()
        for section in doc.sections:
            section.page_width = Mm(p_w_mm)
            section.page_height = Mm(p_h_mm)
            section.top_margin = Mm(0)
            section.bottom_margin = Mm(0)
            section.left_margin = Mm(0)
            section.right_margin = Mm(0)

        doc.add_picture(temp_img_path, width=Mm(p_w_mm), height=Mm(p_h_mm))
        doc.save(save_path)

        if os.path.exists(temp_img_path):
            os.remove(temp_img_path)

        messagebox.showinfo(
            'AR Studio Pro',
            f'MS Word file created successfully!\nSaved at: {save_path}',
        )

    def open_printer_dialog(self):
        if self.app_mode == 'ID_CARD' and not self.id_slots:
            messagebox.showerror('Error', 'Please select at least one card image!')
            return
        elif self.app_mode == 'PASSPORT' and not self.pp_image_path:
            messagebox.showerror('Error', 'Please load a passport photo first!')
            return

        dialog = tk.Toplevel(self.root)
        dialog.title('Select Printer')
        dialog.geometry('380x200')
        dialog.transient(self.root)
        
        # Center the printer popup dialog on the screen
        dialog.update_idletasks()
        dialog_width = dialog.winfo_width()
        dialog_height = dialog.winfo_height()
        screen_width = dialog.winfo_screenwidth()
        screen_height = dialog.winfo_screenheight()
        pos_x = (screen_width // 2) - (dialog_width // 2)
        pos_y = (screen_height // 2) - (dialog_height // 2)
        dialog.geometry(f'{dialog_width}x{dialog_height}+{pos_x}+{pos_y}')
        
        dialog.grab_set()

        lbl = tk.Label(
            dialog,
            text='Select the printer',
            font=('Segoe UI', 9, 'bold'),
        )
        lbl.pack(pady=10)

        printers = []
        if win32print:
            printer_objs = win32print.EnumPrinters(
                win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            )
            printers = [p[2] for p in printer_objs]
        else:
            printers = ['Default Printer']

        combo_printer = ttk.Combobox(
            dialog, values=printers, state='readonly', width=32
        )
        combo_printer.pack(pady=8)

        if win32print:
            try:
                default_printer = win32print.GetDefaultPrinter()
                combo_printer.set(default_printer)
            except Exception:
                if printers:
                    combo_printer.current(0)
        elif printers:
            combo_printer.current(0)

        def do_print():
            selected_printer = combo_printer.get()
            if not selected_printer:
                messagebox.showerror('Error', 'Please select a printer!')
                return

            try:
                # Direct GDI Printing via win32ui to prevent Windows Error 31 ShellExecute issue
                full_img = self.generate_full_res_image(transparent=False).convert("RGB")
                
                hDC = win32ui.CreateDC()
                hDC.CreatePrinterDC(selected_printer)
                hDC.StartDoc("AR Studio Print Job")
                hDC.StartPage()

                printable_area = hDC.GetDeviceCaps(8), hDC.GetDeviceCaps(10)  # HORZRES, VERTRES
                dib = ImageWin.Dib(full_img)
                dib.draw(hDC.GetHandleOutput(), (0, 0, printable_area[0], printable_area[1]))

                hDC.EndPage()
                hDC.EndDoc()
                hDC.DeleteDC()

                messagebox.showinfo(
                    'Success', f'The print command has been sent to {selected_printer}'
                )
                dialog.destroy()
            except Exception as e:
                # Fallback to ShellExecute if win32ui direct print fails
                try:
                    temp_pdf = os.path.join(tempfile.gettempdir(), 'AR_Print_Sheet.pdf')
                    self.save_pdf(temp_pdf)
                    if win32api:
                        win32api.ShellExecute(0, 'printto', temp_pdf, f'"{selected_printer}"', '.', 0)
                    else:
                        os.startfile(temp_pdf, 'print')
                    messagebox.showinfo('Success', f'The print command has been sent to {selected_printer}')
                    dialog.destroy()
                except Exception as ex:
                    messagebox.showerror(
                        'Print Error', f'An error occurred while printing:\n{str(ex)}'
                    )

        btn_confirm_print = RoundedButton(
            dialog,
            text='Print Now',
            gradient_colors=('#11998e', '#38ef7d'),
            height=30,
            command=do_print,
        )
        btn_confirm_print.pack(pady=12, padx=20, fill='x')


if __name__ == '__main__':
    root = tk.Tk()
    app = ARCardStudioApp(root)
    root.mainloop()