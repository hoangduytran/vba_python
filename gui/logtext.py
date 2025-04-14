import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.font as tkFont
from enum import Enum

logger = None

# Lớp ToolTip: hiển thị tooltip khi con trỏ chuột di chuyển vào widget
class ToolTip:
    def __init__(self, widget, text='Thông tin'):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        self.id = None
        self.x = self.y = 0
        widget.bind("<Enter>", self.enter)
        widget.bind("<Leave>", self.leave)
        
    def enter(self, event=None):
        self.schedule()
        
    def leave(self, event=None):
        self.unschedule()
        self.hidetip()
        
    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(500, self.showtip)
        
    def unschedule(self):
        id_ = self.id
        self.id = None
        if id_:
            self.widget.after_cancel(id_)
            
    def showtip(self, event=None):
        if self.tipwindow or not self.text:
            return
        # Get the absolute x coordinate of the widget's left edge.
        x = self.widget.winfo_rootx()
        # Get the y coordinate by adding the widget's height to its top coordinate.
        y = self.widget.winfo_rooty() + self.widget.winfo_height()
        # Optionally, add some extra vertical offset if needed.
        y += 5  # thêm khoảng cách 5 pixel bên dưới nút

        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(tw, text=self.text, justify=tk.LEFT,
                        background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                        font=("tahoma", "15", "normal"))
        label.pack(ipadx=1)
        
    def hidetip(self):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

# Lớp Emoji: định nghĩa các emoji cho các nút trên thanh công cụ
class Emoji(Enum):
    SAVE = "💾"
    COPY = "📋"
    PASTE = "📥"
    SELECT_FONTS = "🔤"
    FONT_SIZE_UP = "🔼"
    FONT_SIZE_DOWN = "🔽"
    WRAP_TEXT = "↩️"

# Lớp LogText: tạo một khung chứa thanh công cụ và vùng hiển thị log
class LogText(tk.Frame):
    def __init__(self, master=None, mp_logging = None, **kwargs):
        """
        Khởi tạo LogText:
          - Tạo một frame chứa thanh công cụ nằm ở trên và vùng hiển thị log phía dưới.
          - Bao gồm phương thức chèn log và các chức năng của thanh công cụ.
        """
        super().__init__(master, **kwargs)
        global logger

        self.mp_logging = mp_logging
        logger = self.mp_logging.logger  # set global logger

        # Tạo khung thanh công cụ (toolbar) ở trên cùng
        self.toolbar = tk.Frame(self)
        self.toolbar.pack(side="top", fill="x")
        
        # Tạo vùng Text để hiển thị log. Chế độ wrap mặc định là "word", undo=True cho phép hoàn tác
        self.log_text = tk.Text(self, wrap="word", undo=True)
        self.log_text.pack(side="bottom", fill="both", expand=True)
        
        # Tạo thanh cuộn (scrollbar) cho vùng Text
        self.scrollbar = tk.Scrollbar(self.log_text, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=self.scrollbar.set)
        self.scrollbar.pack(side="right", fill="y")
        
        # Thiết lập font mặc định cho vùng Text: ví dụ, Arial kích thước 12
        self.font = tkFont.Font(family="Arial", size=12)
        self.log_text.configure(font=self.font)
        
        # Gọi hàm tạo các nút trên thanh công cụ
        self.create_toolbar()
        
    def create_toolbar(self):
        """
        Tạo các nút trên thanh công cụ sử dụng emoji được định nghĩa trong lớp Emoji.
        Các nút có kích thước lớn (width, height và font được tăng) và mỗi nút có tooltip với tiêu đề tiếng Việt.
        """
        # Cấu hình cho các nút: bao gồm emoji, command và tooltip (tiêu đề tiếng Việt)
        buttons_config = [
            {"emoji": Emoji.SAVE.value, "command": self.save_log, "tooltip": "Lưu log vào tập tin"},
            {"emoji": Emoji.COPY.value, "command": self.copy_text, "tooltip": "Sao chép văn bản đã chọn"},
            {"emoji": Emoji.PASTE.value, "command": self.paste_text, "tooltip": "Dán văn bản từ clipboard"},
            {"emoji": Emoji.SELECT_FONTS.value, "command": self.select_fonts, "tooltip": "Chọn phông chữ"},
            {"emoji": Emoji.FONT_SIZE_UP.value, "command": self.font_size_up, "tooltip": "Tăng kích cỡ phông chữ"},
            {"emoji": Emoji.FONT_SIZE_DOWN.value, "command": self.font_size_down, "tooltip": "Giảm kích cỡ phông chữ"},
            {"emoji": Emoji.WRAP_TEXT.value, "command": self.toggle_wrap, "tooltip": "Bật/Tắt tự động xuống dòng"}
        ]
        # Cấu hình kiểu chung cho các nút: tăng kích thước (ví dụ width=5, height=2, font kích thước 20)
        btn_style = {"font": ("Arial", 20, "bold"), "width": 5, "height": 2}
        # Tạo và đóng gói các nút vào thanh công cụ
        for config in buttons_config:
            btn = tk.Button(self.toolbar, text=config["emoji"], command=config["command"], **btn_style)
            btn.pack(side="left", padx=5, pady=5)
            # Thêm tooltip cho mỗi nút với tiêu đề tiếng Việt
            ToolTip(btn, text=config["tooltip"])
    
    def insert_log(self, text):
        """
        Chèn nội dung log vào vùng Text.
        Nội dung được thêm vào cuối cùng và cuộn xuống.
        """
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, text + "\n")
        self.log_text.configure(state="normal")
        self.log_text.see(tk.END)
    
    def save_log(self):
        """
        When the user chooses to save the log, open a save dialog with file type options for both
        a text file (*.txt) and a JSON file (*.json). If a text file is chosen, write out the content
        of the log text widget. If a JSON file is chosen, copy the temporary log file (which contains
        JSON-formatted logs) to the chosen filename.
        """
        import shutil
        # Open the save dialog with two format choices.
        path = filedialog.asksaveasfilename(
            title="Lưu Log vào tập tin",
            defaultextension=".txt",
            filetypes=[("Tệp văn bản (*.txt)", "*.txt"), ("Tệp JSON (*.json)", "*.json")]
        )
        if path:
            try:
                if path.lower().endswith(".json"):
                    # For JSON, simply copy the temporary log file,
                    # which is written in JSON-line format.
                    shutil.copyfile(self.mp_logging.log_temp_file_path, path)
                else:
                    # For .txt, write out the human-readable log text (from the LogText widget).
                    with open(path, "w", encoding="utf-8") as f:
                        f.write(self.log_text.get("1.0", tk.END))
                messagebox.showinfo("Thông báo", "Log đã được lưu thành công.")
                logger.info("Log đã được lưu thành công")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Lỗi khi lưu log: {e}")
    
    def copy_text(self):
        """
        Sao chép văn bản được chọn từ vùng Text vào clipboard.
        """
        try:
            selected_text = self.log_text.get("sel.first", "sel.last")
            self.clipboard_clear()
            self.clipboard_append(selected_text)
        except tk.TclError:
            pass
    
    def paste_text(self):
        """
        Dán văn bản từ clipboard vào vị trí con trỏ trong vùng Text.
        """
        try:
            clipboard_text = self.clipboard_get()
            self.log_text.insert(tk.INSERT, clipboard_text)
        except tk.TclError:
            pass
    
    def select_fonts(self):
        """
        Cho phép người dùng chọn phông chữ từ một danh sách đã được định sẵn.
        Một cửa sổ con sẽ xuất hiện với các tùy chọn phông chữ.
        """
        top = tk.Toplevel(self)
        top.title("Chọn phông chữ")
        tk.Label(top, text="Phông chữ:").pack(side="left", padx=5, pady=5)
        
        font_options = ["Arial", "Courier New", "Times New Roman", "Verdana", "Tahoma"]
        var = tk.StringVar(value=self.font.actual("family"))
        option_menu = tk.OptionMenu(top, var, *font_options)
        option_menu.pack(side="left", padx=5, pady=5)
        
        def update_font():
            self.font.configure(family=var.get())
            top.destroy()
        tk.Button(top, text="OK", command=update_font).pack(side="left", padx=5, pady=5)
    
    def font_size_up(self):
        """
        Tăng kích thước phông chữ của vùng Text.
        """
        current_size = self.font.actual("size")
        self.font.configure(size=current_size + 2)
    
    def font_size_down(self):
        """
        Giảm kích thước phông chữ của vùng Text (không giảm dưới 1).
        """
        current_size = self.font.actual("size")
        new_size = current_size - 2 if current_size > 2 else 1
        self.font.configure(size=new_size)
    
    def toggle_wrap(self):
        """
        Bật/Tắt chế độ tự động xuống dòng cho vùng Text.
        Nếu hiện tại wrap là "word", chuyển sang "none"; ngược lại chuyển về "word".
        """
        current_wrap = self.log_text.cget("wrap")
        new_wrap = "none" if current_wrap == "word" else "word"
        self.log_text.configure(wrap=new_wrap)
