import tkinter as tk
import tkinter.font as tkFont
from tkinter import ttk, filedialog, messagebox
import tkinter.font as tkFont
from enum import Enum
from gv import Gvar as gv, COMMON_WIDGET_STYLE, FONT_BASIC, font_options
from mpp_logger import LOG_LEVELS
from gui_actions import action_list

logger = None
option_style = None

# Lớp ToolTip: hiển thị tooltip khi con trỏ chuột di chuyển vào widget
import tkinter as tk
from tkinter import font as tkFont

class ToolTip:
    def __init__(self, widget, text='Thông tin', max_width=400):
        """
        widget    : widget mà tooltip gắn vào
        text      : nội dung văn bản hiển thị
        max_width : giới hạn chiều rộng tính theo pixel (nếu text quá dài, sẽ wrap)
        """
        self.widget = widget
        self.text = text
        self.tipwindow = None
        self.id = None
        self.x = self.y = 0
        self.max_width = max_width

        # Gán sự kiện di chuyển vào/ra widget
        widget.bind("<Enter>", self.enter)
        widget.bind("<Leave>", self.leave)

        # Khởi tạo font cho tooltip
        self.label_font = tkFont.Font(family="Arial", size=15)
        
        # Đo chiều rộng văn bản tính theo pixel
        text_width_px = self.label_font.measure(text)
        
        # Nếu văn bản + chút lề > max_width, ta sẽ dùng wrap
        if text_width_px > self.max_width:
            self.wraplength = self.max_width
        else:
            self.wraplength = 0  # =0 nghĩa là không wrap

    def enter(self, event=None):
        # Lên lịch hiển thị tooltip sau 500ms
        self.schedule()
        
    def leave(self, event=None):
        self.unschedule()
        self.hidetip()
        
    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(1000, self.showtip)
        
    def unschedule(self):
        if self.id:
            self.widget.after_cancel(self.id)
        self.id = None

    def showtip(self, event=None):
        """
        Tạo một cửa sổ toplevel, hiển thị text (có wrap nếu cần).
        Đặt vị trí ngay dưới widget.
        """
        # Nếu đã có tipwindow hoặc text rỗng, bỏ qua
        if self.tipwindow or not self.text:
            return

        # Tính toạ độ cho tooltip: ngay cạnh dưới widget
        x = self.widget.winfo_rootx()
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5

        # Tạo cửa sổ Toplevel cho tooltip
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)  # Xoá viền, thanh title
        tw.wm_geometry("+%d+%d" % (x, y))

        # Tạo nhãn để hiển thị text trong tipwindow
        label = tk.Label(
            tw,
            text=self.text,
            justify=tk.LEFT,
            background="#b7c8be",
            relief=tk.SOLID,
            borderwidth=0,
            font=self.label_font
        )

        # Nếu cần wrap (wraplength > 0), cấu hình wraplength
        if self.wraplength > 0:
            label.config(wraplength=self.wraplength)

        label.pack(ipadx=5, ipady=5)  # Thêm chút padding trong Label

    def hidetip(self):
        if self.tipwindow:
            self.tipwindow.destroy()
        self.tipwindow = None


# Lớp Emoji: định nghĩa các emoji cho các nút trên thanh công cụ
class Emoji(Enum):
    SAVE = "💾"
    COPY = "📋"
    PASTE = "📥"
    SELECT_FONTS = "🔤"
    FONT_SIZE_UP = "🔼"
    FONT_SIZE_DOWN = "🔽"
    WRAP_TEXT = "↩️"
    CLEAR = "🗑"  # Emoji cho nút xóa log

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
        # Thiết lập style TTK cho nút (Button) và checkbutton
        global option_style

        self.style = ttk.Style(self)
        # Style cho nút lớn, font Arial 18 đậm
        self.style.configure(
            "App.TMenubutton",  # Tên bạn chọn, thường kèm hậu tố .TMenubutton
            font=("Arial", 18, "normal"),
            padding=8
        )        
        

        # -----------------------------------------
        # ADD the log_level_menu at the end of toolbar
        # Tạo Biến chuỗi cho log_level_var, tham chiếu gv.log_level_var
        # (assuming we still keep it in gv)
        gv.log_level_var = tk.StringVar(value="DEBUG")
        # Drop-down
        self.log_level_menu = ttk.OptionMenu(
            self.toolbar,
            gv.log_level_var,
            gv.log_level_var.get(),  # initial
            *LOG_LEVELS.keys(),
            command=action_list["select_log_level"]
        )
        # Style or config if you want
        self.log_level_menu.config(style="App.TMenubutton", width=15)
        self.log_level_menu["menu"].config(font=FONT_BASIC)
        self.log_level_menu.pack(side="left", padx=5)
        
        ToolTip(self.log_level_menu, text="Chọn mức log (DEBUG -> CRITICAL)")

        # exact_check
        gv.is_exact_var = tk.BooleanVar(value=True)
        self.exact_check = tk.Checkbutton(
            self.toolbar,
            text="Duy Cấp Độ",
            variable=gv.is_exact_var,
            font=FONT_BASIC,
            command=action_list["update_gui_filter"],
        )
        self.exact_check.pack(side="left", padx=5)
        ToolTip(self.exact_check, text="Chỉ hiển thị đúng cấp độ này hay từ nó trở lên")

        # Cấu hình cho các nút: bao gồm emoji, command và tooltip (tiêu đề tiếng Việt)
        buttons_config = [
            {"emoji": Emoji.SAVE.value, "command": self.save_log, "tooltip": "Lưu toàn bộ log vào tập tin"},
            {"emoji": Emoji.CLEAR.value, "command": self.clear_log, "tooltip": "Xóa log"},  # Nút xóa log mới
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
    
    def clear_log(self):
        """
        Xóa toàn bộ nội dung của vùng Text log.
        """
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state="disabled")

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
            expl = ""
            try:
                if path.lower().endswith(".json"):
                    # For JSON, simply copy the temporary log file,
                    # which is written in JSON-line format.
                    shutil.copyfile(self.mp_logging.log_temp_file_path, path)
                    expl = "toàn bộ nội dung trong định dạng JSON"
                else:
                    # For .txt, write out the human-readable log text (from the LogText widget).
                    with open(path, "w", encoding="utf-8") as f:
                        f.write(self.log_text.get("1.0", tk.END))
                    expl = "duy nội dung trong hộp văn bản ở định dạng văn bản thường"
                messagebox.showinfo("Thông báo", f"Log đã được lưu thành công với {expl}")
                logger.info(f"Log đã được lưu thành công với {expl}")
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
        
        
        var = tk.StringVar(value=self.font.actual("family"))
        option_menu = tk.OptionMenu(top, var, *font_options)
        option_menu["menu"].config(font=FONT_BASIC)
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
