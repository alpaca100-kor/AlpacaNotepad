import tkinter as tk
from tkinter import messagebox, filedialog, Toplevel, font, ttk
import json
import os
import sys
import configparser
from datetime import datetime

# 엑셀 파일 처리를 위한 라이브러리. (없으면 엑셀 내보내기 비활성화)
try:
    import openpyxl
except ImportError:
    openpyxl = None

# Windows 11 스타일(Sun Valley) ttk 테마. (없으면 기본 ttk 테마로 동작)
try:
    import sv_ttk
except ImportError:
    sv_ttk = None


def resource_path(relative_path):
    """PyInstaller로 생성된 exe의 리소스 경로를 가져옴"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def get_app_dir():
    """실행 파일(exe) 또는 스크립트가 실제로 위치한 폴더 경로를 반환.

    resource_path()의 sys._MEIPASS는 PyInstaller onefile 실행 시 임시로 압축 해제되는
    폴더라 프로그램 종료 후 사라지므로, memos.json/settings.ini처럼 계속 남아있어야 하는
    사용자 데이터 파일 경로에는 사용하면 안 됨. 이 함수는 exe(또는 .py) 자체가 있는
    폴더를 반환해서, 실행 시 현재 작업 디렉터리(cwd)가 무엇이든 관계없이 항상 같은
    위치에 데이터를 저장/로드하도록 함.
    """
    if getattr(sys, "frozen", False):
        # PyInstaller로 빌드된 exe: exe 파일이 있는 폴더
        return os.path.dirname(sys.executable)
    # 일반 .py 스크립트로 실행되는 경우: 스크립트 파일이 있는 폴더
    return os.path.dirname(os.path.abspath(__file__))


# 다크/라이트 테마별 색상 팔레트 (sv_ttk 패키지의 실제 팔레트 값 기준)
# ttk가 직접 테마를 입히지 못하는 Listbox/Text 위젯에 수동으로 적용하기 위함
# 빠른 입력(Alt+1~Alt+0) 슬롯 키 순서: 1,2,...,9,0
QUICK_INPUT_KEYS = [str(i) for i in range(1, 10)] + ["0"]

# Windows 가상 키코드(VK_0~VK_9) -> 숫자 문자열 매핑
# (<Alt-1> 같은 개별 keysym 바인딩이 Windows에서 씹히는 문제를 우회하기 위해
#  <Alt-KeyPress>로 받은 뒤 keycode로 눌린 키를 직접 판별하는 데 사용)
ALT_DIGIT_VK_CODES = {48: "0", 49: "1", 50: "2", 51: "3", 52: "4",
                       53: "5", 54: "6", 55: "7", 56: "8", 57: "9"}

THEME_COLORS = {
    "light": {
        "bg": "#fafafa",
        "fg": "#1c1c1c",
        "border": "#e0e0e0",
        "accent": "#005fb8",
        "list_select_bg": "#0067c0",
        "list_select_fg": "#ffffff",
        "text_select_bg": "#2f60d8",
        "text_select_fg": "#ffffff",
        "insert_bg": "#1c1c1c",
        "disabled_bg": "#f3f3f3",
        "success_fg": "#0067c0",
    },
    "dark": {
        "bg": "#1c1c1c",
        "fg": "#fafafa",
        "border": "#3a3a3a",
        "accent": "#57c8ff",
        "list_select_bg": "#0067c0",
        "list_select_fg": "#ffffff",
        "text_select_bg": "#2f60d8",
        "text_select_fg": "#ffffff",
        "insert_bg": "#fafafa",
        "disabled_bg": "#252525",
        "success_fg": "#ffd600",
    },
}


def apply_titlebar_theme(root, dark):
    """Windows 10/11에서 창 제목표시줄 색상까지 다크모드에 맞춤.
    Windows가 아니거나 API를 지원하지 않으면 조용히 무시함(다른 OS 크래시 방지)."""
    try:
        import ctypes
        root.update_idletasks()

        hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
        if not hwnd:
            # GetParent로 못 찾으면 창 제목으로 실제 최상위 창을 직접 검색 (폴백)
            hwnd = ctypes.windll.user32.FindWindowW(None, root.title())
        if not hwnd:
            return

        DWMWA_USE_IMMERSIVE_DARK_MODE = 20
        value = ctypes.c_int(1 if dark else 0)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(
            hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ctypes.byref(value), ctypes.sizeof(value)
        )

        # DwmSetWindowAttribute만으로는 이미 화면에 떠 있는 창의 타이틀바가
        # 곧바로 다시 그려지지 않는 경우가 있어, 프레임을 강제로 한 번 갱신시켜준다.
        SWP_NOMOVE, SWP_NOSIZE, SWP_NOZORDER, SWP_FRAMECHANGED = 0x0002, 0x0001, 0x0004, 0x0020
        ctypes.windll.user32.SetWindowPos(
            hwnd, 0, 0, 0, 0, 0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED
        )
    except Exception:
        pass


class MemoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("알파카 메모장 (Alpaca Notepad)")
        self.root.minsize(800, 600)
        # 아이콘 설정 (오류 발생 시 무시)
        try:
            self.root.iconbitmap(resource_path("an.ico"))
        except Exception as e:
            print(f"아이콘 로드 실패: {e}")
            pass

        self.file_path = os.path.join(get_app_dir(), "memos.json")
        self.settings_file = os.path.join(get_app_dir(), "settings.ini")
        self.quick_input_file = os.path.join(get_app_dir(), "quick_inputs.json")
        self.memos = self.load_memos()
        self.settings = self.load_settings()
        self.quick_inputs = self.load_quick_inputs()
        self.current_index = -1

        # Windows 11 스타일(sv_ttk) 테마 적용 (sv_ttk 미설치 시 기본 ttk 테마 사용)
        self.theme_mode = self.settings.get("theme_mode", "light")
        if sv_ttk:
            sv_ttk.set_theme(self.theme_mode, self.root)
        apply_titlebar_theme(self.root, self.theme_mode == "dark")
        # sv_ttk가 테마 변경 시 내부적으로 <<ThemeChanged>> 이벤트(큐에 쌓임)로
        # 위젯 색상/팔레트를 갱신하는데, 이 처리가 끝나기 전에 아래에서 위젯을 만들면
        # 메뉴/버튼/라벨 일부가 예전 회색(#d9d9d9 계열)으로 만들어지는 문제가 있어
        # 위젯 생성 전에 강제로 큐를 한 번 비워준다
        self.root.update_idletasks()

        self.ui_font = ("맑은 고딕", 12)
        self.content_font = (self.settings.get("font_family"), self.settings.get("font_size"))
        self.copy_shortcut = self.settings.get("copy_shortcut")

        self.create_menu()

        # 상태표시줄 생성
        self.status_bar = ttk.Label(root, text="", relief=tk.SUNKEN, anchor=tk.W, font=("맑은 고딕", 10))
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.main_pane = ttk.PanedWindow(root, orient=tk.HORIZONTAL)
        self.main_pane.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        main_pane = self.main_pane

        # 왼쪽 패널: 메모 리스트
        left_panel = ttk.Frame(main_pane)
        main_pane.add(left_panel, weight=0)

        list_frame = ttk.Frame(left_panel)
        list_frame.pack(fill=tk.BOTH, expand=True)

        self.listbox = tk.Listbox(list_frame, exportselection=False, font=self.ui_font)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.listbox.bind("<<ListboxSelect>>", self.on_memo_select)
        self.listbox.bind("<Delete>", lambda event: self.remove_memo())
        self.listbox.bind("<Home>", self.on_home_key)
        self.listbox.bind("<End>", self.on_end_key)

        # 드래그 앤 드롭 관련 변수 초기화
        self.drag_start_index = None
        
        # 드래그 앤 드롭 이벤트 바인딩
        self.listbox.bind("<ButtonPress-1>", self.on_drag_start)
        self.listbox.bind("<B1-Motion>", self.on_drag_motion)
        self.listbox.bind("<ButtonRelease-1>", self.on_drag_drop)

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.config(yscrollcommand=scrollbar.set)

        self.update_listbox()

        # 버튼 프레임
        button_frame = ttk.Frame(left_panel)
        button_frame.pack(fill=tk.X, pady=5)

        add_button = ttk.Button(button_frame, text="추가", command=self.add_memo)
        add_button.pack(side=tk.LEFT, expand=True, fill=tk.X)
        remove_button = ttk.Button(button_frame, text="제거", command=self.remove_memo)
        remove_button.pack(side=tk.LEFT, expand=True, fill=tk.X)
        up_button = ttk.Button(button_frame, text="▲", command=self.move_memo_up)
        up_button.pack(side=tk.LEFT, expand=True, fill=tk.X)
        down_button = ttk.Button(button_frame, text="▼", command=self.move_memo_down)
        down_button.pack(side=tk.LEFT, expand=True, fill=tk.X)

        # 오른쪽 패널: 제목 + 내용
        right_panel = ttk.Frame(main_pane)
        main_pane.add(right_panel, weight=1)

        title_label = ttk.Label(right_panel, text="메모 제목", font=self.ui_font)
        title_label.pack(anchor="w")
        self.title_entry = ttk.Entry(right_panel, font=self.ui_font)
        self.title_entry.pack(fill=tk.X, pady=(0, 10))
        self.title_entry.bind("<KeyRelease>", self.update_memo_realtime)
        self.title_entry.bind("<Control-t>", self.focus_on_title)

        content_label = ttk.Label(right_panel, text="메모 내용", font=self.ui_font)
        content_label.pack(anchor="w")
        self.content_text = tk.Text(right_panel, font=self.content_font, padx=10, pady=8)
        self.content_text.pack(fill=tk.BOTH, expand=True)
        self.content_text.bind("<KeyRelease>", lambda event: self.update_memo_realtime(event, update_list=False))
        self.content_text.bind("<ButtonRelease-1>", self.update_status_bar)
        self.content_text.bind("<Control-t>", self.focus_on_title)

        # 복사 버튼과 복사 완료 메시지를 위한 프레임
        copy_frame = ttk.Frame(right_panel)
        copy_frame.pack(anchor="e", pady=5)
        
        self.copy_status_label = ttk.Label(copy_frame, text="", font=("맑은 고딕", 11, "bold"))
        self.copy_status_label.pack(side=tk.LEFT, padx=(0, 10))
        
        self.copy_button = ttk.Button(copy_frame, text="클립보드로 복사", command=self.copy_to_clipboard)
        self.copy_button.pack(side=tk.LEFT)

        # 단축키 등록
        self.root.bind("<Control-n>", lambda event: self.add_memo())
        self.root.bind("<Prior>", lambda event: self.move_memo_up())
        self.root.bind("<Next>", lambda event: self.move_memo_down())
        self.root.bind("<Control-l>", self.focus_on_listbox)
        self.root.bind("<Control-t>", self.focus_on_title)
        self.root.bind("<Alt-t>", self.insert_datetime)
        # 빠른 입력 단축키 (Alt+1 ~ Alt+9, Alt+0)
        # Windows에서는 <Alt-1>처럼 특정 숫자 keysym을 그대로 bind()하면
        # WM_SYSKEYDOWN 처리 특성상 이벤트가 씹혀 동작하지 않는 경우가 있어,
        # <Alt-KeyPress>로 Alt+모든 키 입력을 받은 뒤 내부에서 눌린 키를 판별한다.
        self.root.bind("<Alt-KeyPress>", self._on_alt_number_keypress)
        # 복사 단축키는 설정에 따라 바인딩
        self.bind_copy_shortcut()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.toggle_right_panel(False)
        self.apply_classic_widget_colors()
        self.update_status_bar()
        
        # UI 생성 완료 후 창 위치/크기 복원
        self.restore_window_geometry()

        # ⚠️ 시작 시 테마 미적용 문제 보완:
        # 위젯이 아직 없거나 창이 화면에 표시되기 전에 테마를 적용하면
        # 메뉴/버튼/라벨/타이틀바 일부가 테마를 제대로 못 받는 경우가 있음
        # (설정창에서 같은 테마를 다시 '저장'하면 정상으로 보이는 것과 동일한 현상).
        # 창이 완전히 그려진 뒤 테마를 한 번 더 적용해 이를 자동으로 바로잡는다.
        self.root.after(150, self._reapply_theme_on_startup)

    def _reapply_theme_on_startup(self):
        """창이 완전히 표시된 뒤 테마를 다시 한번 적용해, 시작 시 일부 위젯이
        테마를 제대로 반영하지 못하는 렌더링 문제를 보정한다."""
        if sv_ttk:
            sv_ttk.set_theme(self.theme_mode, self.root)
        apply_titlebar_theme(self.root, self.theme_mode == "dark")
        self.apply_classic_widget_colors()

    def restore_window_geometry(self):
        """창 위치와 크기를 복원"""
        geometry = self.settings.get("window_geometry", "800x600+100+100")
        
        try:
            # geometry 문자열 검증: "800x600+100+100" 형식
            if geometry and 'x' in geometry and '+' in geometry:
                # UI 레이아웃 완료 대기
                self.root.update_idletasks()
                self.root.geometry(geometry)
                print(f"✅ 창 위치 복원: {geometry}")
            else:
                # 잘못된 형식이면 기본값 사용
                print(f"⚠️ 잘못된 geometry 형식: {geometry}")
                self.root.geometry("800x600+100+100")
        except Exception as e:
            print(f"❌ 창 위치 복원 실패: {e}")
            self.root.geometry("800x600+100+100")

        # 좌측 리스트 패널 초기 폭 지정
        # (ttk.PanedWindow는 add()에서 width/minsize를 못 받아 sashpos로 대신 처리)
        try:
            self.root.update_idletasks()
            self.main_pane.sashpos(0, 250)
        except Exception:
            pass
        self.main_pane.bind("<ButtonRelease-1>", self.enforce_min_sash)

    def enforce_min_sash(self, event=None):
        """왼쪽 리스트 패널이 너무 좁아지지 않도록 최소 폭(200px)을 보정"""
        try:
            if self.main_pane.sashpos(0) < 200:
                self.main_pane.sashpos(0, 200)
        except Exception:
            pass

    def create_menu(self):
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="메모 가져오기...", command=self.import_memos)
        file_menu.add_command(label="메모 내보내기...", command=self.export_memos)
        file_menu.add_separator()
        file_menu.add_command(label="종료", command=self.on_closing)
        menubar.add_cascade(label="파일", menu=file_menu)

        settings_menu = tk.Menu(menubar, tearoff=0)
        settings_menu.add_command(label="글꼴 설정...", command=self.open_font_settings)
        settings_menu.add_command(label="단축키 설정...", command=self.open_shortcut_settings)
        settings_menu.add_command(label="빠른 입력 설정...", command=self.open_quick_input_settings)
        settings_menu.add_command(label="테마 설정...", command=self.open_theme_settings)
        menubar.add_cascade(label="설정", menu=settings_menu)

        self.root.config(menu=menubar)

    def focus_on_listbox(self, event=None):
        self.listbox.focus_set()
        if self.current_index != -1:
            self.listbox.selection_set(self.current_index)
            self.listbox.activate(self.current_index)
        return "break"

    def focus_on_title(self, event=None):
        # ttk.Entry의 cget('state')는 일반 str이 아닌 Tcl 객체를 반환하므로 str()로 변환 후 비교해야 함
        if str(self.title_entry.cget('state')) == tk.NORMAL:
            self.title_entry.focus_set()
            self.title_entry.select_range(0, tk.END)
        return "break"

    def _insert_text_at_focus(self, text):
        """포커스된 위젯(제목/내용)의 커서 위치에 문자열을 삽입하는 공통 로직.
        (Alt+T 날짜/시간 삽입, Alt+숫자 빠른 입력에서 공용으로 사용)"""
        focused = self.root.focus_get()

        if focused is self.title_entry:
            try:
                self.title_entry.delete("sel.first", "sel.last")
            except tk.TclError:
                pass
            self.title_entry.insert(tk.INSERT, text)
            self.update_memo_realtime(update_list=True)
        else:
            # content_text에 포커스가 있는 경우 커서 위치에, 그 외에는 내용 끝에 삽입
            try:
                self.content_text.delete("sel.first", "sel.last")
            except tk.TclError:
                pass
            if focused is self.content_text:
                self.content_text.insert(tk.INSERT, text)
            else:
                self.content_text.insert(tk.END, text)
            self.update_memo_realtime(update_list=False)

    def insert_datetime(self, event=None):
        """Alt+T: 포커스된 위젯의 커서 위치에 현재 날짜/시간을 삽입 (yyyy-mm-dd hh:mm:ss)"""
        if self.current_index == -1:
            return "break"
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._insert_text_at_focus(now_str)
        return "break"

    def insert_quick_text(self, key, event=None):
        """Alt+1~Alt+0: 설정 메뉴 > 빠른 입력 설정에서 미리 지정한 문구를 삽입"""
        if self.current_index == -1:
            return "break"
        text = self.quick_inputs.get(key, "")
        if not text:
            # 해당 슬롯에 등록된 문구가 없으면 아무 동작도 하지 않음
            return "break"
        self._insert_text_at_focus(text)
        return "break"

    def _on_alt_number_keypress(self, event):
        """<Alt-KeyPress> 범용 핸들러: Alt+숫자(1~9,0) 입력만 골라내 빠른 입력을 실행.

        Windows에서 <Alt-1>처럼 특정 keysym을 직접 bind()하면 WM_SYSKEYDOWN 처리
        특성상 이벤트가 씹혀 아예 안 잡히는 경우가 있어, Alt+모든 키를 통째로 받은 뒤
        1) Windows 가상 키코드(keycode)로 먼저 판별하고,
        2) 안되면 keysym으로도 판별하는(리눅스/맥 등 대비) 이중 방식을 사용한다.
        해당하지 않는 Alt 조합(Alt+T 등)은 그대로 통과시켜 다른 바인딩이 처리하게 둔다.
        """
        key = ALT_DIGIT_VK_CODES.get(event.keycode)
        if key is None and event.keysym in QUICK_INPUT_KEYS:
            key = event.keysym
        if key is not None:
            return self.insert_quick_text(key, event)
        return None

    def on_home_key(self, event=None):
        if not self.memos:
            return "break"
        self.listbox.focus_set()
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(0)
        self.on_memo_select(None)
        self.listbox.activate(0)
        self.listbox.see(0)
        return "break"

    def on_end_key(self, event=None):
        if not self.memos:
            return "break"
        last_index = len(self.memos) - 1
        self.listbox.focus_set()
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(last_index)
        self.on_memo_select(None)
        self.listbox.activate(last_index)
        self.listbox.see(last_index)
        return "break"

    def copy_to_clipboard(self, event=None):
        if self.current_index == -1:
            return "break"
        try:
            text_to_copy = self.content_text.get("1.0", tk.END).strip()
            if not text_to_copy:
                messagebox.showinfo("알림", "복사할 내용이 없습니다.", parent=self.root)
                return "break"
            self.root.clipboard_clear()
            self.root.clipboard_append(text_to_copy)
            
            # "복사 완료!" 메시지 표시
            self.show_copy_success_message()
            
        except tk.TclError:
            messagebox.showerror("오류", "클립보드에 접근할 수 없습니다.", parent=self.root)
        except Exception as e:
            messagebox.showerror("오류", f"알 수 없는 오류 발생: {e}", parent=self.root)
        return "break"
    
    def show_copy_success_message(self):
        """복사 완료 메시지를 1초 동안 표시 (라이트=파란색, 다크=노란색)"""
        colors = THEME_COLORS.get(self.theme_mode, THEME_COLORS["light"])
        self.copy_status_label.config(text="복사 완료!", foreground=colors["success_fg"])
        # 1초(1000ms) 후에 메시지 제거
        self.root.after(1000, lambda: self.copy_status_label.config(text=""))

    def update_status_bar(self, event=None):
        try:
            memo_count = len(self.memos)
            status_text = f"총 {memo_count}개의 메모"
            if self.current_index != -1 and self.content_text.cget('state') == tk.NORMAL:
                cursor_pos = self.content_text.index(tk.INSERT)
                row, column = cursor_pos.split('.')
                row = int(row)
                column = int(column) + 1
                content = self.content_text.get("1.0", tk.END).strip()
                char_count = len(content)
                status_text += f" / 커서 위치: 줄 {row}, 열 {column} / 글자 수: {char_count}자"
            self.status_bar.config(text=status_text)
        except Exception:
            pass

    def load_settings(self):
        config = configparser.ConfigParser()
        default_settings = {
            'font_family': '맑은 고딕',
            'font_size': 12,
            'window_geometry': '800x600+100+100',
            'copy_shortcut': 'Ctrl+Shift+C',
            'theme_mode': 'light'
        }

        if not os.path.exists(self.settings_file):
            return default_settings

        try:
            config.read(self.settings_file, encoding='utf-8')
            font_family = config.get('Font', 'family', fallback=default_settings['font_family'])
            font_size = config.getint('Font', 'size', fallback=default_settings['font_size'])
            window_geometry = config.get('Window', 'geometry', fallback=default_settings['window_geometry'])
            copy_shortcut = config.get('Shortcuts', 'copy', fallback=default_settings['copy_shortcut'])
            theme_mode = config.get('Theme', 'mode', fallback=default_settings['theme_mode'])
            if theme_mode not in ("light", "dark"):
                theme_mode = default_settings['theme_mode']
            return {
                'font_family': font_family,
                'font_size': font_size,
                'window_geometry': window_geometry,
                'copy_shortcut': copy_shortcut,
                'theme_mode': theme_mode
            }
        except (configparser.Error, ValueError):
            return default_settings

    def save_settings(self):
        config = configparser.ConfigParser()
        config['Font'] = {
            'family': self.settings.get('font_family', '맑은 고딕'),
            'size': str(self.settings.get('font_size', 12))
        }
        config['Window'] = {
            'geometry': self.root.geometry()
        }
        config['Shortcuts'] = {
            'copy': self.settings.get('copy_shortcut', 'Ctrl+Shift+C')
        }
        config['Theme'] = {
            'mode': self.theme_mode
        }
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as configfile:
                config.write(configfile)
            print(f"✅ 설정 저장: {self.root.geometry()}")
        except Exception as e:
            print(f"❌ 설정 저장 실패: {e}")

    def open_font_settings(self):
        settings_win = Toplevel(self.root)
        settings_win.title("글꼴 설정")
        settings_win.geometry("350x150")
        settings_win.resizable(False, False)
        settings_win.transient(self.root)
        settings_win.grab_set()
        settings_win.focus_force()

        ttk.Label(settings_win, text="글꼴:", font=self.ui_font).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        font_families = sorted(font.families())
        font_var = tk.StringVar(value=self.content_font[0])
        font_combo = ttk.Combobox(settings_win, textvariable=font_var, values=font_families, state="readonly")
        font_combo.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        font_combo.focus_set()

        ttk.Label(settings_win, text="크기:", font=self.ui_font).grid(row=1, column=0, padx=10, pady=10, sticky="w")
        size_var = tk.StringVar(value=str(self.content_font[1]))
        size_spinbox = ttk.Spinbox(settings_win, from_=8, to=72, textvariable=size_var, width=5)
        size_spinbox.grid(row=1, column=1, padx=10, pady=10, sticky="w")

        def apply_and_save():
            new_font_family = font_var.get()
            try:
                new_font_size = int(size_var.get())
                self.content_font = (new_font_family, new_font_size)
                self.content_text.config(font=self.content_font)
                self.settings["font_family"] = new_font_family
                self.settings["font_size"] = new_font_size
                self.save_settings()
            except ValueError:
                messagebox.showerror("오류", "올바른 글자 크기를 입력하세요.", parent=settings_win)

        def save_action():
            apply_and_save()
            settings_win.destroy()

        def cancel_action():
            settings_win.destroy()

        button_frame = ttk.Frame(settings_win)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)
        save_button = ttk.Button(button_frame, text="저장", command=save_action, width=10)
        save_button.pack(side=tk.LEFT, padx=5)
        cancel_button = ttk.Button(button_frame, text="취소", command=cancel_action, width=10)
        cancel_button.pack(side=tk.LEFT, padx=5)

    def open_shortcut_settings(self):
        settings_win = Toplevel(self.root)
        settings_win.title("단축키 설정")
        settings_win.geometry("350x150")
        settings_win.resizable(False, False)
        settings_win.transient(self.root)
        settings_win.grab_set()
        settings_win.focus_force()

        ttk.Label(settings_win, text="클립보드 복사:", font=self.ui_font).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        shortcut_options = ["Ctrl+Shift+C", "Ctrl+Alt+C", "Alt+Shift+C"]
        shortcut_var = tk.StringVar(value=self.copy_shortcut)
        shortcut_combo = ttk.Combobox(settings_win, textvariable=shortcut_var, values=shortcut_options, state="readonly", width=20)
        shortcut_combo.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        shortcut_combo.focus_set()

        def apply_and_save():
            new_shortcut = shortcut_var.get()
            # 기존 단축키 바인딩 제거
            self.unbind_copy_shortcut()
            # 새 단축키 설정
            self.copy_shortcut = new_shortcut
            self.settings["copy_shortcut"] = new_shortcut
            # 새 단축키 바인딩
            self.bind_copy_shortcut()
            self.save_settings()

        def save_action():
            apply_and_save()
            settings_win.destroy()

        def cancel_action():
            settings_win.destroy()

        button_frame = ttk.Frame(settings_win)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)
        save_button = ttk.Button(button_frame, text="저장", command=save_action, width=10)
        save_button.pack(side=tk.LEFT, padx=5)
        cancel_button = ttk.Button(button_frame, text="취소", command=cancel_action, width=10)
        cancel_button.pack(side=tk.LEFT, padx=5)

    def open_theme_settings(self):
        settings_win = Toplevel(self.root)
        settings_win.title("테마 설정")
        settings_win.geometry("300x140")
        settings_win.resizable(False, False)
        settings_win.transient(self.root)
        settings_win.grab_set()
        settings_win.focus_force()

        ttk.Label(settings_win, text="테마:", font=self.ui_font).grid(row=0, column=0, padx=10, pady=10, sticky="w")

        theme_labels = {"light": "밝게 (Light)", "dark": "어둡게 (Dark)"}
        label_to_mode = {v: k for k, v in theme_labels.items()}
        theme_var = tk.StringVar(value=theme_labels.get(self.theme_mode, theme_labels["light"]))
        theme_combo = ttk.Combobox(
            settings_win, textvariable=theme_var,
            values=list(theme_labels.values()), state="readonly", width=15
        )
        theme_combo.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        theme_combo.focus_set()

        def apply_and_save():
            new_mode = label_to_mode.get(theme_var.get(), "light")
            self.theme_mode = new_mode
            self.settings["theme_mode"] = new_mode
            if sv_ttk:
                sv_ttk.set_theme(new_mode, self.root)
            apply_titlebar_theme(self.root, new_mode == "dark")
            self.apply_classic_widget_colors()
            self.save_settings()

        def save_action():
            apply_and_save()
            settings_win.destroy()

        def cancel_action():
            settings_win.destroy()

        button_frame = ttk.Frame(settings_win)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)
        save_button = ttk.Button(button_frame, text="저장", command=save_action, width=10)
        save_button.pack(side=tk.LEFT, padx=5)
        cancel_button = ttk.Button(button_frame, text="취소", command=cancel_action, width=10)
        cancel_button.pack(side=tk.LEFT, padx=5)

    def open_quick_input_settings(self):
        """Alt+1~Alt+0 빠른 입력 문구를 설정하는 팝업창"""
        settings_win = Toplevel(self.root)
        settings_win.title("빠른 입력 설정")
        settings_win.resizable(False, False)
        settings_win.transient(self.root)
        settings_win.grab_set()
        settings_win.focus_force()

        info_label = ttk.Label(
            settings_win,
            text="Alt+숫자 키를 눌렀을 때 삽입할 문구를 입력하세요.",
            font=self.ui_font,
        )
        info_label.grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="w")

        entries = {}
        for row, key in enumerate(QUICK_INPUT_KEYS, start=1):
            ttk.Label(settings_win, text=f"Alt+{key} :", font=self.ui_font).grid(
                row=row, column=0, padx=(10, 5), pady=4, sticky="w"
            )
            entry = ttk.Entry(settings_win, width=40, font=self.ui_font)
            entry.insert(0, self.quick_inputs.get(key, ""))
            entry.grid(row=row, column=1, padx=(0, 10), pady=4, sticky="ew")
            entries[key] = entry

        entries[QUICK_INPUT_KEYS[0]].focus_set()

        def apply_and_save():
            for key, entry in entries.items():
                self.quick_inputs[key] = entry.get()
            self.save_quick_inputs()

        def apply_action():
            apply_and_save()
            settings_win.destroy()

        def cancel_action():
            settings_win.destroy()

        button_frame = ttk.Frame(settings_win)
        button_frame.grid(row=len(QUICK_INPUT_KEYS) + 1, column=0, columnspan=2, pady=10)
        apply_button = ttk.Button(button_frame, text="적용", command=apply_action, width=10)
        apply_button.pack(side=tk.LEFT, padx=5)
        cancel_button = ttk.Button(button_frame, text="취소", command=cancel_action, width=10)
        cancel_button.pack(side=tk.LEFT, padx=5)

    def bind_copy_shortcut(self):
        """현재 설정된 복사 단축키를 바인딩"""
        shortcut_map = {
            "Ctrl+Shift+C": "<Control-Shift-C>",
            "Ctrl+Alt+C": "<Control-Alt-c>",
            "Alt+Shift+C": "<Alt-Shift-C>"
        }
        binding = shortcut_map.get(self.copy_shortcut, "<Control-Shift-C>")
        self.root.bind(binding, self.copy_to_clipboard)

    def unbind_copy_shortcut(self):
        """모든 복사 단축키 바인딩 제거"""
        try:
            self.root.unbind("<Control-Shift-C>")
        except:
            pass
        try:
            self.root.unbind("<Control-Alt-c>")
        except:
            pass
        try:
            self.root.unbind("<Alt-Shift-C>")
        except:
            pass

    def import_memos(self):
        filepath = filedialog.askopenfilename(
            title="메모 파일 가져오기",
            filetypes=[("JSON 파일", "*.json"), ("모든 파일", "*.*")]
        )
        if not filepath:
            return
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                content = f.read()
                if not content.strip():
                    messagebox.showerror("오류", "파일이 비어있습니다.")
                    return
                new_memos = json.loads(content)
            if not isinstance(new_memos, list):
                raise TypeError("데이터가 리스트 형식이 아닙니다.")
            if not all(isinstance(m, dict) and "title" in m and "content" in m for m in new_memos):
                raise ValueError("일부 메모 항목의 구조가 올바르지 않습니다.")
            if messagebox.askyesno("확인", "기존 메모를 덮어쓰고 가져오시겠습니까?"):
                self.memos = new_memos
                self.save_memos()
                self.current_index = -1
                self.title_entry.delete(0, tk.END)
                self.content_text.delete("1.0", tk.END)
                self.toggle_right_panel(False)
                self.update_listbox()
                messagebox.showinfo("성공", "메모를 성공적으로 가져왔습니다.")
        except json.JSONDecodeError:
            messagebox.showerror("오류", "올바른 JSON 파일이 아닙니다.")
        except (TypeError, ValueError) as e:
            messagebox.showerror("오류", f"메모 파일 구조가 올바르지 않습니다: {e}")
        except Exception as e:
            messagebox.showerror("오류", f"파일을 가져오는 중 오류 발생: {e}")

    def export_memos(self):
        filepath = filedialog.asksaveasfilename(
            title="메모 내보내기", defaultextension=".json",
            filetypes=[("JSON 파일", "*.json"), ("텍스트 파일", "*.txt"), ("Excel 파일", "*.xlsx")]
        )
        if not filepath:
            return
        file_ext = os.path.splitext(filepath)[1].lower()
        try:
            if file_ext == ".json":
                self.save_memos()
                with open(self.file_path, 'r', encoding='utf-8') as f_in, open(filepath, 'w', encoding='utf-8') as f_out:
                    f_out.write(f_in.read())
            elif file_ext == ".txt":
                with open(filepath, "w", encoding="utf-8") as f:
                    for memo in self.memos:
                        f.write(f"제목: {memo['title']}\n" + "-"*20 + f"\n{memo['content']}\n\n" + "="*20 + "\n\n")
            elif file_ext == ".xlsx":
                if not openpyxl:
                    messagebox.showerror("오류", "Excel 내보내기에는 openpyxl이 필요합니다.")
                    return
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "메모"
                ws.append(["제목", "내용"])
                for memo in self.memos:
                    ws.append([memo["title"], memo["content"]])
                wb.save(filepath)
            messagebox.showinfo("성공", f"메모를 {filepath} 파일로 내보냈습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"파일 내보내기 중 오류 발생: {e}")

    def toggle_right_panel(self, enabled):
        state = tk.NORMAL if enabled else tk.DISABLED
        colors = THEME_COLORS.get(self.theme_mode, THEME_COLORS["light"])
        self.title_entry.config(state=state)
        self.content_text.config(state=state, bg=colors["bg"] if enabled else colors["disabled_bg"])
        self.copy_button.config(state=state)
        self.update_status_bar()

    def apply_classic_widget_colors(self):
        """ttk가 테마를 입히지 못하는 Listbox/Text 위젯에 현재 테마 색상을 적용.
        (Frame/Label/Entry/Button 등은 sv_ttk가 자동으로 처리하지만,
        Listbox와 Text는 ttk 위젯이 아니라서 색상을 직접 맞춰줘야 함)"""
        colors = THEME_COLORS.get(self.theme_mode, THEME_COLORS["light"])

        self.listbox.config(
            bg=colors["bg"], fg=colors["fg"],
            selectbackground=colors["list_select_bg"],
            selectforeground=colors["list_select_fg"],
            relief=tk.FLAT, borderwidth=0,
            highlightthickness=1,
            highlightbackground=colors["border"],
            highlightcolor=colors["accent"],
        )

        enabled = self.content_text.cget("state") == tk.NORMAL
        self.content_text.config(
            bg=colors["bg"] if enabled else colors["disabled_bg"],
            fg=colors["fg"],
            insertbackground=colors["insert_bg"],
            selectbackground=colors["text_select_bg"],
            selectforeground=colors["text_select_fg"],
            relief=tk.FLAT, borderwidth=0,
            highlightthickness=1,
            highlightbackground=colors["border"],
            highlightcolor=colors["accent"],
        )

    def load_memos(self):
        if not os.path.exists(self.file_path):
            return []
        try:
            with open(self.file_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            return []

    def save_memos(self):
        with open(self.file_path, "w", encoding="utf-8") as f:
            json.dump(self.memos, f, ensure_ascii=False, indent=4)

    def load_quick_inputs(self):
        """Alt+1~Alt+0에 대응하는 빠른 입력 문구를 불러옴 (없으면 전부 빈 문자열)"""
        empty = {key: "" for key in QUICK_INPUT_KEYS}
        if not os.path.exists(self.quick_input_file):
            return empty
        try:
            with open(self.quick_input_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, dict):
                return empty
            return {key: data.get(key, "") for key in QUICK_INPUT_KEYS}
        except (json.JSONDecodeError, IOError):
            return empty

    def save_quick_inputs(self):
        try:
            with open(self.quick_input_file, "w", encoding="utf-8") as f:
                json.dump(self.quick_inputs, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"❌ 빠른 입력 저장 실패: {e}")

    def update_listbox(self, preserve_scroll=False):
        """
        리스트박스를 업데이트하면서 스크롤 위치를 유지할 수 있는 옵션 추가
        
        Args:
            preserve_scroll: True이면 현재 스크롤 위치를 유지
        """
        # 스크롤 위치 저장
        if preserve_scroll:
            try:
                # yview()는 (첫번째_보이는_비율, 마지막_보이는_비율) 튜플 반환
                scroll_position = self.listbox.yview()
            except:
                scroll_position = None
        
        # 리스트박스 업데이트
        self.listbox.delete(0, tk.END)
        for memo in self.memos:
            self.listbox.insert(tk.END, memo["title"])
        
        # 스크롤 위치 복원
        if preserve_scroll and scroll_position:
            try:
                self.listbox.yview_moveto(scroll_position[0])
            except:
                pass
        
        self.update_status_bar()

    def on_memo_select(self, event):
        selected_indices = self.listbox.curselection()
        if not selected_indices:
            return
        self.current_index = selected_indices[0]
        memo = self.memos[self.current_index]
        self.toggle_right_panel(True)
        self.title_entry.delete(0, tk.END)
        self.title_entry.insert(0, memo["title"])
        self.content_text.delete("1.0", tk.END)
        self.content_text.insert("1.0", memo["content"])
        self.update_status_bar()

    def add_memo(self):
        new_memo = {"title": "새 메모", "content": ""}
        insert_pos = self.current_index + 1 if self.current_index != -1 else len(self.memos)
        self.memos.insert(insert_pos, new_memo)
        self.update_listbox()
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(insert_pos)
        self.listbox.activate(insert_pos)
        self.on_memo_select(None)
        self.save_memos()

    def remove_memo(self):
        if self.current_index == -1:
            messagebox.showwarning("경고", "삭제할 메모를 선택하세요.")
            return
        if messagebox.askyesno("확인", "선택한 메모를 제거하시겠습니까?"):
            del self.memos[self.current_index]
            self.current_index = -1
            self.title_entry.delete(0, tk.END)
            self.content_text.delete("1.0", tk.END)
            self.toggle_right_panel(False)
            self.update_listbox()
            self.save_memos()

    def move_memo_up(self):
        if self.current_index > 0:
            self.memos.insert(self.current_index - 1, self.memos.pop(self.current_index))
            self.current_index -= 1
            self.update_listbox_selection()

    def move_memo_down(self):
        if 0 <= self.current_index < len(self.memos) - 1:
            self.memos.insert(self.current_index + 1, self.memos.pop(self.current_index))
            self.current_index += 1
            self.update_listbox_selection()

    def update_listbox_selection(self):
        """메모 위치 변경 시 스크롤 위치를 유지하면서 업데이트"""
        # ✅ preserve_scroll=True로 스크롤 위치 유지
        self.update_listbox(preserve_scroll=True)
        self.listbox.selection_set(self.current_index)
        self.listbox.activate(self.current_index)
        # ✅ see()를 통해 이동된 항목이 보이도록 보장 (스크롤 최소 이동)
        self.listbox.see(self.current_index)
        self.on_memo_select(None)
        self.save_memos()

    def update_memo_realtime(self, event=None, update_list=True):
        """실시간으로 메모를 업데이트

        Args:
            update_list: True면 리스트박스에 보이는 제목도 함께 갱신한다.
                title_entry 입력처럼 화면에 보이는 제목이 바뀌는 경우에만 True로 호출하고,
                content_text 입력처럼 제목은 그대로인 경우 False로 호출하면 매 키 입력마다
                리스트박스 전체를 delete+재삽입하지 않아도 되어 더 가볍다.
        """
        if self.current_index == -1:
            return
        
        title = self.title_entry.get()
        content = self.content_text.get("1.0", tk.END).strip()
        self.memos[self.current_index] = {"title": title, "content": content}
        
        if update_list:
            # 전체 delete(0, END)+재삽입 대신, 바뀐 항목 하나만 갱신
            # (다른 항목은 그대로라 스크롤 위치도 자연히 유지됨)
            self.listbox.delete(self.current_index)
            self.listbox.insert(self.current_index, title)
            self.listbox.selection_set(self.current_index)
        
        self.save_memos()
        self.update_status_bar()

    def on_drag_start(self, event):
        """드래그 시작: 시작 인덱스 저장"""
        self.drag_start_index = self.listbox.nearest(event.y)
        # 기존 선택 이벤트가 먼저 발생하도록 함
        return

    def on_drag_motion(self, event):
        """드래그 중: 현재 위치 하이라이트"""
        if self.drag_start_index is None:
            return
        
        current_index = self.listbox.nearest(event.y)
        
        # 현재 마우스 위치의 항목을 시각적으로 표시
        if 0 <= current_index < len(self.memos):
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(current_index)

    def on_drag_drop(self, event):
        """드롭: 메모 순서 변경"""
        if self.drag_start_index is None:
            return
        
        drop_index = self.listbox.nearest(event.y)
        
        # 유효한 범위 체크
        if not (0 <= drop_index < len(self.memos)):
            self.drag_start_index = None
            return
        
        # 같은 위치면 무시
        if self.drag_start_index == drop_index:
            self.drag_start_index = None
            return
        
        # 메모 순서 변경
        memo = self.memos.pop(self.drag_start_index)
        self.memos.insert(drop_index, memo)
        
        # 현재 인덱스 업데이트
        self.current_index = drop_index
        
        # UI 업데이트
        self.update_listbox_selection()
        
        # 드래그 상태 초기화
        self.drag_start_index = None

    def on_closing(self):
        # 종료 전 설정 및 메모 저장 (창 크기, 위치 포함)
        self.save_memos()
        self.save_settings()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = MemoApp(root)
    root.mainloop()