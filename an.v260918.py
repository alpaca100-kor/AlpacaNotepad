import tkinter as tk
from tkinter import messagebox, filedialog, Toplevel, font, ttk
import json
import os
import re
import sys
import configparser
from datetime import datetime, date, timedelta
import calendar

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

# 실시간 자동저장 디바운스 지연시간(ms) - 이 시간 안에 새 키 입력이 있으면
# 저장을 다시 미루고, 입력이 멈추고 이 시간이 지나야 실제로 디스크에 씀
AUTOSAVE_DEBOUNCE_MS = 500

# 달력메모 탭 달력의 날짜 칸(Canvas) 크기(px)
CAL_CELL_W = 40
CAL_CELL_H = 34
# 달력메모 탭 왼쪽(달력) 영역의 고정 폭(px) - 월별보기/1년전체 폭을 통일하고,
# 사용자가 크기조절 막대로 바꿀 수 없도록 PanedWindow 대신 고정폭 Frame에 사용
CAL_LEFT_WIDTH = 360

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
        "sunday_fg": "#c42b1c",
        "muted_fg": "#9a9a9a",
        "has_memo_bg": "#cfe3f7",
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
        "sunday_fg": "#ff8a80",
        "muted_fg": "#6b6b6b",
        "has_memo_bg": "#1c3a52",
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
        self.calendar_file_path = os.path.join(get_app_dir(), "memos_calendar.json")
        self.holiday_file = os.path.join(get_app_dir(), "holidays.json")
        self.memos = self.load_memos()
        self.settings = self.load_settings()
        self.quick_inputs = self.load_quick_inputs()
        self.calendar_memos = self.load_calendar_memos()
        self.holidays = self.load_holidays()  # {"YYYY-MM-DD": "공휴일 이름"} - 참고용, 앱이 쓰지는 않음
        self._holiday_tooltip = None  # 공휴일 이름 풍선말 Toplevel (없으면 None)
        self.current_index = -1
        # 디바운스된 자동저장 예약을 key별로 추적 ({key: after_id})
        self._pending_save_ids = {}
        # 자동저장이 실패한 key를 기록 - 같은 원인(디스크 꽉 참, 권한 없음 등)이
        # 계속되는 동안 키 입력마다 반복해서 오류 팝업이 뜨지 않도록, 이미 알린
        # key는 그 저장이 다시 성공하기 전까지는 조용히 넘어감
        self._save_failed_keys = set()

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

        # 메뉴막대 아래 [일반메모]/[달력메모] 탭 - 하위의 [월별보기]/[1년전체]
        # 탭과 똑같아 보이지 않도록, 최상위 탭 전용 스타일(더 큰 볼드체 + 넉넉한 여백)을
        # 적용해 시각적 위계를 분명히 함. (색상이 아닌 폰트/여백만 다르게 하는 이유:
        # sv_ttk는 자기가 아는 기본 스타일 이름만 관리하므로, 이렇게 별도 이름으로
        # 만든 커스텀 스타일은 라이트/다크 전환 시에도 다시 적용할 필요 없이 유지됨)
        top_tab_style = ttk.Style()
        top_tab_style.configure("TopLevel.TNotebook.Tab", font=("맑은 고딕", 13, "bold"),
                                padding=(24, 4))
        top_tab_style.configure("TopLevel.TNotebook", tabmargins=(6, 8, 6, 0))

        self.notebook = ttk.Notebook(root, style="TopLevel.TNotebook")
        self.notebook.pack(fill=tk.BOTH, expand=True)
        self.weekday_label_widgets = []  # (라벨위젯, "sun"|"sat"|"normal") - 테마 갱신용

        general_tab = ttk.Frame(self.notebook)
        calendar_tab = ttk.Frame(self.notebook)
        self.notebook.add(general_tab, text="📝 일반메모")
        self.notebook.add(calendar_tab, text="📅 달력메모")

        self.main_pane = ttk.PanedWindow(general_tab, orient=tk.HORIZONTAL)
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
        # height=1: 지정 안 하면 Text 기본값(24줄) 기준으로 자연 요구 크기가 매우 커져서,
        # 창이 작을 때 fill/expand로 실제 크기가 줄어들어도 레이아웃 계산에서 아래
        # [복사] 버튼 등 형제 위젯이 창 밖으로 밀려나는 원인이 됨. height=1로 자연
        # 요구 크기를 최소화하고, 실제 크기는 fill=BOTH+expand=True가 결정하게 함
        self.content_text = tk.Text(right_panel, font=self.content_font, padx=10, pady=8, height=1)
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

        # 달력메모 탭 구성
        self.build_calendar_tab(calendar_tab)

        # 단축키 등록
        self.root.bind("<Control-n>", self._on_ctrl_n_key)
        self.root.bind("<Control-d>", self._on_ctrl_d_key)
        self.root.bind("<Control-Tab>", lambda event: self._cycle_top_tab(1))
        self.root.bind("<Control-Shift-Tab>", lambda event: self._cycle_top_tab(-1))
        self.root.bind("<Control-m>", self._on_ctrl_m_key)
        # 고급 사용자용 숨김 기능: 버튼 없이 단축키로만 진입 (탭 구분 없이 항상 동작)
        self.root.bind("<Control-Shift-D>", self.open_bulk_delete_dialog)
        self.root.bind("<Alt-m>", self._on_alt_m_key)
        self.root.bind("<Alt-y>", self._on_alt_y_key)
        self.root.bind("<Alt-Left>", self._on_alt_left_key)
        self.root.bind("<Alt-Right>", self._on_alt_right_key)
        self.root.bind("<Prior>", self._on_prior_key)
        self.root.bind("<Next>", self._on_next_key)
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
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

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
            # geometry 문자열 검증: "800x600+100+100" 형식.
            # 좌표는 +100처럼 양수뿐 아니라 -50처럼 음수(창이 화면 왼쪽/위쪽 바깥에
            # 걸쳐 있던 경우)일 수도 있어, 단순히 '+' 문자 존재 여부만 보면 x/y가
            # 둘 다 음수인 "800x600-50-30" 같은 유효한 형식을 잘못된 것으로 오판해
            # 위치가 조용히 기본값으로 리셋되는 문제가 있었음. 정규식으로 폭x높이,
            # 그리고 +/- 부호가 붙은 x,y 좌표까지 정확히 매치하는지 확인함.
            if geometry and re.fullmatch(r"\d+x\d+[+-]\d+[+-]\d+", geometry):
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
        self.main_pane.bind("<ButtonRelease-1>", lambda e: self.enforce_min_sash(e, self.main_pane, 200))

    def enforce_min_sash(self, event=None, pane=None, min_width=200):
        """왼쪽 리스트 패널(일반메모 탭)이 너무 좁아지지 않도록 최소 폭을 보정.
        (달력메모 탭은 크기조절이 불가능한 고정 폭이라 이 함수를 쓰지 않음)"""
        pane = pane if pane is not None else self.main_pane
        try:
            if pane.sashpos(0) < min_width:
                pane.sashpos(0, min_width)
        except Exception:
            pass

    def create_menu(self):
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="일반메모 가져오기...", command=self.import_memos)
        file_menu.add_command(label="일반메모 내보내기...", command=self.export_memos)
        file_menu.add_separator()
        file_menu.add_command(label="달력메모 가져오기...", command=self.import_calendar_memos)
        file_menu.add_command(label="달력메모 내보내기...", command=self.export_calendar_memos)
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

    def _on_alt_m_key(self, event=None):
        """Alt+M: 달력메모 탭이 활성화된 상태에서만, [월별보기] 서브탭 활성화
        (Ctrl+M은 이미 "메모 내용 편집창에 포커스"로 양쪽 탭에서 쓰이고 있어서
        서브탭 전환은 Alt 계열로 분리함)"""
        if self._is_general_tab_active():
            return
        self.cal_view_notebook.select(0)
        return "break"

    def _on_alt_y_key(self, event=None):
        """Alt+Y: 달력메모 탭이 활성화된 상태에서만, [1년전체] 서브탭 활성화"""
        if self._is_general_tab_active():
            return
        self.cal_view_notebook.select(1)
        return "break"

    def _on_alt_left_key(self, event=None):
        """Alt+왼쪽 방향키: 달력메모 탭이 활성화된 상태에서만, 월별보기면 이전 달로,
        1년전체면 이전 해로 이동"""
        if self._is_general_tab_active():
            return
        if self._is_year_view_active():
            self.go_to_year(self.cal_year_year - 1)
        else:
            self.go_to_month(self.cal_year, self.cal_month - 1)
        return "break"

    def _on_alt_right_key(self, event=None):
        """Alt+오른쪽 방향키: _on_alt_left_key 참고 (반대 방향)"""
        if self._is_general_tab_active():
            return
        if self._is_year_view_active():
            self.go_to_year(self.cal_year_year + 1)
        else:
            self.go_to_month(self.cal_year, self.cal_month + 1)
        return "break"

    def _cycle_top_tab(self, direction):
        """Ctrl+Tab(+1)/Ctrl+Shift+Tab(-1): 일반메모/달력메모 탭 전환.
        (ttk.Notebook 자체의 기본 Ctrl+Tab 처리는 노트북 위젯 본인이 포커스일
        때만 동작해 평소 편집 중엔 적용되지 않으므로, root 레벨에서 직접 처리함)"""
        tabs = self.notebook.tabs()
        current = self.notebook.index(self.notebook.select())
        self.notebook.select(tabs[(current + direction) % len(tabs)])
        return "break"

    def _on_ctrl_m_key(self, event=None):
        """Ctrl+M: 현재 탭에 맞는 메모 내용 편집창에 포커스만 이동
        (일반메모 탭: content_text / 달력메모 탭: date_content_text).
        Ctrl+T(제목)와 달리 전체선택은 하지 않고 포커스만 옮김"""
        if self._is_general_tab_active():
            if str(self.content_text.cget("state")) == tk.NORMAL:
                self.content_text.focus_set()
        else:
            self.date_content_text.focus_set()
        return "break"

    def focus_on_listbox(self, event=None):
        if not self._is_general_tab_active():
            return "break"
        self.listbox.focus_set()
        if self.current_index != -1:
            self.listbox.selection_set(self.current_index)
            self.listbox.activate(self.current_index)
        return "break"

    def focus_on_title(self, event=None):
        if not self._is_general_tab_active():
            # 달력메모 탭에서 Ctrl+T: 월별보기면 [오늘], 1년전체면 [올해] 버튼과 동일하게 동작
            if self._is_year_view_active():
                self.go_to_year(datetime.now().year)
            else:
                self.go_to_today()
            return "break"
        # ttk.Entry의 cget('state')는 일반 str이 아닌 Tcl 객체를 반환하므로 str()로 변환 후 비교해야 함
        if str(self.title_entry.cget('state')) == tk.NORMAL:
            self.title_entry.focus_set()
            self.title_entry.select_range(0, tk.END)
        return "break"

    def _insert_text_at_focus(self, text):
        """포커스된 위젯(제목/내용/달력메모 내용)의 커서 위치에 문자열을 삽입하는 공통 로직.
        (Alt+T 날짜/시간 삽입, Alt+숫자 빠른 입력에서 공용으로 사용)
        제목 입력창/메모 내용/달력메모 내용, 이 세 위젯 중 하나에 실제로 포커스가
        있을 때만 삽입한다. (이전에는 셋 다 아니면 무조건 content_text 끝에 삽입해서,
        예를 들어 일반메모 편집 중 달력메모 탭으로 이동해 아무 것도 포커스하지 않은
        상태로 Alt+T/Alt+숫자를 눌러도 일반메모의 마지막 메모에 엉뚱하게 삽입되는
        문제가 있었음 - PageUp/PageDown이 다른 위젯에서도 리스트 순서를 바꾸던
        버그와 같은 종류의 문제)"""
        focused = self.root.focus_get()

        if focused is self.title_entry:
            try:
                self.title_entry.delete("sel.first", "sel.last")
            except tk.TclError:
                pass
            self.title_entry.insert(tk.INSERT, text)
            self.update_memo_realtime(update_list=True)
            return

        if focused is self.date_content_text:
            try:
                self.date_content_text.delete("sel.first", "sel.last")
            except tk.TclError:
                pass
            self.date_content_text.insert(tk.INSERT, text)
            self.save_date_memo_realtime()
            return

        if focused is self.content_text:
            try:
                self.content_text.delete("sel.first", "sel.last")
            except tk.TclError:
                pass
            self.content_text.insert(tk.INSERT, text)
            self.update_memo_realtime(update_list=False)
            return

        # 위 세 위젯 중 어디에도 포커스가 없으면 아무 것도 하지 않음

    def insert_datetime(self, event=None):
        """Alt+T: 포커스된 위젯의 커서 위치에 현재 날짜/시간을 삽입 (yyyy-mm-dd hh:mm:ss)"""
        if self._is_general_tab_active():
            if self.current_index == -1:
                return "break"
        elif not self.selected_date:
            return "break"
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._insert_text_at_focus(now_str)
        return "break"

    def insert_quick_text(self, key, event=None):
        """Alt+1~Alt+0: 설정 메뉴 > 빠른 입력 설정에서 미리 지정한 문구를 삽입"""
        if self._is_general_tab_active():
            if self.current_index == -1:
                return "break"
        elif not self.selected_date:
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
        """활성 탭에 맞는 내용(일반메모의 content_text 또는 달력메모의 date_content_text)을 복사"""
        if self._is_general_tab_active():
            if self.current_index == -1:
                return "break"
            return self._copy_widget_text_to_clipboard(self.content_text, self.copy_status_label)
        if not self.selected_date:
            return "break"
        return self._copy_widget_text_to_clipboard(self.date_content_text, self.date_copy_status_label)

    def _copy_widget_text_to_clipboard(self, text_widget, status_label):
        try:
            text_to_copy = text_widget.get("1.0", tk.END).strip()
            if not text_to_copy:
                messagebox.showinfo("알림", "복사할 내용이 없습니다.", parent=self.root)
                return "break"
            self.root.clipboard_clear()
            self.root.clipboard_append(text_to_copy)

            # "복사 완료!" 메시지 표시
            self._show_copy_success(status_label)

        except tk.TclError:
            messagebox.showerror("오류", "클립보드에 접근할 수 없습니다.", parent=self.root)
        except Exception as e:
            messagebox.showerror("오류", f"알 수 없는 오류 발생: {e}", parent=self.root)
        return "break"

    def _show_copy_success(self, label):
        """복사 완료 메시지를 1초 동안 표시 (라이트=파란색, 다크=노란색)"""
        colors = THEME_COLORS.get(self.theme_mode, THEME_COLORS["light"])
        label.config(text="복사 완료!", foreground=colors["success_fg"])
        # 1초(1000ms) 후에 메시지 제거
        self.root.after(1000, lambda: label.config(text=""))

    def update_status_bar(self, event=None):
        try:
            if self._is_general_tab_active():
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
            else:
                memo_count = len(self.calendar_memos)
                status_text = f"총 {memo_count}개의 메모"
                if self.selected_date:
                    cursor_pos = self.date_content_text.index(tk.INSERT)
                    row, column = cursor_pos.split('.')
                    row = int(row)
                    column = int(column) + 1
                    content = self.date_content_text.get("1.0", tk.END).strip()
                    char_count = len(content)
                    status_text += f" / 선택한 날짜: {self.selected_date} / 커서 위치: 줄 {row}, 열 {column} / 글자 수: {char_count}자"
            self.status_bar.config(text=status_text)
        except Exception:
            pass

    def _atomic_write(self, path, write_func):
        """write_func(파일객체)로 실제 내용을 쓰되, 같은 폴더의 임시 파일에
        먼저 쓴 뒤 os.replace()로 교체하는 원자적 저장. os.replace()는 파일
        내용을 옮기는 게 아니라 이름(포인터)만 바꾸는 연산이라 OS가 원자적으로
        처리하므로, 쓰는 도중 앱이 강제 종료되거나 오류가 나도 원본 파일(path)은
        손상되지 않고 그대로 남음. (임시 파일은 반드시 같은 폴더에 있어야
        os.replace()의 원자성이 보장됨)"""
        tmp_path = path + ".tmp"
        try:
            with open(tmp_path, "w", encoding="utf-8") as f:
                write_func(f)
            os.replace(tmp_path, path)
        except Exception:
            # write_func 내부 오류(TypeError 등)든 디스크 I/O 오류(OSError)든
            # 종류를 가리지 않고 원본 파일은 아직 손대지 않았으므로 안전함.
            # 남은 임시 파일만 정리 시도 후 예외를 그대로 위로 전달.
            try:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
            except OSError:
                pass
            raise

    def _atomic_write_json(self, path, data):
        """JSON 데이터를 원자적으로 저장 (_atomic_write 참고)"""
        self._atomic_write(path, lambda f: json.dump(data, f, ensure_ascii=False, indent=4))

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
            if not (8 <= font_size <= 72):
                # settings.ini가 손상되거나 직접 편집되어 범위를 벗어난 값(예: 음수)이면
                # 글꼴 설정창의 허용 범위(8~72)를 기준으로 기본값으로 되돌림
                font_size = default_settings['font_size']
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
            self._atomic_write(self.settings_file, config.write)
            print(f"✅ 설정 저장: {self.root.geometry()}")
        except Exception as e:
            print(f"❌ 설정 저장 실패: {e}")

    def open_bulk_delete_dialog(self, event=None):
        """Ctrl+Shift+D (고급 사용자용 숨김 기능, 버튼 없음): 특정 날짜 이전에
        저장된 달력메모를 한꺼번에 삭제하는 팝업. 일반메모는 작성/저장 날짜를
        따로 기록하지 않는 데이터 구조라 이 기능의 대상이 될 수 없으므로,
        달력메모(날짜별로 저장되는 메모)에만 적용됨.
        달력메모 탭이 활성화된 상태에서만 동작함."""
        if self._is_general_tab_active():
            return
        dialog = Toplevel(self.root)
        dialog.title("메모 일괄 삭제")
        dialog.geometry("380x250")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.focus_force()

        default_date = (date.today() - timedelta(days=1)).strftime("%Y-%m-%d")
        date_var = tk.StringVar(value=default_date)

        ttk.Label(dialog, text="기준 날짜:", font=self.ui_font).grid(
            row=0, column=0, padx=10, pady=(15, 5), sticky="w")
        date_spin = ttk.Spinbox(dialog, textvariable=date_var, width=12, font=self.ui_font)
        date_spin.grid(row=0, column=1, padx=10, pady=(15, 5), sticky="w")
        date_spin.focus_set()
        date_spin.icursor(tk.END)

        def adjust_date(delta):
            """스핀박스 위/아래 버튼: 날짜를 하루씩 증가/감소 (월/연 경계도 올바르게 처리)"""
            try:
                d = datetime.strptime(date_var.get().strip(), "%Y-%m-%d").date()
            except ValueError:
                d = date.today() - timedelta(days=1)
            date_var.set((d + timedelta(days=delta)).strftime("%Y-%m-%d"))

        # ttk.Spinbox의 기본 클래스 바인딩(TSpinbox)이 <<Increment>>/<<Decrement>>에서
        # 텍스트를 숫자로 취급해 자체적으로 증감을 시도하는데(break 없음), 그대로 두면
        # 우리가 막 넣은 날짜 문자열을 다시 덮어써버려서(예: "0") "break"로 반드시 막아야 함
        def on_increment(e=None):
            adjust_date(1)
            return "break"

        def on_decrement(e=None):
            adjust_date(-1)
            return "break"

        date_spin.bind("<<Increment>>", on_increment)
        date_spin.bind("<<Decrement>>", on_decrement)

        colors = THEME_COLORS.get(self.theme_mode, THEME_COLORS["light"])
        desc = (
            "선택한 날짜를 포함하여 그 이전 날짜에 저장된 달력메모 내용을\n"
            "모두 삭제합니다.\n\n"
            "※ 삭제 후에는 되돌릴 수 없으니 날짜를 확인한 뒤 실행하세요."
        )
        ttk.Label(dialog, text=desc, font=("맑은 고딕", 9), foreground=colors["sunday_fg"],
                  justify=tk.LEFT).grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="w")

        def on_confirm():
            raw = date_var.get().strip()
            try:
                cutoff = datetime.strptime(raw, "%Y-%m-%d").date().strftime("%Y-%m-%d")
            except ValueError:
                messagebox.showerror("오류", "날짜를 YYYY-MM-DD 형식으로 입력하세요.", parent=dialog)
                return
            to_delete = [d for d in self.calendar_memos if d <= cutoff]
            for d in to_delete:
                del self.calendar_memos[d]
            self.save_calendar_memos()
            self._refresh_calendar_views()
            self.on_calendar_date_click(self.selected_date)
            dialog.destroy()
            if to_delete:
                messagebox.showinfo("완료", f"{len(to_delete)}개의 달력메모를 삭제했습니다.")
            else:
                messagebox.showinfo("완료", "삭제할 메모가 없습니다.")

        def on_cancel():
            dialog.destroy()

        dialog.bind("<Escape>", lambda e: on_cancel())

        button_frame = ttk.Frame(dialog)
        button_frame.grid(row=2, column=0, columnspan=2, pady=15)
        ttk.Button(button_frame, text="확인", command=on_confirm, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="취소", command=on_cancel, width=10).pack(side=tk.LEFT, padx=5)

    def open_font_settings(self):
        settings_win = Toplevel(self.root)
        settings_win.title("글꼴 설정")
        settings_win.geometry("350x150")
        settings_win.resizable(False, False)
        settings_win.transient(self.root)
        settings_win.grab_set()
        settings_win.focus_force()
        # cancel_action은 이 함수 안에서 나중에 정의되지만, 람다는 호출되는 시점(Esc를
        # 누르는 시점)에 이름을 찾으므로 여기서 미리 바인딩해도 문제 없음
        settings_win.bind("<Escape>", lambda e: cancel_action())

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
                if hasattr(self, "date_content_text"):
                    self.date_content_text.config(font=self.content_font)
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
        # cancel_action은 이 함수 안에서 나중에 정의되지만, 람다는 호출되는 시점(Esc를
        # 누르는 시점)에 이름을 찾으므로 여기서 미리 바인딩해도 문제 없음
        settings_win.bind("<Escape>", lambda e: cancel_action())

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
        # cancel_action은 이 함수 안에서 나중에 정의되지만, 람다는 호출되는 시점(Esc를
        # 누르는 시점)에 이름을 찾으므로 여기서 미리 바인딩해도 문제 없음
        settings_win.bind("<Escape>", lambda e: cancel_action())

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
        # cancel_action은 이 함수 안에서 나중에 정의되지만, 람다는 호출되는 시점(Esc를
        # 누르는 시점)에 이름을 찾으므로 여기서 미리 바인딩해도 문제 없음
        settings_win.bind("<Escape>", lambda e: cancel_action())

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
            # 키가 존재하는지만 보면 title이 숫자거나 content가 객체인 경우도 통과해서
            # 리스트박스/Text에 지저분한 값이 그대로 들어갈 수 있으므로, 값의 타입까지
            # 확인함 (내부 memos.json 로딩(load_memos)과 달리, 가져오기는 사용자가
            # 명시적으로 확인 후 실행하는 동작이라 조용히 보정하지 않고 오류로 알림)
            for m in new_memos:
                if not isinstance(m, dict):
                    raise ValueError("메모 항목이 딕셔너리가 아닙니다.")
                if not isinstance(m.get("title"), str):
                    raise ValueError("메모 제목은 문자열이어야 합니다.")
                if not isinstance(m.get("content"), str):
                    raise ValueError("메모 내용은 문자열이어야 합니다.")
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
            else:
                # .json/.txt/.xlsx가 아닌 확장자로 저장을 시도한 경우(예: 저장 대화상자에서
                # 사용자가 직접 다른 확장자를 입력) - 아무 것도 저장하지 않았는데 아래
                # "성공" 메시지가 뜨는 것을 막기 위해 명시적으로 오류 처리함
                raise ValueError(
                    f"지원하지 않는 파일 형식입니다: {file_ext or '(확장자 없음)'}\n"
                    ".json / .txt / .xlsx 중 하나로 저장해주세요."
                )
            messagebox.showinfo("성공", f"메모를 {filepath} 파일로 내보냈습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"파일 내보내기 중 오류 발생: {e}")

    def import_calendar_memos(self):
        """"달력메모 가져오기...": memos_calendar.json과 같은 {"YYYY-MM-DD": "내용"} 형태의
        JSON 파일을 가져와 기존 달력메모를 덮어씀"""
        filepath = filedialog.askopenfilename(
            title="달력메모 파일 가져오기",
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
                new_calendar_memos = json.loads(content)
            if not isinstance(new_calendar_memos, dict):
                raise TypeError('데이터가 {"날짜": "내용"} 형태의 딕셔너리가 아닙니다.')
            for k, v in new_calendar_memos.items():
                datetime.strptime(k, "%Y-%m-%d")  # 날짜 형식(YYYY-MM-DD) 검증
                if not isinstance(v, str):
                    raise ValueError("일부 항목의 내용이 문자열이 아닙니다.")
            if messagebox.askyesno("확인", "기존 달력메모를 덮어쓰고 가져오시겠습니까?"):
                self.calendar_memos = new_calendar_memos
                self.save_calendar_memos()
                self.date_content_text.delete("1.0", tk.END)
                self.date_content_text.insert("1.0", self.calendar_memos.get(self.selected_date, ""))
                self.render_month_view()
                self.render_year_view()
                self.update_status_bar()
                messagebox.showinfo("성공", "달력메모를 성공적으로 가져왔습니다.")
        except json.JSONDecodeError:
            messagebox.showerror("오류", "올바른 JSON 파일이 아닙니다.")
        except (TypeError, ValueError) as e:
            messagebox.showerror("오류", f"달력메모 파일 구조가 올바르지 않습니다: {e}")
        except Exception as e:
            messagebox.showerror("오류", f"파일을 가져오는 중 오류 발생: {e}")

    def export_calendar_memos(self):
        """"달력메모 내보내기...": JSON/텍스트/Excel로 내보냄 (날짜 오름차순 정렬)"""
        filepath = filedialog.asksaveasfilename(
            title="달력메모 내보내기", defaultextension=".json",
            filetypes=[("JSON 파일", "*.json"), ("텍스트 파일", "*.txt"), ("Excel 파일", "*.xlsx")]
        )
        if not filepath:
            return
        file_ext = os.path.splitext(filepath)[1].lower()
        sorted_items = sorted(self.calendar_memos.items())
        try:
            if file_ext == ".json":
                self.save_calendar_memos()
                with open(self.calendar_file_path, 'r', encoding='utf-8') as f_in, open(filepath, 'w', encoding='utf-8') as f_out:
                    f_out.write(f_in.read())
            elif file_ext == ".txt":
                with open(filepath, "w", encoding="utf-8") as f:
                    for date_key, content in sorted_items:
                        f.write(f"{date_key}\n" + "-"*20 + f"\n{content}\n\n" + "="*20 + "\n\n")
            elif file_ext == ".xlsx":
                if not openpyxl:
                    messagebox.showerror("오류", "Excel 내보내기에는 openpyxl이 필요합니다.")
                    return
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "달력메모"
                ws.append(["날짜", "내용"])
                for date_key, content in sorted_items:
                    ws.append([date_key, content])
                wb.save(filepath)
            else:
                raise ValueError(
                    f"지원하지 않는 파일 형식입니다: {file_ext or '(확장자 없음)'}\n"
                    ".json / .txt / .xlsx 중 하나로 저장해주세요."
                )
            messagebox.showinfo("성공", f"달력메모를 {filepath} 파일로 내보냈습니다.")
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
        """ttk가 테마를 입히지 못하는 Listbox/Text/달력 Canvas 위젯에 현재 테마 색상을 적용.
        (Frame/Label/Entry/Button 등은 sv_ttk가 자동으로 처리하지만,
        Listbox와 Text, 달력의 Canvas/Label은 ttk 위젯이 아니라서 색상을 직접 맞춰줘야 함)"""
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

        # 달력메모 탭의 내용 입력창 (생성된 이후에만 존재)
        if hasattr(self, "date_content_text"):
            self.date_content_text.config(
                bg=colors["bg"], fg=colors["fg"],
                insertbackground=colors["insert_bg"],
                selectbackground=colors["text_select_bg"],
                selectforeground=colors["text_select_fg"],
                relief=tk.FLAT, borderwidth=0,
                highlightthickness=1,
                highlightbackground=colors["border"],
                highlightcolor=colors["accent"],
            )

        # 달력 요일 헤더(일~토) 라벨 색상
        for lbl, kind in getattr(self, "weekday_label_widgets", []):
            if kind == "sun":
                lbl.config(bg=colors["bg"], fg=colors["sunday_fg"])
            elif kind == "sat":
                lbl.config(bg=colors["bg"], fg=colors["accent"])
            else:
                lbl.config(bg=colors["bg"], fg=colors["fg"])

        # 1년전체의 스크롤바 (항상 뚜렷하게 보이도록 기본 Tk 스크롤바 사용 - 테마색 적용)
        if hasattr(self, "year_scrollbar"):
            self.year_scrollbar.config(
                bg=colors["border"], troughcolor=colors["bg"],
                activebackground=colors["accent"], highlightthickness=0,
                relief=tk.FLAT, borderwidth=0,
            )

        # 달력 그리드: 칸 사이 여백이 은은한 격자선처럼 보이도록 배경색을 맞추고 다시 그림
        if hasattr(self, "month_grid_frame"):
            self.month_grid_frame.config(bg=colors["border"])
            self.render_month_view()
        if hasattr(self, "year_grid_frame"):
            self.year_grid_frame.config(bg=colors["border"])
            self.render_year_view()

    def _debounced_save(self, key, save_func, delay_ms=AUTOSAVE_DEBOUNCE_MS):
        """key로 구분되는 저장을 delay_ms 뒤로 미루고, 그 사이에 또 호출되면
        이전 예약을 취소하고 다시 미룸. (연속으로 키 입력이 들어오는 동안은
        디스크에 쓰지 않다가, 입력이 멈추고 delay_ms가 지나야 실제로 한 번
        저장 - 매 키 입력마다 파일 전체를 쓰는 부담을 줄임)"""
        existing = self._pending_save_ids.pop(key, None)
        if existing is not None:
            self.root.after_cancel(existing)

        def _run():
            self._pending_save_ids.pop(key, None)
            result = save_func()
            # save_func가 명시적으로 False를 반환한 경우만 저장 실패로 간주함
            # (달력 그리드 다시 그리기처럼 반환값이 없는 콜백은 영향받지 않음)
            if result is False:
                self._notify_save_failure(key)

        self._pending_save_ids[key] = self.root.after(delay_ms, _run)

    def _notify_save_failure(self, key):
        """자동저장 실패를 사용자에게 알림. 콘솔 출력만으로는 콘솔 없이 실행되는
        빌드(exe)에서 사용자가 저장 실패 사실을 전혀 알 수 없으므로 팝업으로 알림.
        단, 같은 원인(디스크 꽉 참 등)으로 계속 실패하는 동안 키 입력마다 팝업이
        반복해서 뜨지 않도록, 이미 알린 key는 그 저장이 다시 성공할 때까지 무시함."""
        if key in self._save_failed_keys:
            return
        self._save_failed_keys.add(key)
        label = {"memos": "일반메모", "calendar_memos": "달력메모"}.get(key, key)
        messagebox.showerror(
            "저장 실패",
            f"{label} 자동저장에 실패했습니다.\n"
            "디스크 공간이나 저장 폴더의 쓰기 권한을 확인해주세요.\n"
            "입력한 내용은 화면에 남아있지만 파일에는 아직 저장되지 않았습니다.",
            parent=self.root,
        )

    def _cancel_pending_saves(self):
        """예약된 디바운스 저장을 전부 취소만 함 (뒤이어 직접·즉시 저장을
        호출할 예정일 때 사용 - 예: 종료 직전)"""
        for after_id in self._pending_save_ids.values():
            self.root.after_cancel(after_id)
        self._pending_save_ids.clear()

    def load_memos(self):
        if not os.path.exists(self.file_path):
            return []
        try:
            with open(self.file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, list):
                return []
            # memos.json이 외부에서 손상되거나 title/content 키가 없는 항목이 섞여
            # 있어도, 이후 코드(on_memo_select 등)가 두 키가 항상 있다고 가정하고
            # 접근하다가 KeyError로 죽는 일이 없도록 여기서 구조를 정규화함
            valid_memos = []
            for memo in data:
                if not isinstance(memo, dict):
                    continue
                title = memo.get("title", "")
                content = memo.get("content", "")
                if not isinstance(title, str):
                    title = str(title)
                if not isinstance(content, str):
                    content = str(content)
                valid_memos.append({"title": title, "content": content})
            return valid_memos
        except (json.JSONDecodeError, IOError):
            return []

    def save_memos(self):
        try:
            self._atomic_write_json(self.file_path, self.memos)
            self._save_failed_keys.discard("memos")
            return True
        except Exception as e:
            print(f"❌ 메모 저장 실패: {e}")
            return False

    def load_holidays(self):
        """대한민국 공휴일 정보를 holidays.json에서 불러옴 (없거나 손상되었으면 빈 딕셔너리).
        {"YYYY-MM-DD": "공휴일 이름"} 형태이며, memos_calendar.json과 달리 앱이 이 파일에
        쓰지는 않는 참고용 데이터임 - 최신 연도를 쓰려면 이 파일을 직접 교체/추가해야 함"""
        if not os.path.exists(self.holiday_file):
            return {}
        try:
            with open(self.holiday_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        except (json.JSONDecodeError, OSError):
            return {}

    def load_calendar_memos(self):
        """달력메모(memos_calendar.json)를 불러옴: {"YYYY-MM-DD": "내용", ...} 형태.
        (가져오기 기능(import_calendar_memos)과 동일한 수준으로 날짜 형식과 값이
        문자열인지 검증함 - 둘의 검증 수준이 달라 파일을 직접 열었을 때만 이상한
        데이터가 통과하는 일이 없도록 함)"""
        if not os.path.exists(self.calendar_file_path):
            return {}
        try:
            with open(self.calendar_file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, dict):
                return {}
            valid_memos = {}
            for date_key, content in data.items():
                try:
                    datetime.strptime(date_key, "%Y-%m-%d")
                except (ValueError, TypeError):
                    continue
                if isinstance(content, str):
                    valid_memos[date_key] = content
            return valid_memos
        except (json.JSONDecodeError, IOError):
            return {}

    def save_calendar_memos(self):
        try:
            self._atomic_write_json(self.calendar_file_path, dict(sorted(self.calendar_memos.items())))
            self._save_failed_keys.discard("calendar_memos")
            return True
        except Exception as e:
            print(f"❌ 달력메모 저장 실패: {e}")
            return False

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
            self._atomic_write_json(self.quick_input_file, self.quick_inputs)
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

    def _is_listbox_focused(self):
        """PageUp/PageDown 순서변경 단축키를 리스트박스에 포커스가 있을 때만
        허용하기 위한 확인. (Text/Entry의 기본 Prior/Next 바인딩은 break를
        호출하지 않아 이벤트가 그대로 root까지 전파되는 tkinter 특성 때문에,
        이 확인이 없으면 메모 내용을 스크롤하려고 PageUp/PageDown을 누르는
        것만으로 (심지어 다른 탭에서도) 메모 순서가 같이 바뀌는 문제가 있었음)"""
        return self.root.focus_get() is self.listbox

    def _on_ctrl_n_key(self, event=None):
        """Ctrl+N: 일반메모 탭이 활성화된 상태에서만 새 메모를 추가.
        (다른 대부분의 단축키처럼 _is_general_tab_active()를 확인함 - 이 확인이
        없으면 달력메모 탭에서 무심코 Ctrl+N을 눌렀을 때 화면엔 아무 변화가 없는데
        일반메모 목록에 빈 메모가 조용히 하나 추가되는 문제가 있었음)"""
        if not self._is_general_tab_active():
            return "break"
        self.add_memo()
        return "break"

    def _on_ctrl_d_key(self, event=None):
        """Ctrl+D: 일반메모 탭에서는 제목/내용에 포커스가 있을 때 선택된 메모를
        삭제, 달력메모 탭에서는 (포커스 위치와 무관하게) 선택된 날짜의 메모
        내용을 삭제. 메모 목록 자체에 포커스가 있을 때는 이미 Delete 키가 그
        역할을 하고 있으므로 Ctrl+D는 관여하지 않음 (Alt+T/PageUp/PageDown과
        같은, 포커스·탭 상태를 직접 확인하는 방식)"""
        if not self._is_general_tab_active():
            self.remove_date_memo()
            return "break"
        focused = self.root.focus_get()
        if focused is self.title_entry or focused is self.content_text:
            self.remove_memo()
            return "break"

    def _on_prior_key(self, event=None):
        """전역 PageUp 단축키: 달력메모 탭이 활성화된 상태면 (어느 위젯에 포커스가
        있든) 항상 이전 메모 날짜로 이동. 일반메모 탭에서는 리스트박스에 포커스가
        있을 때만 메모를 위로 이동.
        (date_content_text에 포커스가 있는 경우는 그 위젯의 인스턴스 바인딩이
        먼저 가로채 break로 처리하므로 이 함수까지 전파되지 않지만, 그 외 달력
        탭 내 다른 위젯(버튼, 달력 칸 등)에 포커스가 있는 경우를 위해 필요함)"""
        if not self._is_general_tab_active():
            self.go_to_adjacent_memo_date(-1)
            return "break"
        if self._is_listbox_focused():
            self.move_memo_up()

    def _on_next_key(self, event=None):
        """전역 PageDown 단축키: _on_prior_key 참고 (반대 방향)"""
        if not self._is_general_tab_active():
            self.go_to_adjacent_memo_date(1)
            return "break"
        if self._is_listbox_focused():
            self.move_memo_down()

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
        # get("1.0", tk.END)는 Tk Text 위젯이 항상 자동으로 붙이는 마지막 개행까지
        # 포함해서 반환하므로, 그 한 글자만 제외하는 "end-1c"를 사용함. 이전에는
        # .strip()으로 앞뒤 공백/빈 줄까지 제거했는데, 사용자가 일부러 넣은 끝줄
        # 공백이나 빈 줄이 저장 시 사라지는 문제가 있어 원본 그대로 저장하도록 변경함
        content = self.content_text.get("1.0", "end-1c")
        self.memos[self.current_index] = {"title": title, "content": content}
        
        if update_list:
            # 전체 delete(0, END)+재삽입 대신, 바뀐 항목 하나만 갱신
            # (다른 항목은 그대로라 스크롤 위치도 자연히 유지됨)
            self.listbox.delete(self.current_index)
            self.listbox.insert(self.current_index, title)
            self.listbox.selection_set(self.current_index)
        
        self._debounced_save("memos", self.save_memos)
        self.update_status_bar()

    def on_drag_start(self, event):
        """드래그 시작: 시작 인덱스 저장"""
        self.drag_start_index = self.listbox.nearest(event.y)
        # 기존 선택 이벤트가 먼저 발생하도록 함
        return

    def on_drag_motion(self, event):
        """드래그 중: 현재 마우스 위치를 활성(active) 표시로만 안내.
        (selection_set() 대신 activate()로 바꾼 것만으로는 부족했음 - Listbox
        위젯 자체의 기본 클래스 바인딩이 <B1-Motion>에서 마우스가 지나가는
        항목을 스스로 선택(selection)하도록 되어 있어서, 우리 쪽 인스턴스
        바인딩과는 별개로 계속 실행되며 <<ListboxSelect>>를 발생시켜 편집창이
        바뀌는 문제가 그대로 남아 있었음. 반드시 "break"를 반환해서 Listbox의
        기본 동작 자체가 실행되지 않도록 막아야 함)"""
        if self.drag_start_index is None:
            return
        
        current_index = self.listbox.nearest(event.y)
        
        # 현재 마우스 위치의 항목을 active 표시로만 안내 (선택 자체는 바꾸지 않음)
        if 0 <= current_index < len(self.memos):
            self.listbox.activate(current_index)
        
        # Listbox의 기본 클래스 바인딩(<B1-Motion>)이 이어서 실행되며 실제
        # 선택(selection)을 옮기는 것을 막기 위해 반드시 break를 반환해야 함
        return "break"

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

    # ------------------------------------------------------------------
    # 달력메모(달력) 탭 관련 메서드
    # ------------------------------------------------------------------

    def build_calendar_tab(self, parent):
        """"달력메모" 탭 UI 구성: 왼쪽 달력([월별보기]/[1년전체]) + 오른쪽 메모 내용(제목 없음)"""
        now = datetime.now()
        self.cal_year, self.cal_month = now.year, now.month
        self.cal_year_year = now.year
        self.selected_date = self._today_str()

        cal_container = ttk.Frame(parent)
        cal_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.calendar_pane = cal_container

        # ---- 왼쪽: 달력 (고정 폭 - 사용자가 크기조절 막대로 바꿀 수 없도록 PanedWindow 대신 사용) ----
        cal_left = ttk.Frame(cal_container, width=CAL_LEFT_WIDTH)
        cal_left.pack(side=tk.LEFT, fill=tk.Y)
        cal_left.pack_propagate(False)

        ttk.Separator(cal_container, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=8)

        self.cal_view_notebook = ttk.Notebook(cal_left)

        # [이전 메모]/[다음 메모] 버튼 - 달력 영역(cal_left) 맨 아래, 상태표시줄 바로 위.
        # side=BOTTOM으로 먼저 배치해야 아래에서 expand=True로 채워지는 노트북이
        # 이 영역을 침범하지 않음 (pack은 호출 순서대로 공간을 배정함)
        cal_nav_buttons = ttk.Frame(cal_left)
        cal_nav_buttons.pack(side=tk.BOTTOM, fill=tk.X, pady=(6, 0))
        self.prev_memo_button = ttk.Button(cal_nav_buttons, text="◀ 이전 메모",
                                            command=lambda: self.go_to_adjacent_memo_date(-1))
        self.prev_memo_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 2))
        self.next_memo_button = ttk.Button(cal_nav_buttons, text="다음 메모 ▶",
                                            command=lambda: self.go_to_adjacent_memo_date(1))
        self.next_memo_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(2, 2))
        # 선택된 날짜에 이미 내용이 있을 때만 활성화 (초기 상태 - 아래 date_content_text가
        # 아직 만들어지기 전이므로 위젯이 아닌 calendar_memos 딕셔너리로 직접 확인)
        initial_has_content = bool(self.calendar_memos.get(self.selected_date, "").strip())
        self.remove_date_memo_button = ttk.Button(
            cal_nav_buttons, text="메모 제거", command=self.remove_date_memo,
            state=(tk.NORMAL if initial_has_content else tk.DISABLED))
        self.remove_date_memo_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(2, 0))

        self.cal_view_notebook.pack(fill=tk.BOTH, expand=True)
        self.cal_view_notebook.bind("<<NotebookTabChanged>>", self._on_cal_view_tab_changed)

        month_tab = ttk.Frame(self.cal_view_notebook)
        year_tab = ttk.Frame(self.cal_view_notebook)
        self.cal_view_notebook.add(month_tab, text="월별보기")
        self.cal_view_notebook.add(year_tab, text="1년전체")

        # -- 월별보기 --
        self.month_year_var = tk.StringVar(value=str(self.cal_year))
        self.month_month_var = tk.StringVar(value=f"{self.cal_month}월")

        # 1행: 연도/월 선택, 2행: 이전/오늘/다음 (한 줄에 다 넣으면 폰트에 따라 "오늘" 버튼이
        # 고정 폭 밖으로 밀려날 수 있어 두 줄로 나눔 - 달력 영역 너비는 그대로 유지)
        # 1행: 연도/월 선택 - 요일(일~토) 폭 안에 가운데 정렬
        # 2행: [◀:일~월][오늘:화~목][▶:금~토] - 요일 레이블 폭에 정확히 맞춰 배치하고
        #      버튼을 더 크게 키워 사용성을 높임 (ttk.Frame을 써서 0열 여백도 테마색이 자동 적용됨)
        m_nav = ttk.Frame(month_tab)
        m_nav.pack(anchor="w", padx=6, pady=(8, 2))
        m_nav.grid_columnconfigure(0, minsize=30)
        for c in range(7):
            m_nav.grid_columnconfigure(c + 1, minsize=CAL_CELL_W + 2)

        nav_controls = ttk.Frame(m_nav)
        nav_controls.grid(row=0, column=1, columnspan=7)
        year_spin = ttk.Spinbox(nav_controls, from_=1900, to=2100, textvariable=self.month_year_var,
                                 width=6, command=self._commit_month_year)
        year_spin.pack(side=tk.LEFT)
        year_spin.bind("<Return>", self._commit_month_year)
        year_spin.bind("<FocusOut>", self._commit_month_year)
        ttk.Label(nav_controls, text="년").pack(side=tk.LEFT, padx=(2, 8))

        month_combo = ttk.Combobox(nav_controls, textvariable=self.month_month_var,
                                    values=[f"{m}월" for m in range(1, 13)],
                                    state="readonly", width=5)
        month_combo.pack(side=tk.LEFT)
        month_combo.bind("<<ComboboxSelected>>", self._on_month_combo_change)

        m_nav2 = ttk.Frame(month_tab)
        m_nav2.pack(anchor="w", padx=6, pady=(0, 6))
        m_nav2.grid_columnconfigure(0, minsize=30)
        for c in range(7):
            m_nav2.grid_columnconfigure(c + 1, minsize=CAL_CELL_W + 2)

        # 기본 ttk 버튼의 좌우 패딩이 넓어서 2칸(84px)에 맞추면 오히려 칸이 밀려나므로,
        # 이 버튼들만 좌우 패딩을 줄인 전용 스타일을 사용해 요일 칸 폭에 정확히 맞춤
        nav_btn_style = ttk.Style()
        nav_btn_style.configure("CalNav.TButton", padding=(2, 4), width=1)

        ttk.Button(m_nav2, text="◀", style="CalNav.TButton",
                   command=lambda: self.go_to_month(self.cal_year, self.cal_month - 1)
                   ).grid(row=0, column=1, columnspan=2, sticky="nsew", padx=1, pady=1, ipady=4)
        ttk.Button(m_nav2, text="오늘", style="CalNav.TButton", command=self.go_to_today
                   ).grid(row=0, column=3, columnspan=3, sticky="nsew", padx=1, pady=1, ipady=4)
        ttk.Button(m_nav2, text="▶", style="CalNav.TButton",
                   command=lambda: self.go_to_month(self.cal_year, self.cal_month + 1)
                   ).grid(row=0, column=6, columnspan=2, sticky="nsew", padx=1, pady=1, ipady=4)

        m_weekday_frame = tk.Frame(month_tab)
        m_weekday_frame.pack(anchor="w", padx=6)
        m_weekday_frame.grid_columnconfigure(0, minsize=30)
        m_corner_lbl = tk.Label(m_weekday_frame, text="", width=3)
        m_corner_lbl.grid(row=0, column=0, sticky="nsew")
        self.weekday_label_widgets.append((m_corner_lbl, "normal"))
        weekday_names = ["일", "월", "화", "수", "목", "금", "토"]
        for i, wd in enumerate(weekday_names):
            m_weekday_frame.grid_columnconfigure(i + 1, minsize=CAL_CELL_W + 2)
            kind = "sun" if i == 0 else ("sat" if i == 6 else "normal")
            lbl = tk.Label(m_weekday_frame, text=wd, width=4, font=("맑은 고딕", 9, "bold"))
            lbl.grid(row=0, column=i + 1, sticky="nsew")
            self.weekday_label_widgets.append((lbl, kind))

        self.month_grid_frame = tk.Frame(month_tab)
        self.month_grid_frame.pack(anchor="w", padx=6, pady=(2, 8))

        # -- 1년전체 --
        self.year_year_var = tk.StringVar(value=str(self.cal_year_year))

        y_nav = ttk.Frame(year_tab)
        y_nav.pack(fill=tk.X, padx=6, pady=(8, 4))
        year_spin_y = ttk.Spinbox(y_nav, from_=1900, to=2100, textvariable=self.year_year_var,
                                   width=6, command=self._commit_year_year)
        year_spin_y.pack(side=tk.LEFT)
        year_spin_y.bind("<Return>", self._commit_year_year)
        year_spin_y.bind("<FocusOut>", self._commit_year_year)
        ttk.Label(y_nav, text="년").pack(side=tk.LEFT, padx=(2, 8))
        ttk.Button(y_nav, text="◀", width=3,
                   command=lambda: self.go_to_year(self.cal_year_year - 1)).pack(side=tk.LEFT, padx=2)
        ttk.Button(y_nav, text="▶", width=3,
                   command=lambda: self.go_to_year(self.cal_year_year + 1)).pack(side=tk.LEFT, padx=2)
        ttk.Button(y_nav, text="올해",
                   command=lambda: self.go_to_year(datetime.now().year)).pack(side=tk.LEFT, padx=(6, 0))

        y_weekday_frame = tk.Frame(year_tab)
        y_weekday_frame.pack(fill=tk.X, padx=6)
        y_weekday_frame.grid_columnconfigure(0, minsize=30)
        corner_lbl = tk.Label(y_weekday_frame, text="", width=3)
        corner_lbl.grid(row=0, column=0, sticky="nsew")
        self.weekday_label_widgets.append((corner_lbl, "normal"))
        for i, wd in enumerate(weekday_names):
            y_weekday_frame.grid_columnconfigure(i + 1, minsize=CAL_CELL_W + 2)
            kind = "sun" if i == 0 else ("sat" if i == 6 else "normal")
            lbl = tk.Label(y_weekday_frame, text=wd, width=4, font=("맑은 고딕", 9, "bold"))
            lbl.grid(row=0, column=i + 1, sticky="nsew")
            self.weekday_label_widgets.append((lbl, kind))

        y_scroll_container = ttk.Frame(year_tab)
        y_scroll_container.pack(fill=tk.BOTH, expand=True, padx=6, pady=(2, 8))
        self.year_canvas = tk.Canvas(y_scroll_container, highlightthickness=0)
        # sv_ttk 테마의 스크롤바는 매우 얇아서 눈에 잘 안 띄기 때문에, 항상 뚜렷하게 보이는
        # 기본 Tk 스크롤바를 사용해 사용자가 현재 위치를 가늠할 수 있게 함
        self.year_scrollbar = tk.Scrollbar(y_scroll_container, orient="vertical",
                                            command=self.year_canvas.yview, width=16)
        self.year_canvas.configure(yscrollcommand=self.year_scrollbar.set)
        self.year_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.year_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.year_grid_frame = tk.Frame(self.year_canvas)
        self.year_canvas.create_window((0, 0), window=self.year_grid_frame, anchor="nw")
        self.year_grid_frame.bind(
            "<Configure>",
            lambda e: self.year_canvas.configure(scrollregion=self.year_canvas.bbox("all"))
        )
        self.year_canvas.bind("<Enter>", lambda e: self.year_canvas.bind_all("<MouseWheel>", self._on_year_mousewheel))
        self.year_canvas.bind("<Leave>", lambda e: self.year_canvas.unbind_all("<MouseWheel>"))

        # ---- 오른쪽: 달력 메모 내용 (제목 필드 없음) ----
        cal_right = ttk.Frame(cal_container)
        cal_right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.date_label_var = tk.StringVar(value=self._format_date_label(self.selected_date))
        ttk.Label(cal_right, textvariable=self.date_label_var, font=self.ui_font).pack(anchor="w")

        # height=1: content_text와 같은 이유(자연 요구 크기 최소화) - 현재는 이 열에
        # 아래로 다른 형제 위젯이 없어 당장 잘리는 문제는 없지만 동일하게 맞춰둠
        self.date_content_text = tk.Text(cal_right, font=self.content_font, padx=10, pady=8, height=1)
        self.date_content_text.pack(fill=tk.BOTH, expand=True)
        self.date_content_text.insert("1.0", self.calendar_memos.get(self.selected_date, ""))
        self.date_content_text.bind("<KeyRelease>", self.save_date_memo_realtime)
        self.date_content_text.bind("<ButtonRelease-1>", self.update_status_bar)
        # PageUp/PageDown은 내용을 편집 중이어도 항상 이전/다음 메모로 이동해야 하므로,
        # Text의 기본 클래스 바인딩(페이지 스크롤)이 실행되기 전에 인스턴스 바인딩에서
        # break로 가로챔 (Listbox의 기본 B1-Motion 동작을 막을 때와 같은 방식)
        self.date_content_text.bind("<Prior>", self._on_prior_key)
        self.date_content_text.bind("<Next>", self._on_next_key)

        date_copy_frame = ttk.Frame(cal_right)
        date_copy_frame.pack(anchor="e", pady=5)
        self.date_copy_status_label = ttk.Label(date_copy_frame, text="", font=("맑은 고딕", 11, "bold"))
        self.date_copy_status_label.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(date_copy_frame, text="클립보드로 복사", command=self.copy_to_clipboard).pack(side=tk.LEFT)

    def _today_str(self):
        return datetime.now().strftime("%Y-%m-%d")

    def _format_date_label(self, date_key):
        try:
            d = datetime.strptime(date_key, "%Y-%m-%d")
            weekdays_kr = ["월", "화", "수", "목", "금", "토", "일"]
            return f"{date_key} ({weekdays_kr[d.weekday()]}) 메모 내용"
        except Exception:
            return f"{date_key} 메모 내용"

    def _is_general_tab_active(self):
        try:
            return self.notebook.index(self.notebook.select()) == 0
        except Exception:
            return True

    def _is_year_view_active(self):
        try:
            return self.cal_view_notebook.index(self.cal_view_notebook.select()) == 1
        except Exception:
            return False

    def on_tab_changed(self, event=None):
        """[일반메모]/[달력메모] 탭 전환 시 상태표시줄을 갱신하고, 달력메모 탭으로
        전환되는 경우 "오늘" 표시가 최신 날짜를 반영하도록 다시 그림(자정 경과 대비)."""
        if not self._is_general_tab_active():
            if self._is_year_view_active():
                self.render_year_view()
                self._year_view_stale = False
            else:
                self.render_month_view()
                self._month_view_stale = False
        self.update_status_bar()

    def _on_cal_view_tab_changed(self, event=None):
        """[월별보기]/[1년전체] 전환 시, 그 사이 다른 날짜를 클릭해 갱신이 미뤄져 있었다면
        (성능을 위해 보이지 않는 뷰는 즉시 다시 그리지 않으므로) 지금 그려준다."""
        if self._is_year_view_active():
            if getattr(self, "_year_view_stale", False):
                self.render_year_view()
                self._year_view_stale = False
        else:
            if getattr(self, "_month_view_stale", False):
                self.render_month_view()
                self._month_view_stale = False

    def _commit_month_year(self, event=None):
        try:
            y = int(self.month_year_var.get())
        except (ValueError, TypeError):
            y = self.cal_year
        self.go_to_month(y, self.cal_month)

    def _commit_year_year(self, event=None):
        try:
            y = int(self.year_year_var.get())
        except (ValueError, TypeError):
            y = self.cal_year_year
        self.go_to_year(y)

    def _on_month_combo_change(self, event=None):
        try:
            m = int(self.month_month_var.get().replace("월", ""))
        except (ValueError, TypeError):
            m = self.cal_month
        self.go_to_month(self.cal_year, m)

    def _on_year_mousewheel(self, event):
        self.year_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def go_to_month(self, year, month):
        while month < 1:
            month += 12
            year -= 1
        while month > 12:
            month -= 12
            year += 1
        self.cal_year, self.cal_month = year, month
        self.render_month_view()

    def go_to_year(self, year):
        self.cal_year_year = year
        self.render_year_view()

    def go_to_today(self):
        """[월별보기]의 "오늘" 버튼: 화면을 오늘이 있는 달로 이동함과 동시에 오늘 날짜를
        선택하여 오른쪽 편집창도 오늘 메모로 전환한다.
        (반면 [1년전체]의 "올해" 버튼은 go_to_year를 그대로 호출해 화면 이동만 하고
        편집 중인 날짜는 바꾸지 않는다.)"""
        self.on_calendar_date_click(self._today_str())

    def on_calendar_date_click(self, date_key):
        """달력의 날짜 칸 클릭: 해당 날짜를 선택하고 오른쪽에 그 날짜의 메모 내용을 표시.
        (다른 달/해의 흐린 날짜를 클릭한 경우 그 달/해로 이동도 함께 처리)"""
        try:
            y, m, _ = date_key.split("-")
            self.cal_year, self.cal_month = int(y), int(m)
            self.cal_year_year = int(y)
        except Exception:
            pass

        self.selected_date = date_key
        self.date_label_var.set(self._format_date_label(date_key))
        self.date_content_text.delete("1.0", tk.END)
        self.date_content_text.insert("1.0", self.calendar_memos.get(date_key, ""))
        self._update_remove_date_memo_button_state()
        self._refresh_calendar_views()
        self.update_status_bar()

    def go_to_adjacent_memo_date(self, direction):
        """[이전 메모]/[다음 메모] 버튼 및 PageUp(-1)/PageDown(+1): 메모가 있는
        날짜 중 현재 선택된 날짜보다 이전/이후인 가장 가까운 날짜로 이동.
        (현재 선택된 날짜 자체에 메모가 없어도, 그 날짜를 기준으로 찾는다)"""
        if direction < 0:
            candidates = [d for d in self.calendar_memos if d < self.selected_date]
            if not candidates:
                messagebox.showinfo("알림", "이전 메모가 없습니다.")
                return
            target = max(candidates)
        else:
            candidates = [d for d in self.calendar_memos if d > self.selected_date]
            if not candidates:
                messagebox.showinfo("알림", "다음 메모가 없습니다.")
                return
            target = min(candidates)
        self.on_calendar_date_click(target)

    def _refresh_calendar_views(self):
        """현재 보이는 달력 뷰만 즉시 다시 그리고, 다른 쪽은 다음에 그 탭으로
        전환될 때 그리도록 표시만 해둔다(불필요한 위젯 재생성을 피해 반응성을 유지)."""
        if self._is_year_view_active():
            self.render_year_view()
            self._month_view_stale = True
        else:
            self.render_month_view()
            self._year_view_stale = True

    def save_date_memo_realtime(self, event=None):
        """달력메모 내용을 실시간으로 memos_calendar.json에 저장.
        (내용을 모두 지우면 해당 날짜 항목 자체를 제거해, 메모 표시 점과 파일을 깔끔하게 유지)"""
        if not self.selected_date:
            return
        had_memo = self.selected_date in self.calendar_memos
        # 빈 날짜 판정(자동 삭제 여부)에는 .strip()으로 공백만 있는지 확인하되,
        # 실제로 저장하는 값은 원본을 그대로 사용해 사용자가 입력한 끝줄 공백/빈
        # 줄이 저장 시 사라지지 않도록 함 ("end-1c"는 Tk가 자동으로 붙이는 마지막
        # 개행 한 글자만 제외함)
        content = self.date_content_text.get("1.0", "end-1c")
        if content.strip():
            self.calendar_memos[self.selected_date] = content
        else:
            self.calendar_memos.pop(self.selected_date, None)
        self._update_remove_date_memo_button_state()
        self._debounced_save("calendar_memos", self.save_calendar_memos)
        # 빈 날짜에 처음 메모를 쓰거나 마지막 내용을 지운 경우, 즉 "메모 있음" 여부
        # 자체가 바뀐 경우에만 달력 칸의 진한 배경색이 바뀌므로, 그 순간에만 debounce로
        # 달력 그리드를 다시 그림(계속 타이핑 중에는 저장과 마찬가지로 미뤄지므로,
        # 매 키 입력마다 칸을 재생성해 버벅이는 일 없이 입력이 멈췄을 때 한 번만 반영됨)
        if had_memo != bool(content.strip()):
            self._debounced_save("calendar_view_refresh", self._refresh_calendar_views)
        self.update_status_bar()

    def _update_remove_date_memo_button_state(self):
        """[메모 제거] 버튼을 현재 date_content_text 내용이 있을 때만 활성화.
        (Ctrl+D 단축키도 이 버튼 상태를 그대로 확인해서 동작 여부를 결정함)"""
        has_content = bool(self.date_content_text.get("1.0", tk.END).strip())
        self.remove_date_memo_button.config(state=(tk.NORMAL if has_content else tk.DISABLED))

    def remove_date_memo(self, event=None):
        """[메모 제거] 버튼 및 Ctrl+D: 선택된 날짜의 메모 내용을 확인 후 삭제.
        버튼이 비활성 상태(이미 내용 없음)면 아무 것도 하지 않음."""
        if str(self.remove_date_memo_button.cget("state")) != tk.NORMAL:
            return "break"
        if messagebox.askyesno("확인", f"{self.selected_date} 메모 내용을 삭제하시겠습니까?"):
            self.calendar_memos.pop(self.selected_date, None)
            self.save_calendar_memos()
            self.on_calendar_date_click(self.selected_date)
        return "break"

    def render_month_view(self):
        """[월별보기] 그리드를 현재 self.cal_year/self.cal_month 기준으로 다시 그림
        (1년전체 보기와 폭을 맞추기 위해 왼쪽에 빈 칸(0열)을 두고 요일 칸은 1~7열에 그림)"""
        self._hide_holiday_tooltip()  # 재구성 전, 떠 있을 수 있는 풍선말 정리
        for widget in self.month_grid_frame.winfo_children():
            widget.destroy()
        self.month_grid_frame.grid_columnconfigure(0, minsize=30)
        for c in range(1, 8):
            self.month_grid_frame.grid_columnconfigure(c, minsize=CAL_CELL_W + 2)

        year, month = self.cal_year, self.cal_month
        self.month_year_var.set(str(year))
        self.month_month_var.set(f"{month}월")

        first_weekday_mon0, total_days = calendar.monthrange(year, month)
        first_weekday = (first_weekday_mon0 + 1) % 7  # 월요일=0 -> 일요일=0 기준으로 변환

        prev_month = 12 if month == 1 else month - 1
        prev_year = year - 1 if month == 1 else year
        prev_days = calendar.monthrange(prev_year, prev_month)[1]

        next_month = 1 if month == 12 else month + 1
        next_year = year + 1 if month == 12 else year

        cells = []
        for i in range(first_weekday - 1, -1, -1):
            d = prev_days - i
            cells.append((d, f"{prev_year:04d}-{prev_month:02d}-{d:02d}", True))
        for d in range(1, total_days + 1):
            cells.append((d, f"{year:04d}-{month:02d}-{d:02d}", False))
        remainder = len(cells) % 7
        trailing = 0 if remainder == 0 else 7 - remainder
        for d in range(1, trailing + 1):
            cells.append((d, f"{next_year:04d}-{next_month:02d}-{d:02d}", True))

        for idx, (day, key, other) in enumerate(cells):
            row, col = divmod(idx, 7)
            if other:
                kind = "other"
            elif col == 0:
                kind = "sun"
            elif col == 6:
                kind = "sat"
            else:
                kind = "normal"
            self._make_day_cell(self.month_grid_frame, row, col + 1, day, kind, key)

    def render_year_view(self):
        """[1년전체] 그리드를 현재 self.cal_year_year 기준으로 다시 그림
        (1월 1일이 속한 주의 일요일부터 12월 31일이 속한 주의 토요일까지 이어서 표시)"""
        self._hide_holiday_tooltip()  # 재구성 전, 떠 있을 수 있는 풍선말 정리
        for widget in self.year_grid_frame.winfo_children():
            widget.destroy()
        self.year_grid_frame.grid_columnconfigure(0, minsize=30)
        for c in range(1, 8):
            self.year_grid_frame.grid_columnconfigure(c, minsize=CAL_CELL_W + 2)

        year = self.cal_year_year
        self.year_year_var.set(str(year))
        colors = THEME_COLORS.get(self.theme_mode, THEME_COLORS["light"])

        jan1 = date(year, 1, 1)
        start_pad = (jan1.weekday() + 1) % 7
        cursor = jan1 - timedelta(days=start_pad)

        dec31 = date(year, 12, 31)
        end_pad = 6 - ((dec31.weekday() + 1) % 7)
        end_date = dec31 + timedelta(days=end_pad)

        total_rows = ((end_date - cursor).days + 1) // 7

        d = cursor
        for r in range(total_rows):
            row_dates = [d + timedelta(days=c) for c in range(7)]
            d = d + timedelta(days=7)

            month_start = next((dt for dt in row_dates if dt.year == year and dt.day == 1), None)
            if month_start:
                label_cell = tk.Label(self.year_grid_frame, text=str(month_start.month), width=3,
                                       bg=colors["fg"], fg=colors["bg"], font=("맑은 고딕", 9, "bold"))
            else:
                label_cell = tk.Label(self.year_grid_frame, text="", width=3, bg=colors["bg"])
            label_cell.grid(row=r, column=0, sticky="nsew", padx=1, pady=1)

            for col, dt in enumerate(row_dates):
                in_year = dt.year == year
                key = dt.strftime("%Y-%m-%d")
                if not in_year:
                    kind = "other"
                elif col == 0:
                    kind = "sun"
                elif col == 6:
                    kind = "sat"
                else:
                    kind = "normal"
                self._make_day_cell(self.year_grid_frame, r, col + 1, dt.day, kind, key)

    def _make_day_cell(self, parent, row, col, day_num, kind, date_key):
        """달력의 날짜 한 칸을 그림(오늘=칠해진 원, 선택된 날짜=테두리 원, 메모 있음=진한 배경색)"""
        colors = THEME_COLORS.get(self.theme_mode, THEME_COLORS["light"])
        has_memo = bool(str(self.calendar_memos.get(date_key, "")).strip())
        cell_bg = colors["has_memo_bg"] if has_memo else colors["bg"]
        canvas = tk.Canvas(parent, width=CAL_CELL_W, height=CAL_CELL_H,
                            highlightthickness=0, bg=cell_bg, cursor="hand2")
        canvas.grid(row=row, column=col, sticky="nsew", padx=1, pady=1)

        holiday_name = self.holidays.get(date_key)
        # 흐리게 표시되는 다른 달 날짜("other")는 일요일/토요일도 특별 색을 안 쓰는
        # 기존 규칙과 똑같이, 공휴일이어도 색을 강조하지 않음(시각적 일관성 유지)
        if kind == "other":
            fg = colors["muted_fg"]
            holiday_name = None
        elif holiday_name:
            fg = colors["sunday_fg"]  # 공휴일은 토요일이어도 일요일과 같은 빨간색
        elif kind == "sun":
            fg = colors["sunday_fg"]
        elif kind == "sat":
            fg = colors["accent"]
        else:
            fg = colors["fg"]

        cx, cy = CAL_CELL_W / 2, CAL_CELL_H / 2 - 2
        is_today = (date_key == self._today_str())
        is_selected = (date_key == self.selected_date)

        if is_today:
            r = 12
            canvas.create_oval(cx - r, cy - r, cx + r, cy + r, fill=colors["list_select_bg"], outline="")
            fg = colors["list_select_fg"]
        if is_selected:
            r = 14
            canvas.create_oval(cx - r, cy - r, cx + r, cy + r, outline=colors["accent"], width=2)

        canvas.create_text(cx, cy, text=str(day_num), fill=fg, font=("맑은 고딕", 10))

        canvas.bind("<Button-1>", lambda e, dk=date_key: self.on_calendar_date_click(dk))
        if holiday_name:
            canvas.bind("<Enter>", lambda e, name=holiday_name: self._show_holiday_tooltip(e, name))
            canvas.bind("<Leave>", self._hide_holiday_tooltip)
        return canvas

    def _show_holiday_tooltip(self, event, text):
        """공휴일 날짜 칸에 마우스를 올리면 공휴일 이름을 작은 풍선말로 표시.
        (공휴일 정보는 달력메모 내용이 아니라 이렇게 별도 풍선말로만 안내함)"""
        self._hide_holiday_tooltip()
        tooltip = tk.Toplevel(self.root)
        tooltip.wm_overrideredirect(True)
        try:
            tooltip.wm_attributes("-topmost", True)
        except tk.TclError:
            pass
        tooltip.wm_geometry(f"+{event.x_root + 12}+{event.y_root + 12}")
        tk.Label(tooltip, text=text, background="#ffffe0", foreground="#000000",
                 relief=tk.SOLID, borderwidth=1, font=("맑은 고딕", 9),
                 padx=6, pady=3).pack()
        self._holiday_tooltip = tooltip

    def _hide_holiday_tooltip(self, event=None):
        """떠 있는 공휴일 풍선말이 있으면 닫음"""
        if self._holiday_tooltip is not None:
            self._holiday_tooltip.destroy()
            self._holiday_tooltip = None

    def on_closing(self):
        # 디바운스로 미뤄둔 저장이 있다면 취소하고(중복 저장 방지),
        # 아래에서 최신 상태를 바로 저장하므로 유실되는 내용은 없음
        self._cancel_pending_saves()
        # 종료 전 설정 및 메모 저장 (창 크기, 위치 포함)
        memos_ok = self.save_memos()
        calendar_ok = self.save_calendar_memos()
        self.save_settings()
        if not (memos_ok and calendar_ok):
            # 디스크 공간 부족/권한 문제 등으로 마지막 저장이 실패했는데도 그냥
            # 종료해버리면 방금 작성한 내용이 그대로 사라지므로, 사용자에게 알리고
            # 종료 여부를 직접 선택하게 함
            if not messagebox.askyesno(
                "저장 실패",
                "메모 저장에 실패했습니다. 지금 종료하면 마지막으로 입력한 내용이\n"
                "저장되지 않을 수 있습니다. 그래도 종료하시겠습니까?",
                parent=self.root,
            ):
                return
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = MemoApp(root)
    root.mainloop()