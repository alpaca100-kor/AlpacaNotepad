import tkinter as tk
from tkinter import messagebox, PanedWindow, filedialog, Toplevel, font, ttk
import json
import os
import sys
import configparser

# 엑셀 파일 처리를 위한 라이브러리. (없으면 엑셀 내보내기 비활성화)
try:
    import openpyxl
except ImportError:
    openpyxl = None


def resource_path(relative_path):
    """PyInstaller로 생성된 exe의 리소스 경로를 가져옴"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


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

        self.file_path = "memos.json"
        self.settings_file = "settings.ini"
        self.memos = self.load_memos()
        self.settings = self.load_settings()
        self.current_index = -1

        self.ui_font = ("굴림체", 12)
        self.content_font = (self.settings.get("font_family"), self.settings.get("font_size"))
        self.copy_shortcut = self.settings.get("copy_shortcut")

        self.create_menu()

        # 상태표시줄 생성
        self.status_bar = tk.Label(root, text="", bd=1, relief=tk.SUNKEN, anchor=tk.W, font=("굴림체", 10))
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        main_pane = PanedWindow(root, sashrelief=tk.RAISED, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 왼쪽 패널: 메모 리스트
        left_panel = tk.Frame(main_pane)
        main_pane.add(left_panel, width=250)
        main_pane.paneconfig(left_panel, minsize=200)

        list_frame = tk.Frame(left_panel)
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

        scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.config(yscrollcommand=scrollbar.set)

        self.update_listbox()

        # 버튼 프레임
        button_frame = tk.Frame(left_panel)
        button_frame.pack(fill=tk.X, pady=5)

        add_button = tk.Button(button_frame, text="추가", command=self.add_memo)
        add_button.pack(side=tk.LEFT, expand=True, fill=tk.X)
        remove_button = tk.Button(button_frame, text="제거", command=self.remove_memo)
        remove_button.pack(side=tk.LEFT, expand=True, fill=tk.X)
        up_button = tk.Button(button_frame, text="▲", command=self.move_memo_up)
        up_button.pack(side=tk.LEFT, expand=True, fill=tk.X)
        down_button = tk.Button(button_frame, text="▼", command=self.move_memo_down)
        down_button.pack(side=tk.LEFT, expand=True, fill=tk.X)

        # 오른쪽 패널: 제목 + 내용
        right_panel = tk.Frame(main_pane)
        main_pane.add(right_panel)

        title_label = tk.Label(right_panel, text="메모 제목", font=self.ui_font)
        title_label.pack(anchor="w")
        self.title_entry = tk.Entry(right_panel, font=self.ui_font)
        self.title_entry.pack(fill=tk.X, pady=(0, 10))
        self.title_entry.bind("<KeyRelease>", self.update_memo_realtime)
        self.title_entry.bind("<Control-t>", self.focus_on_title)

        content_label = tk.Label(right_panel, text="메모 내용", font=self.ui_font)
        content_label.pack(anchor="w")
        self.content_text = tk.Text(right_panel, font=self.content_font)
        self.content_text.pack(fill=tk.BOTH, expand=True)
        self.content_text.bind("<KeyRelease>", self.update_memo_realtime)
        self.content_text.bind("<ButtonRelease-1>", self.update_status_bar)
        self.content_text.bind("<Control-t>", self.focus_on_title)

        # 복사 버튼과 복사 완료 메시지를 위한 프레임
        copy_frame = tk.Frame(right_panel)
        copy_frame.pack(anchor="e", pady=5)
        
        self.copy_status_label = tk.Label(copy_frame, text="", font=("굴림체", 11, "bold"), fg="blue")
        self.copy_status_label.pack(side=tk.LEFT, padx=(0, 10))
        
        self.copy_button = tk.Button(copy_frame, text="클립보드로 복사", command=self.copy_to_clipboard, font=self.ui_font)
        self.copy_button.pack(side=tk.LEFT)

        # 단축키 등록
        self.root.bind("<Control-n>", lambda event: self.add_memo())
        self.root.bind("<Prior>", lambda event: self.move_memo_up())
        self.root.bind("<Next>", lambda event: self.move_memo_down())
        self.root.bind("<Control-l>", self.focus_on_listbox)
        self.root.bind("<Control-t>", self.focus_on_title)
        # 복사 단축키는 설정에 따라 바인딩
        self.bind_copy_shortcut()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.toggle_right_panel(False)
        self.update_status_bar()
        
        # UI 생성 완료 후 창 위치/크기 복원
        self.restore_window_geometry()

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

    def create_menu(self):
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="메모 가져오기...", command=self.import_memos)
        file_menu.add_command(label="메모 내보내기...", command=self.export_memos)
        file_menu.add_separator()
        file_menu.add_command(label="종료", command=self.root.quit)
        menubar.add_cascade(label="파일", menu=file_menu)

        settings_menu = tk.Menu(menubar, tearoff=0)
        settings_menu.add_command(label="글꼴 설정...", command=self.open_font_settings)
        settings_menu.add_command(label="단축키 설정...", command=self.open_shortcut_settings)
        menubar.add_cascade(label="설정", menu=settings_menu)

        self.root.config(menu=menubar)

    def focus_on_listbox(self, event=None):
        self.listbox.focus_set()
        if self.current_index != -1:
            self.listbox.selection_set(self.current_index)
            self.listbox.activate(self.current_index)
        return "break"

    def focus_on_title(self, event=None):
        if self.title_entry.cget('state') == tk.NORMAL:
            self.title_entry.focus_set()
            self.title_entry.select_range(0, tk.END)
        return "break"

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
        """복사 완료 메시지를 1초 동안 표시"""
        self.copy_status_label.config(text="복사 완료!")
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
            'font_family': '굴림체',
            'font_size': 12,
            'window_geometry': '800x600+100+100',
            'copy_shortcut': 'Ctrl+Shift+C'
        }

        if not os.path.exists(self.settings_file):
            return default_settings

        try:
            config.read(self.settings_file, encoding='utf-8')
            font_family = config.get('Font', 'family', fallback=default_settings['font_family'])
            font_size = config.getint('Font', 'size', fallback=default_settings['font_size'])
            window_geometry = config.get('Window', 'geometry', fallback=default_settings['window_geometry'])
            copy_shortcut = config.get('Shortcuts', 'copy', fallback=default_settings['copy_shortcut'])
            return {
                'font_family': font_family,
                'font_size': font_size,
                'window_geometry': window_geometry,
                'copy_shortcut': copy_shortcut
            }
        except (configparser.Error, ValueError):
            return default_settings

    def save_settings(self):
        config = configparser.ConfigParser()
        config['Font'] = {
            'family': self.settings.get('font_family', '굴림체'),
            'size': str(self.settings.get('font_size', 12))
        }
        config['Window'] = {
            'geometry': self.root.geometry()
        }
        config['Shortcuts'] = {
            'copy': self.settings.get('copy_shortcut', 'Ctrl+Shift+C')
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
        settings_win.grab_set()

        tk.Label(settings_win, text="글꼴:", font=self.ui_font).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        font_families = sorted(font.families())
        font_var = tk.StringVar(value=self.content_font[0])
        font_combo = ttk.Combobox(settings_win, textvariable=font_var, values=font_families, state="readonly")
        font_combo.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        tk.Label(settings_win, text="크기:", font=self.ui_font).grid(row=1, column=0, padx=10, pady=10, sticky="w")
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

        def ok_action():
            apply_and_save()
            settings_win.destroy()

        button_frame = tk.Frame(settings_win)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)
        ok_button = tk.Button(button_frame, text="확인", command=ok_action, width=10)
        ok_button.pack(side=tk.LEFT, padx=5)
        apply_button = tk.Button(button_frame, text="적용", command=apply_and_save, width=10)
        apply_button.pack(side=tk.LEFT, padx=5)

    def open_shortcut_settings(self):
        settings_win = Toplevel(self.root)
        settings_win.title("단축키 설정")
        settings_win.geometry("350x150")
        settings_win.resizable(False, False)
        settings_win.grab_set()

        tk.Label(settings_win, text="클립보드 복사:", font=self.ui_font).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        shortcut_options = ["Ctrl+Shift+C", "Ctrl+Alt+C", "Alt+Shift+C"]
        shortcut_var = tk.StringVar(value=self.copy_shortcut)
        shortcut_combo = ttk.Combobox(settings_win, textvariable=shortcut_var, values=shortcut_options, state="readonly", width=20)
        shortcut_combo.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

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

        def ok_action():
            apply_and_save()
            settings_win.destroy()

        button_frame = tk.Frame(settings_win)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)
        ok_button = tk.Button(button_frame, text="확인", command=ok_action, width=10)
        ok_button.pack(side=tk.LEFT, padx=5)
        apply_button = tk.Button(button_frame, text="적용", command=apply_and_save, width=10)
        apply_button.pack(side=tk.LEFT, padx=5)

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
        bg_color = "white" if enabled else "#f0f0f0"
        self.title_entry.config(state=state, bg=bg_color)
        self.content_text.config(state=state, bg=bg_color)
        self.copy_button.config(state=state)
        self.update_status_bar()

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

    def update_memo_realtime(self, event=None):
        """실시간으로 메모를 업데이트하면서 스크롤 위치 유지"""
        if self.current_index == -1:
            return
        
        title = self.title_entry.get()
        content = self.content_text.get("1.0", tk.END).strip()
        self.memos[self.current_index] = {"title": title, "content": content}
        
        # ✅ preserve_scroll=True로 스크롤 위치 유지
        self.update_listbox(preserve_scroll=True)
        
        # 현재 선택 유지
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