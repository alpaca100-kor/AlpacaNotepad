import tkinter as tk
from tkinter import messagebox, PanedWindow, filedialog, Toplevel, font, ttk
import json
import os
import configparser

# 엑셀 파일 처리를 위한 라이브러리. (없으면 엑셀 내보내기 비활성화)
try:
    import openpyxl
except ImportError:
    openpyxl = None


class MemoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("알파카 메모장")

        self.file_path = "memos.json"
        self.settings_file = "settings.ini"
        self.memos = self.load_memos()
        self.settings = self.load_settings()
        self.current_index = -1

        # ✅ 창 크기 및 위치 복원 (기본값 800x600+100+100)
        geometry = self.settings.get("window_geometry", "800x600+100+100")
        self.root.geometry(geometry)

        self.ui_font = ("굴림체", 12)
        self.content_font = (self.settings.get("font_family"), self.settings.get("font_size"))

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

        self.copy_button = tk.Button(right_panel, text="클립보드로 복사", command=self.copy_to_clipboard, font=self.ui_font)
        self.copy_button.pack(anchor="e", pady=5)

        # 단축키 등록
        self.root.bind("<Control-n>", lambda event: self.add_memo())
        self.root.bind("<Prior>", lambda event: self.move_memo_up())
        self.root.bind("<Next>", lambda event: self.move_memo_down())
        self.root.bind("<Control-l>", self.focus_on_listbox)
        self.root.bind("<Control-t>", self.focus_on_title)
        self.root.bind("<Control-Shift-C>", self.copy_to_clipboard)

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.toggle_right_panel(False)
        self.update_status_bar()

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
            messagebox.showinfo("성공", "메모 내용이 클립보드에 복사되었습니다.", parent=self.root)
        except tk.TclError:
            messagebox.showerror("오류", "클립보드에 접근할 수 없습니다.", parent=self.root)
        except Exception as e:
            messagebox.showerror("오류", f"알 수 없는 오류 발생: {e}", parent=self.root)
        return "break"

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
            'window_geometry': '800x600+100+100'
        }

        if not os.path.exists(self.settings_file):
            return default_settings

        try:
            config.read(self.settings_file, encoding='utf-8')
            font_family = config.get('Font', 'family', fallback=default_settings['font_family'])
            font_size = config.getint('Font', 'size', fallback=default_settings['font_size'])
            window_geometry = config.get('Window', 'geometry', fallback=default_settings['window_geometry'])
            return {
                'font_family': font_family,
                'font_size': font_size,
                'window_geometry': window_geometry
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
        with open(self.settings_file, 'w', encoding='utf-8') as configfile:
            config.write(configfile)

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

    def update_listbox(self):
        self.listbox.delete(0, tk.END)
        for memo in self.memos:
            self.listbox.insert(tk.END, memo["title"])
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
        self.update_listbox()
        self.listbox.selection_set(self.current_index)
        self.listbox.activate(self.current_index)
        self.listbox.see(self.current_index)
        self.on_memo_select(None)
        self.save_memos()

    def update_memo_realtime(self, event=None):
        if self.current_index == -1:
            return
        title = self.title_entry.get()
        content = self.content_text.get("1.0", tk.END).strip()
        self.memos[self.current_index] = {"title": title, "content": content}
        self.update_listbox()
        self.listbox.selection_set(self.current_index)
        self.save_memos()
        self.update_status_bar()

    def on_closing(self):
        # 종료 전 설정 및 메모 저장 (창 크기, 위치 포함)
        self.save_memos()
        self.save_settings()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = MemoApp(root)
    root.mainloop()
