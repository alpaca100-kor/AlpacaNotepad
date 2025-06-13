import tkinter as tk
from tkinter import messagebox, PanedWindow, filedialog, Toplevel, font, ttk
import json
import os

# 엑셀 파일 처리를 위한 라이브러리. 설치 필요 (pip install openpyxl)
try:
    import openpyxl
except ImportError:
    openpyxl = None

class MemoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("알파카 메모장")
        self.file_path = "memos.json"
        self.settings_file = "settings.json"
        self.memos = self.load_memos()
        self.settings = self.load_settings()
        self.current_index = -1

        self.root.geometry("800x600")
        
        # UI 기본 글꼴
        self.ui_font = ("굴림체", 12)
        # 메모 내용 글꼴 (설정에서 변경 가능)
        self.content_font = (self.settings.get("font_family", "굴림체"), self.settings.get("font_size", 12))

        # --- 메뉴바 생성 ---
        self.create_menu()

        main_pane = PanedWindow(root, sashrelief=tk.RAISED, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 왼쪽 패널
        left_panel = tk.Frame(main_pane)
        main_pane.add(left_panel, width=250)
        main_pane.paneconfig(left_panel, minsize=200)

        list_frame = tk.Frame(left_panel)
        list_frame.pack(fill=tk.BOTH, expand=True)

        self.listbox = tk.Listbox(list_frame, exportselection=False, font=self.ui_font)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.listbox.bind("<<ListboxSelect>>", self.on_memo_select)
        self.listbox.bind("<Delete>", lambda event: self.remove_memo())

        scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.config(yscrollcommand=scrollbar.set)

        self.update_listbox()

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

        # 오른쪽 패널
        right_panel = tk.Frame(main_pane)
        main_pane.add(right_panel)

        title_label = tk.Label(right_panel, text="메모 제목", font=self.ui_font)
        title_label.pack(anchor="w")

        self.title_entry = tk.Entry(right_panel, font=self.ui_font)
        self.title_entry.pack(fill=tk.X, pady=(0, 10))
        self.title_entry.bind("<KeyRelease>", self.update_memo_realtime)

        content_label = tk.Label(right_panel, text="메모 내용", font=self.ui_font)
        content_label.pack(anchor="w")

        self.content_text = tk.Text(right_panel, font=self.content_font)
        self.content_text.pack(fill=tk.BOTH, expand=True)
        self.content_text.bind("<KeyRelease>", self.update_memo_realtime)
        
        self.root.bind("<Control-n>", lambda event: self.add_memo())
        self.root.bind("<Prior>", lambda event: self.move_memo_up())
        self.root.bind("<Next>", lambda event: self.move_memo_down())

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.toggle_right_panel(False)

    def create_menu(self):
        menubar = tk.Menu(self.root)
        
        # 파일 메뉴
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="메모 가져오기...", command=self.import_memos)
        file_menu.add_command(label="메모 내보내기...", command=self.export_memos)
        file_menu.add_separator()
        file_menu.add_command(label="종료", command=self.root.quit)
        menubar.add_cascade(label="파일", menu=file_menu)

        # 설정 메뉴
        settings_menu = tk.Menu(menubar, tearoff=0)
        settings_menu.add_command(label="글꼴 설정...", command=self.open_font_settings)
        menubar.add_cascade(label="설정", menu=settings_menu)

        self.root.config(menu=menubar)

    def load_settings(self):
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, "r", encoding="utf-8") as f:
                    return json.load(f)
            except (json.JSONDecodeError, IOError):
                return {}
        return {}

    def save_settings(self):
        with open(self.settings_file, "w", encoding="utf-8") as f:
            json.dump(self.settings, f, ensure_ascii=False, indent=4)

    def open_font_settings(self):
        settings_win = Toplevel(self.root)
        settings_win.title("글꼴 설정")
        settings_win.geometry("350x150")
        settings_win.resizable(False, False)
        settings_win.grab_set() # 모달 창으로 만들기

        # 글꼴 선택
        tk.Label(settings_win, text="글꼴:", font=self.ui_font).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        font_families = sorted(font.families())
        font_var = tk.StringVar(value=self.content_font[0])
        font_combo = ttk.Combobox(settings_win, textvariable=font_var, values=font_families, state="readonly")
        font_combo.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        # 글자 크기 선택
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
                new_memos = json.load(f)
            
            if isinstance(new_memos, list) and all("title" in m and "content" in m for m in new_memos):
                if messagebox.askyesno("확인", "기존 메모를 덮어쓰고 가져오시겠습니까?"):
                    self.memos = new_memos
                    self.save_memos()
                    self.current_index = -1
                    self.title_entry.delete(0, tk.END)
                    self.content_text.delete("1.0", tk.END)
                    self.toggle_right_panel(False)
                    self.update_listbox()
                    messagebox.showinfo("성공", "메모를 성공적으로 가져왔습니다.")
            else:
                 messagebox.showerror("오류", "올바른 메모 파일 형식이 아닙니다.")
        except Exception as e:
            messagebox.showerror("오류", f"파일을 가져오는 중 오류가 발생했습니다:\n{e}")

    def export_memos(self):
        filepath = filedialog.asksaveasfilename(
            title="메모 내보내기",
            defaultextension=".json",
            filetypes=[
                ("JSON 파일", "*.json"),
                ("텍스트 파일", "*.txt"),
                ("Excel 파일", "*.xlsx")
            ]
        )
        if not filepath:
            return

        file_ext = os.path.splitext(filepath)[1].lower()

        try:
            if file_ext == ".json":
                self.save_memos() # 현재 메모를 memos.json에 저장 후
                with open(self.file_path, 'r', encoding='utf-8') as f_in, open(filepath, 'w', encoding='utf-8') as f_out:
                    f_out.write(f_in.read())
            elif file_ext == ".txt":
                with open(filepath, "w", encoding="utf-8") as f:
                    for memo in self.memos:
                        f.write(f"제목: {memo['title']}\n")
                        f.write("-" * 20 + "\n")
                        f.write(f"{memo['content']}\n\n")
                        f.write("=" * 20 + "\n\n")
            elif file_ext == ".xlsx":
                if not openpyxl:
                    messagebox.showerror("오류", "Excel 파일로 내보내려면 'openpyxl' 라이브러리가 필요합니다.\n\n(터미널에서 'pip install openpyxl' 실행)")
                    return
                
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "메모"
                ws.append(["제목", "내용"])
                for memo in self.memos:
                    ws.append([memo["title"], memo["content"]])
                wb.save(filepath)
            
            messagebox.showinfo("성공", f"메모를 {filepath} 파일로 성공적으로 내보냈습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"파일을 내보내는 중 오류가 발생했습니다:\n{e}")

    def toggle_right_panel(self, enabled):
        state = tk.NORMAL if enabled else tk.DISABLED
        bg_color = "white" if enabled else "#f0f0f0"
        self.title_entry.config(state=state, bg=bg_color)
        self.content_text.config(state=state, bg=bg_color)

    def load_memos(self):
        if os.path.exists(self.file_path):
            try:
                with open(self.file_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except (json.JSONDecodeError, IOError):
                 return []
        else:
            return []

    def save_memos(self):
        with open(self.file_path, "w", encoding="utf-8") as f:
            json.dump(self.memos, f, ensure_ascii=False, indent=4)

    def update_listbox(self):
        self.listbox.delete(0, tk.END)
        for memo in self.memos:
            self.listbox.insert(tk.END, memo["title"])

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
            self.update_listbox()
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(self.current_index)
            self.listbox.activate(self.current_index)
            self.save_memos()

    def move_memo_down(self):
        if 0 <= self.current_index < len(self.memos) - 1:
            self.memos.insert(self.current_index + 1, self.memos.pop(self.current_index))
            self.current_index += 1
            self.update_listbox()
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(self.current_index)
            self.listbox.activate(self.current_index)
            self.save_memos()

    def update_memo_realtime(self, event):
        if self.current_index == -1 or self.title_entry.cget('state') == tk.DISABLED:
            return
        title = self.title_entry.get()
        content = self.content_text.get("1.0", tk.END).strip()
        self.memos[self.current_index] = {"title": title, "content": content}
        self.listbox.delete(self.current_index)
        self.listbox.insert(self.current_index, title)
        self.listbox.selection_set(self.current_index)
        self.save_memos()

    def on_closing(self):
        self.save_memos()
        self.save_settings()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = MemoApp(root)
    root.mainloop()
