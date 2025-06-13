import tkinter as tk
from tkinter import messagebox, ttk
from tkcalendar import Calendar
import webbrowser
import os
import sys
import subprocess
from docx import Document
from datetime import datetime, date
import pandas as pd

# ——— 전역 설정 ———
if getattr(sys, 'frozen', False):
    base_dir = sys._MEIPASS
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

def load_student_data():
    """학생 명단을 불러와 DataFrame과 이름 리스트를 반환합니다."""
    try:
        df = pd.read_excel(
            os.path.join(base_dir, "student_list", "student_list.xlsx"),
            dtype=str
        )
        return df, df["학생 이름"].tolist()
    except Exception as e:
        tk.Tk().withdraw()
        messagebox.showerror("오류", f"학생 명단을 불러올 수 없습니다:\n{e}")
        return pd.DataFrame(), []

# 데이터 및 출력 디렉토리 초기화
df_students, student_names = load_student_data()
today = datetime.today()
selected_end_year = None
if getattr(sys, 'frozen', False):
    exec_dir = os.path.dirname(sys.executable)
else:
    exec_dir = os.path.dirname(os.path.abspath(__file__))
output_dir = os.path.join(exec_dir, "output")
os.makedirs(output_dir, exist_ok=True)

# ─── 기능 정의 ───

def calculate_days():
    """시작·종료 날짜로부터 일수를 계산해 days_var에 설정합니다."""
    try:
        y, m, d = int(start_year.get()), int(start_month.get()), int(start_day.get())
        em, ed = int(end_month.get()), int(end_day.get())
        ey = selected_end_year or y
        diff = (date(ey, em, ed) - date(y, m, d)).days + 1
        if diff < 1:
            messagebox.showwarning("날짜 오류", "종료 날짜가 시작 날짜보다 이전입니다.")
            days_var.set("")
        else:
            days_var.set(str(diff))
    except ValueError:
        days_var.set("")

def open_calendar(which):
    """guide_frame 내 고정된 calendar_frame에 달력을 표시합니다."""
    for w in calendar_frame.winfo_children():
        w.destroy()
    calendar_frame.grid()
    cal = Calendar(calendar_frame,
                   selectmode="day",
                   year=today.year, month=today.month, day=today.day)
    cal.pack(padx=5, pady=5, fill='both', expand=True)
    def on_select(event):
        sel = cal.selection_get()
        global selected_end_year
        if which == 'start':
            start_year.set(sel.year)
            start_month.set(sel.month)
            start_day.set(sel.day)
        else:
            selected_end_year = sel.year
            end_month.set(sel.month)
            end_day.set(sel.day)
        calculate_days()
    cal.bind("<<CalendarSelected>>", on_select)

def on_student_selected(event=None):
    """콤보박스에서 학생 선택 시 나머지 필드를 자동 채웁니다."""
    name = student_var.get()
    if name in student_names:
        rec = df_students[df_students["학생 이름"] == name].iloc[0]
        grade_var.set(rec["학년"])
        class_var.set(rec["반"])
        number_var.set(rec["번호"])
        name_var.set(rec["학생 이름"])
        parent_var.set(rec["보호자 이름"])

def refresh_file_list():
    """output 폴더의 최신 파일 목록을 갱신합니다."""
    list_files.delete(0, tk.END)
    try:
        files = sorted(
            os.listdir(output_dir),
            key=lambda f: os.path.getmtime(os.path.join(output_dir, f)),
            reverse=True
        )
        for f in files:
            list_files.insert(tk.END, f)
    except Exception:
        pass

def open_file(event=None):
    """리스트에서 더블클릭한 파일을 기본 앱으로 실행합니다."""
    sel = list_files.curselection()
    if not sel:
        return
    path = os.path.join(output_dir, list_files.get(sel[0]))
    try:
        if sys.platform.startswith('win'):
            os.startfile(path)
        elif sys.platform.startswith('darwin'):
            subprocess.call(['open', path])
        else:
            subprocess.call(['xdg-open', path])
    except Exception as e:
        messagebox.showerror("열기 오류", f"파일을 열 수 없습니다:\n{e}")

def generate_document():
    """입력된 데이터를 DOCX 템플릿에 치환하여 저장합니다."""
    data = {
        "학년": grade_var.get(),
        "반": class_var.get(),
        "번호": number_var.get(),
        "이름": name_var.get(),
        "보호자": parent_var.get(),
        "구분": type_var.get(),
        "시작년": start_year.get(),
        "시작월": start_month.get(),
        "시작일": start_day.get(),
        "종료월": end_month.get(),
        "종료일": end_day.get(),
        "며칠간": days_var.get(),
        "사유": reason_text.get("1.0", "end").strip(),
        "오늘날짜": today.strftime("%Y년 %m월 %d일")
    }
    tpl = os.path.join(base_dir, "extract_template", "template_word.docx")
    filename = f"결석신고서_{data['이름']}_{int(data['시작년']):04d}-{int(data['시작월']):02d}-{int(data['시작일']):02d}.docx"
    out_path = os.path.join(output_dir, filename)
    try:
        doc = Document(tpl)
    except Exception as e:
        messagebox.showerror("템플릿 오류", str(e))
        return
    # 본문 치환
    for p in doc.paragraphs:
        txt = "".join(run.text for run in p.runs)
        new_txt = txt
        for k, v in data.items():
            new_txt = new_txt.replace(f"{{{k}}}", v)
        if new_txt != txt:
            p.runs[0].text = new_txt
            for run in p.runs[1:]:
                run.text = ""
    # 테이블 치환
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    txt = "".join(run.text for run in p.runs)
                    new_txt = txt
                    for k, v in data.items():
                        new_txt = new_txt.replace(f"{{{k}}}", v)
                    if new_txt != txt:
                        p.runs[0].text = new_txt
                        for run in p.runs[1:]:
                            run.text = ""
    # 저장
    try:
        doc.save(out_path)
        messagebox.showinfo("성공", f"생성 완료:\n{out_path}")
        refresh_file_list()
    except Exception as e:
        messagebox.showerror("저장 오류", str(e))

# ─── GUI 구성 ───
root = tk.Tk()
root.title("결석계 자동 작성기")
root.configure(bg='#f9f9f9')  # 전체 윈도우 배경색
root.update_idletasks()
root.minsize(root.winfo_width(), root.winfo_height())
for i in range(3):
    root.columnconfigure(i, weight=1)

# 스타일 설정
style = ttk.Style(root)
style.theme_use('default')
style.configure('TFrame', background='#f9f9f9')
style.configure('TLabel', background='#f9f9f9')
style.configure('TLabelframe', background='#f9f9f9')
style.configure('TLabelframe.Label', background='#f9f9f9')
style.configure('Yellow.TCombobox',
                fieldbackground='yellow',
                background='yellow')
style.map('Yellow.TCombobox',
          fieldbackground=[('readonly','yellow'),
                           ('!disabled','yellow')])

# 메인 프레임
main = ttk.Frame(root, padding=15)
main.grid(sticky='nsew')
for i in range(3):
    main.columnconfigure(i, weight=1)

# 제목
ttk.Label(main, text="결석계 자동 작성기",
          font=(None, 20, 'bold')).grid(
    row=0, column=0, columnspan=3, pady=10)

# 정보 입력 프레임
form = ttk.Labelframe(main, text="정보 입력", padding=10)
form.grid(row=1, column=0, sticky='nw', padx=5, pady=5)

# 변수 선언
student_var = tk.StringVar()
grade_var = tk.StringVar()
class_var = tk.StringVar()
number_var = tk.StringVar()
name_var = tk.StringVar()
parent_var = tk.StringVar()
type_var = tk.StringVar(value="출석인정")
start_year = tk.StringVar()
start_month = tk.StringVar()
start_day = tk.StringVar()
end_month = tk.StringVar()
end_day = tk.StringVar()
days_var = tk.StringVar()

# 학생 선택
tk.Label(form, text="학생 선택").grid(
    row=0, column=0, sticky='e', pady=5)
cb = ttk.Combobox(form,
                  textvariable=student_var,
                  values=student_names,
                  state='readonly',
                  style='Yellow.TCombobox')
cb.grid(row=0, column=1, sticky='w', pady=5)
cb.bind('<<ComboboxSelected>>', on_student_selected)

# 수동 입력 필드
fields = [
    ("학년", grade_var),
    ("반", class_var),
    ("번호", number_var),
    ("학생 이름", name_var),
    ("보호자 이름", parent_var)
]
for idx, (lbl, var) in enumerate(fields, start=1):
    ttk.Label(form, text=lbl).grid(
        row=idx, column=0, sticky='e', pady=5)
    ttk.Entry(form, textvariable=var).grid(
        row=idx, column=1, sticky='w', pady=5)

# 결석 구분
tk.Label(form, text="결석 구분").grid(
    row=6, column=0, sticky='e', pady=5)
opt = tk.OptionMenu(form, type_var, "출석인정", "질병", "기타")
opt.config(bg="yellow")
opt.grid(row=6, column=1, sticky='w', pady=5)

# 시작 날짜
tk.Label(form, text="시작").grid(
    row=7, column=0, sticky='e', pady=5)
fr1 = ttk.Frame(form)
fr1.grid(row=7, column=1, sticky='w', pady=5)
for var, w in [(start_year, 4), (start_month, 3), (start_day, 3)]:
    ttk.Entry(fr1, textvariable=var, width=w).pack(side='left', padx=2)
tk.Button(fr1, text="달력", bg="yellow",
          command=lambda: open_calendar('start')
         ).pack(side='left', padx=5)

# 종료 날짜
tk.Label(form, text="종료").grid(
    row=8, column=0, sticky='e', pady=5)
fr2 = ttk.Frame(form)
fr2.grid(row=8, column=1, sticky='w', pady=5)
for var, w in [(end_month, 3), (end_day, 3)]:
    ttk.Entry(fr2, textvariable=var, width=w).pack(side='left', padx=2)
tk.Button(fr2, text="달력", bg="yellow",
          command=lambda: open_calendar('end')
         ).pack(side='left', padx=5)

# 며칠간 (수동 수정 가능, 백그라운드 흰색)
tk.Label(form, text="며칠간").grid(
    row=9, column=0, sticky='e', pady=5)
ent_days = tk.Entry(form, textvariable=days_var,
                    width=5, bg='white')
ent_days.grid(row=9, column=1, sticky='w', pady=5)

# 사유 입력
tk.Label(form, text="사유").grid(
    row=10, column=0, sticky='ne', pady=5)
reason_text = tk.Text(form, width=20, height=3)
reason_text.grid(row=10, column=1, sticky='w', pady=5)

# 생성 버튼
btn_gen = tk.Button(form, text="생성", bg="lightgreen",
                    command=generate_document)
btn_gen.config(width=20, height=2)
btn_gen.grid(row=11, column=0, columnspan=2, pady=10)

# 작성 요령 / 달력 영역
guide_frame = ttk.Labelframe(main,
                             text="작성 요령 / 달력",
                             padding=10)
guide_frame.grid(row=1, column=2, rowspan=2,
                 sticky='nsew', padx=5, pady=5)
msg = (
    "결석 구분: 인정/질병/기타 선택\n"
    "사유 입력: 병명 (생리통, 감기 등)\n"
    "달력 버튼으로 날짜 선택\n"
    "생성 후 개인정보 보호 위해 파일 삭제"
)
ttk.Label(guide_frame, text=msg,
          justify='left', wraplength=200)\
    .grid(row=0, column=0, sticky='nw')
calendar_frame = ttk.Frame(guide_frame)
calendar_frame.grid(row=1, column=0,
                    sticky='nsew', pady=5)
guide_frame.rowconfigure(1, weight=1)
calendar_frame.grid_remove()

# 최근 생성 파일 영역
files_frame = ttk.Labelframe(main,
                             text="최근 생성 파일",
                             padding=10)
files_frame.grid(row=1, column=1, rowspan=2,
                 sticky='nsew', padx=5, pady=5)
files_frame.columnconfigure(0, weight=1)
files_frame.rowconfigure(0, weight=1)
list_files = tk.Listbox(files_frame,
                        width=int(50*1.3/2),
                        height=8)
list_files.grid(row=0, column=0, sticky='nsew')
scroll = ttk.Scrollbar(files_frame,
                       orient='vertical',
                       command=list_files.yview)
scroll.grid(row=0, column=1, sticky='ns')
list_files.config(yscrollcommand=scroll.set)
list_files.bind('<Double-1>', open_file)

# 하단 블로그 링크
def open_blog(evt):
    webbrowser.open("https://blog.naver.com/method917")

link = ttk.Label(main,
                 text="© 2025 메쏘드쌤. All rights reserved.",
                 foreground='blue', cursor='hand2')
link.grid(row=3, column=0, columnspan=3, pady=10)
link.bind('<Button-1>', open_blog)

# 초기 파일 목록 로드
refresh_file_list()
root.mainloop()
