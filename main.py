import tkinter as tk
from tkinter import messagebox, ttk
import webbrowser
import os
import sys
from docx import Document
from datetime import datetime
import pandas as pd

# ——— 전역에서 엑셀 읽기 ———
if getattr(sys, 'frozen', False):
    base_dir = sys._MEIPASS
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

student_dir  = os.path.join(base_dir, "student_list")
student_file = os.path.join(student_dir, "student_list.xlsx")

try:
    df_students = pd.read_excel(student_file, dtype=str)
    student_names = df_students["학생 이름"].tolist()
except Exception as e:
    student_names = []
    # GUI 시작 전에 에러를 띄우려면 root 만들기 전에도 가능하지만,
    # 여기서는 Combobox가 빈 리스트가 되는 정도로 처리합니다.

def on_student_selected(event):
    name = combo_name.get()
    if name not in student_names:
        return
    rec = df_students[df_students["학생 이름"] == name].iloc[0]
    entry_grade.delete(0, "end");  entry_grade.insert(0, rec["학년"])
    entry_class.delete(0, "end");  entry_class.insert(0, rec["반"])
    entry_number.delete(0, "end"); entry_number.insert(0, rec["번호"])
    entry_name.delete(0, "end");   entry_name.insert(0, rec["학생 이름"])
    entry_parent.delete(0, "end"); entry_parent.insert(0, rec["보호자 이름"])

def generate_document():
    # 사용자 입력값
    학년    = entry_grade.get()
    반      = entry_class.get()
    번호    = entry_number.get()
    이름    = entry_name.get()
    보호자  = entry_parent.get()
    구분    = var_type.get()
    시작년  = entry_start_year.get()
    시작월  = entry_start_month.get()
    시작일  = entry_start_day.get()
    종료월  = entry_end_month.get()
    종료일  = entry_end_day.get()
    며칠간  = entry_days.get()
    사유    = entry_reason.get("1.0", "end").strip()
    오늘    = datetime.today().strftime("%Y년 %m월 %d일")

    # 템플릿 경로
    if getattr(sys, 'frozen', False):
        base_dir = sys._MEIPASS
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(base_dir, "extract_template", "template_word.docx")

    # 출력 폴더 경로
    if getattr(sys, 'frozen', False):
        program_dir = os.path.dirname(sys.executable)
    else:
        program_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(program_dir, "output")
    os.makedirs(output_dir, exist_ok=True)

    # 파일명 구성
    start_date_str    = f"{시작년.zfill(4)}-{시작월.zfill(2)}-{시작일.zfill(2)}"
    filename_docx     = f"결석신고서_{이름}_{start_date_str}.docx"
    output_docx_path  = os.path.join(output_dir, filename_docx)

    # 템플릿 열기
    try:
        doc = Document(template_path)
    except Exception as e:
        messagebox.showerror("에러", f"템플릿 파일을 여는 데 실패했습니다:\n{e}")
        return

    # 치환 맵
    replacements = {
        "{학년}": 학년, "{반}": 반, "{번호}": 번호, "{이름}": 이름,
        "{보호자}": 보호자, "{구분}": 구분,
        "{시작년}": 시작년, "{시작월}": 시작월, "{시작일}": 시작일,
        "{종료월}": 종료월, "{종료일}": 종료일, "{며칠간}": 며칠간,
        "{사유}": 사유, "{오늘날짜}": 오늘,
        "{학생서명}": 이름, "{보호자서명}": 보호자
    }

    # 본문 치환
    for p in doc.paragraphs:
        full = "".join(run.text for run in p.runs)
        new  = full
        for k, v in replacements.items():
            new = new.replace(k, v)
        if new != full:
            p.runs[0].text = new
            for run in p.runs[1:]:
                run.text = ""

    # 테이블 치환
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    full = "".join(run.text for run in p.runs)
                    new  = full
                    for k, v in replacements.items():
                        new = new.replace(k, v)
                    if new != full:
                        p.runs[0].text = new
                        for run in p.runs[1:]:
                            run.text = ""

    # 저장
    try:
        doc.save(output_docx_path)
        messagebox.showinfo("성공", f"{output_docx_path}\nDOCX로 저장 완료!")
    except Exception as e:
        with open("error_log.txt", "w", encoding="utf-8") as f:
            f.write(f"DOCX 저장 오류: {e}")
        messagebox.showerror("저장 실패", f"DOCX 저장 중 오류 발생:\n{e}")

# ------------------ GUI ------------------

root = tk.Tk()
root.title("결석계 자동 작성기")
root.geometry("700x460")

# 제목
title_label = tk.Label(root, text="결석계 자동 작성기", font=("맑은 고딕", 16, "bold"))
title_label.grid(row=0, column=0, columnspan=3, pady=(10, 20))

# 학생 선택 드롭다운
tk.Label(root, text="학생 선택").grid(row=1, column=0, sticky="e", padx=5, pady=3)
combo_name = ttk.Combobox(root, values=student_names, state="readonly", width=20)
combo_name.grid(row=1, column=1, padx=5, pady=3, sticky="w")
combo_name.set("선택하세요")
combo_name.bind("<<ComboboxSelected>>", on_student_selected)

# 수동 입력 필드
fields = [
    ("학년", 2), ("반", 3), ("번호", 4),
    ("학생 이름", 5), ("보호자 이름", 6)
]
entries = {}
for label, row in fields:
    tk.Label(root, text=label).grid(row=row, column=0, sticky="e", padx=5, pady=3)
    e = tk.Entry(root)
    e.grid(row=row, column=1, padx=5, pady=3, sticky="w")
    entries[label] = e

entry_grade  = entries["학년"]
entry_class  = entries["반"]
entry_number = entries["번호"]
entry_name   = entries["학생 이름"]
entry_parent = entries["보호자 이름"]

# 결석 구분
tk.Label(root, text="결석 구분").grid(row=7, column=0, sticky="e", padx=5, pady=3)
var_type = tk.StringVar(value="출석인정")
option_menu = tk.OptionMenu(root, var_type, "출석인정", "질병", "기타")
option_menu.config(bg="lightyellow")
option_menu.grid(row=7, column=1, sticky="w", padx=5)

# 시작 날짜
tk.Label(root, text="시작 날짜").grid(row=8, column=0, sticky="e", padx=5, pady=3)
frame_start = tk.Frame(root)
frame_start.grid(row=8, column=1, sticky="w", padx=5)
entry_start_year  = tk.Entry(frame_start, width=5); entry_start_year.pack(side="left")
tk.Label(frame_start, text="년").pack(side="left")
entry_start_month = tk.Entry(frame_start, width=5); entry_start_month.pack(side="left")
tk.Label(frame_start, text="월").pack(side="left")
entry_start_day   = tk.Entry(frame_start, width=5); entry_start_day.pack(side="left")
tk.Label(frame_start, text="일").pack(side="left")

# 종료 날짜
tk.Label(root, text="종료 날짜").grid(row=9, column=0, sticky="e", padx=5, pady=3)
frame_end = tk.Frame(root)
frame_end.grid(row=9, column=1, sticky="w", padx=5)
entry_end_month = tk.Entry(frame_end, width=5); entry_end_month.pack(side="left")
tk.Label(frame_end, text="월").pack(side="left")
entry_end_day   = tk.Entry(frame_end, width=5); entry_end_day.pack(side="left")
tk.Label(frame_end, text="일").pack(side="left")
entry_days      = tk.Entry(frame_end, width=5); entry_days.pack(side="left")
tk.Label(frame_end, text="일간").pack(side="left")

# 사유 입력
tk.Label(root, text="사유").grid(row=10, column=0, sticky="ne", padx=5, pady=3)
entry_reason = tk.Text(root, height=2, width=30)
entry_reason.grid(row=10, column=1, columnspan=2, sticky="w", padx=5, pady=3)

# 버튼
tk.Button(root, text="결석계 생성", command=generate_document, bg="lightgreen")\
    .grid(row=11, column=1, pady=10)

# 작성 요령
guideline_text = (
    "📌 결석계 작성법\n\n"
    "결석 구분은 인정, 질병, 기타 중 선택합니다.\n\n"
    "사유는 병명, 질환명을 입력합니다.\n"
    "- 생리통\n- 발목염좌\n- 감기\n- 인후염\n\n"
    "작성을 마친 후 '결석계 생성' 버튼을 누르세요.\n"
    "출력 후 개인정보 보호를 위해 삭제하세요."
)
guideline_label = tk.Label(
    root, text=guideline_text,
    justify="left", anchor="nw",
    padx=10, pady=10,
    bg="#f4f4f4", relief="groove",
    width=40, height=20
)
guideline_label.grid(row=1, column=4, rowspan=11, sticky="n", padx=20)

# 하이퍼링크
def open_blog(event):
    webbrowser.open("https://blog.naver.com/method917")

copyright_label = tk.Label(
    root,
    text="© 2025 메쏘드쌤. All rights reserved.",
    font=("맑은 고딕", 9, "underline"),
    fg="blue", cursor="hand2"
)
copyright_label.grid(row=12, column=0, columnspan=5, pady=(20, 5))
copyright_label.bind("<Button-1>", open_blog)

root.mainloop()
