import tkinter as tk
import webbrowser
import os
import sys
from tkinter import messagebox
from docx import Document
from datetime import datetime


def generate_document():
    # 사용자 입력값
    학년 = entry_grade.get()
    반 = entry_class.get()
    번호 = entry_number.get()
    이름 = entry_name.get()
    보호자 = entry_parent.get()
    구분 = var_type.get()
    시작년 = entry_start_year.get()
    시작월 = entry_start_month.get()
    시작일 = entry_start_day.get()
    종료월 = entry_end_month.get()
    종료일 = entry_end_day.get()
    며칠간 = entry_days.get()
    사유 = entry_reason.get("1.0", "end").strip()
    오늘 = datetime.today().strftime("%Y년 %m월 %d일")

    # 📂 템플릿 경로: (pyinstaller용 MEIPASS)
    if getattr(sys, 'frozen', False):
        base_dir = sys._MEIPASS
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    template_path = os.path.join(base_dir, "extract_template", "template_word.docx")

    # 📂 출력 폴더 경로
    if getattr(sys, 'frozen', False):
        program_dir = os.path.dirname(sys.executable)
    else:
        program_dir = os.path.dirname(os.path.abspath(__file__))

    output_dir = os.path.join(program_dir, "output")
    os.makedirs(output_dir, exist_ok=True)

    # 📄 파일명 구성
    start_date_str = f"{시작년.zfill(4)}-{시작월.zfill(2)}-{시작일.zfill(2)}"
    filename_docx = f"결석신고서_{이름}_{start_date_str}.docx"
    output_docx_path = os.path.join(output_dir, filename_docx)

    # 템플릿 열기
    try:
        doc = Document(template_path)
    except Exception as e:
        messagebox.showerror("에러", f"템플릿 파일을 여는 데 실패했습니다:\n{e}")
        return

    # 텍스트 치환
    replacements = {
        "{학년}": 학년, "{반}": 반, "{번호}": 번호, "{이름}": 이름,
        "{보호자}": 보호자, "{구분}": 구분,
        "{시작년}": 시작년, "{시작월}": 시작월, "{시작일}": 시작일,
        "{종료월}": 종료월, "{종료일}": 종료일, "{며칠간}": 며칠간,
        "{사유}": 사유, "{오늘날짜}": 오늘,
        "{학생서명}": 이름, "{보호자서명}": 보호자
    }

    # 파라그래프 치환 (run 단위, 스타일 유지)
    for p in doc.paragraphs:
        full_text = "".join(run.text for run in p.runs)
        new_text = full_text
        for key, value in replacements.items():
            new_text = new_text.replace(key, value)
        if new_text != full_text:
            # 기존 run들에 새 텍스트를 나눠서 재삽입
            # 첫 번째 run에 새 텍스트 넣고, 나머지 run은 비우기
            p.runs[0].text = new_text
            for i in range(1, len(p.runs)):
                p.runs[i].text = ""


    # 테이블 치환 (run 단위, 스타일 유지)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    full_text = "".join(run.text for run in p.runs)
                    new_text = full_text
                    for key, value in replacements.items():
                        new_text = new_text.replace(key, value)
                    if new_text != full_text:
                        p.runs[0].text = new_text
                        for i in range(1, len(p.runs)):
                            p.runs[i].text = ""

    # DOCX 저장
    try:
        doc.save(output_docx_path)
        messagebox.showinfo("성공", f"{output_docx_path}\nDOCX로 저장 완료!")
    except Exception as e:
        with open("error_log.txt", "w", encoding="utf-8") as f:
            f.write(f"DOCX 저장 오류: {e}")
        messagebox.showerror("저장 실패", f"DOCX 저장 중 오류 발생:\n{e}")

    # 디버그용 출력
    print(f"[DEBUG] 템플릿 경로: {template_path}")
    print(f"[DEBUG] 출력 경로: {output_docx_path}")


# ------------------ GUI ------------------

root = tk.Tk()
root.title("결석계 자동 작성기")
root.geometry("650x410")

title_label = tk.Label(root, text="결석계 자동 작성기", font=("맑은 고딕", 16, "bold"))
title_label.grid(row=0, column=0, columnspan=3, pady=(10, 20))

fields = [
    ("학년", 1), ("반", 2), ("번호", 3),
    ("학생 이름", 4), ("보호자 이름", 5)
]
entries = {}
for label, row in fields:
    tk.Label(root, text=label).grid(row=row, column=0, sticky="e", padx=5, pady=3)
    e = tk.Entry(root)
    e.grid(row=row, column=1, padx=5, pady=3, sticky="w")
    entries[label] = e

entry_grade = entries["학년"]
entry_class = entries["반"]
entry_number = entries["번호"]
entry_name = entries["학생 이름"]
entry_parent = entries["보호자 이름"]

tk.Label(root, text="결석 구분").grid(row=6, column=0, sticky="e", padx=5, pady=3)
var_type = tk.StringVar()
var_type.set("출석인정")
option_menu = tk.OptionMenu(root, var_type, "출석인정", "질병", "기타")
option_menu.config(bg="lightyellow")
option_menu.grid(row=6, column=1, sticky="w", padx=5)

tk.Label(root, text="시작 날짜").grid(row=7, column=0, sticky="e", padx=5, pady=3)
frame_start = tk.Frame(root)
frame_start.grid(row=7, column=1, sticky="w", padx=5)
entry_start_year = tk.Entry(frame_start, width=5); entry_start_year.pack(side="left")
tk.Label(frame_start, text="년").pack(side="left")
entry_start_month = tk.Entry(frame_start, width=5); entry_start_month.pack(side="left")
tk.Label(frame_start, text="월").pack(side="left")
entry_start_day = tk.Entry(frame_start, width=5); entry_start_day.pack(side="left")
tk.Label(frame_start, text="일").pack(side="left")

tk.Label(root, text="종료 날짜").grid(row=8, column=0, sticky="e", padx=5, pady=3)
frame_end = tk.Frame(root)
frame_end.grid(row=8, column=1, sticky="w", padx=5)
entry_end_month = tk.Entry(frame_end, width=5); entry_end_month.pack(side="left")
tk.Label(frame_end, text="월").pack(side="left")
entry_end_day = tk.Entry(frame_end, width=5); entry_end_day.pack(side="left")
tk.Label(frame_end, text="일").pack(side="left")
entry_days = tk.Entry(frame_end, width=5); entry_days.pack(side="left")
tk.Label(frame_end, text="일간").pack(side="left")

tk.Label(root, text="사유").grid(row=9, column=0, sticky="ne", padx=5, pady=3)
entry_reason = tk.Text(root, height=2, width=30)
entry_reason.grid(row=9, column=1, columnspan=2, sticky="w", padx=5, pady=3)

tk.Button(root, text="결석계 생성", command=generate_document, bg="lightgreen").grid(row=10, column=1, pady=10)

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
guideline_label.grid(row=0, column=4, rowspan=11, sticky="n", padx=20)

# 하이퍼링크
def open_blog(event):
    webbrowser.open("https://blog.naver.com/method917")

copyright_label = tk.Label(
    root,
    text="© 2025 메쏘드쌤. All rights reserved.",
    font=("맑은 고딕", 9, "underline"),
    fg="blue", cursor="hand2"
)
copyright_label.grid(row=11, column=0, columnspan=5, pady=(20, 5))
copyright_label.bind("<Button-1>", open_blog)

root.mainloop()
