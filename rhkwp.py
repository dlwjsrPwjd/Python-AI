import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import sys
import os
import tkinter as tk
from tkinter import ttk, messagebox

# --- 폰트 설정 (운영체제별 한글 폰트 설정) ---
if sys.platform == "darwin":
    MPL_FONT = "AppleGothic"
    TK_FONT = "AppleGothic"
elif sys.platform.startswith("win"):
    MPL_FONT = "Malgun Gothic"
    TK_FONT = "맑은 고딕"
else:
    MPL_FONT = "NanumGothic"
    TK_FONT = "NanumGothic"

plt.rc('font', family=MPL_FONT)
plt.rc('axes', unicode_minus=False)

def resource_path(rel_path):
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, rel_path)

EXCEL_FILE = resource_path("인공지능기술_AI필요성.xlsx")
TARGET_COMPANIES = [
    "제조업",
    "건설업",
    "도매및소매업",
    "정보통신업",
    "전문,과학및기술서비스업",
]
COL_SECTOR1 = "특성별(1)"
COL_VERY = "매우 필요"
COL_SOME = "약간 필요"
COL_LESS = "별로 필요하지 않음"
COL_NEVER = "전혀 필요하지 않음"

def load_data():
    df = pd.read_excel(EXCEL_FILE, header=2)
    df = df[[COL_SECTOR1, COL_VERY, COL_SOME, COL_LESS, COL_NEVER]].copy()
    df.columns = ["company", "very_need", "some_need", "less_need", "never_need"]
    for col in ["very_need", "some_need", "less_need", "never_need"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df = df.dropna(subset=["company"])
    df["need_total"] = df["very_need"] + df["some_need"]
    df["noneed_total"] = df["less_need"] + df["never_need"]
    return df

class AiNeedApp(tk.Tk):
    def __init__(self, df):
        super().__init__()
        self.title("산업별 디지털·AI 활용 필요성 분석")
        self.geometry("640x450") # 창 크기 약간 증가
        self.resizable(False, False)
        self.df = df
        self.company_list = sorted([c for c in self.df["company"].unique() if c in TARGET_COMPANIES])
        self.configure(bg="#f0f0f0") # 전체 배경색 설정
        self.setup_style() # 스타일 설정 메서드 추가
        self.create_widgets()

    def setup_style(self):
        """ttk.Style을 사용하여 위젯의 디자인을 개선합니다."""
        style = ttk.Style(self)
        style.theme_use('clam') # 'clam' 테마 사용 (좀 더 모던함)

        # 폰트 설정 적용
        style.configure('.', font=(TK_FONT, 10))

        # 제목 (Heading) 스타일
        self.option_add('*TButton*Font', (TK_FONT, 10, 'bold'))

        # 콤보박스 스타일
        style.configure('TCombobox', 
                        fieldbackground='white', 
                        background='white',
                        selectbackground='white',
                        selectforeground='black',
                        bordercolor='#a0a0a0',
                        relief='flat')

        # 분석 버튼 스타일 (진한 파랑)
        style.configure('Analyze.TButton', 
                        background='#3b82f6', 
                        foreground='white',
                        font=(TK_FONT, 10, 'bold'),
                        padding=[10, 5])
        style.map('Analyze.TButton', 
                   background=[('active', '#2563eb')])

        # 비교 버튼 스타일 (진한 초록)
        style.configure('Compare.TButton', 
                        background='#10b981', 
                        foreground='white',
                        font=(TK_FONT, 10, 'bold'),
                        padding=[10, 5])
        style.map('Compare.TButton', 
                   background=[('active', '#059669')])
        
    def create_widgets(self):
        # 중앙 정렬을 위한 컨테이너 프레임
        main_frame = tk.Frame(self, bg="#f0f0f0")
        main_frame.pack(pady=20, padx=20, fill='both')

        # 제목 라벨 (tk.Label 사용, 폰트 크기 및 두께 강조)
        title = tk.Label(main_frame, text="산업별 디지털·AI 활용 필요성 분석", 
                         font=(TK_FONT, 18, "bold"), bg="#f0f0f0", fg="#1e3a8a")
        title.pack(pady=5)

        # 설명 라벨
        desc = tk.Label(main_frame, text="산업을 선택하고 분석 버튼을 눌러보세요.", 
                        font=(TK_FONT, 11), bg="#f0f0f0")
        desc.pack(pady=5)

        # 상단 입력/버튼 프레임 (배경색 통일)
        top_frame = tk.Frame(main_frame, bg="#f0f0f0")
        top_frame.pack(pady=10)

        tk.Label(top_frame, text="산업 선택:", font=(TK_FONT, 11, 'bold'), 
                 bg="#f0f0f0").grid(row=0, column=0, padx=5, pady=5)
        
        self.company_var = tk.StringVar()
        # ttk.Combobox 적용
        self.company_combo = ttk.Combobox(top_frame, textvariable=self.company_var, 
                                          values=self.company_list, state="readonly", width=25, 
                                          style='TCombobox')
        if self.company_list:
            self.company_combo.set(self.company_list[0])
        self.company_combo.grid(row=0, column=1, padx=10, pady=5)
        
        # ttk.Button 및 스타일 적용 (Analyze.TButton)
        analyze_btn = ttk.Button(top_frame, text="선택 산업 분석하기", 
                                 command=self.analyze_selected_company, 
                                 style='Analyze.TButton')
        analyze_btn.grid(row=0, column=2, padx=10, pady=5)

        # ttk.Button 및 스타일 적용 (Compare.TButton)
        compare_btn = ttk.Button(top_frame, text="산업 전체 비교 그래프 보기", 
                                 command=self.show_company_comparison, 
                                 style='Compare.TButton')
        compare_btn.grid(row=1, column=1, columnspan=2, pady=10)

        # 결과 표시 영역 (Result Area)
        result_label = tk.Label(main_frame, text="📊 분석 결과:", 
                                font=(TK_FONT, 12, "bold"), bg="#f0f0f0", fg="#1e3a8a")
        result_label.pack(anchor="w", padx=10, pady=(5, 0))
        
        # 텍스트 위젯 가독성 개선: 테두리 제거, 배경 흰색
        self.result_text = tk.Text(main_frame, height=10, width=70, 
                                   bd=0, relief="flat", bg="white", padx=10, pady=10, 
                                   font=(TK_FONT, 10))
        self.result_text.pack(padx=10, pady=5)
        self.result_text.insert(tk.END, "1) 산업을 선택하고 [선택 산업 분석하기] 클릭\n2) 전체 비교 그래프도 확인해보세요!\n")

    # (이하 분석 및 그래프 메서드는 변경 없음)
    def analyze_selected_company(self):
        company = self.company_var.get()
        if not company:
            messagebox.showwarning("주의", "먼저 산업을 선택해주세요.")
            return
        df_sub = self.df[self.df["company"] == company]
        if df_sub.empty:
            messagebox.showinfo("정보", f"{company} 데이터가 없습니다.")
            return
        need = df_sub["need_total"].mean()
        noneed = df_sub["noneed_total"].mean()
        self.result_text.delete("1.0", tk.END)
        self.result_text.insert(tk.END, f"[선택 산업] {company}\n\n")
        self.result_text.insert(tk.END, f"- AI '필요함' 평균 비율: {need:.1f}%\n")
        self.result_text.insert(tk.END, f"- AI '필요하지 않음' 평균 비율: {noneed:.1f}%\n\n")
        if need >= 25:
            level = "AI 필요성이 매우 높은 산업"
        elif need >= 15:
            level = "평균보다 다소 높은 산업"
        else:
            level = "상대적으로 낮은 산업"
        self.result_text.insert(tk.END, f"[해석]\n{level}으로 볼 수 있습니다.\n")
        show = messagebox.askyesno("그래프 보기", "해당 산업을 그래프로 볼까요?")
        if show:
            plt.figure()
            plt.bar(["필요함", "필요하지 않음"], [need, noneed], color=['#2563eb', '#9ca3af']) # 막대 색상 지정
            plt.ylim(0, 100)
            plt.title(f"{company} - AI 필요성")
            plt.ylabel("비율(%)")
            plt.show()

    def show_company_comparison(self):
        df_sub = self.df[self.df["company"].isin(TARGET_COMPANIES)].copy()
        if df_sub.empty:
            messagebox.showinfo("정보", "비교할 데이터가 없습니다.")
            return
        grouped = (df_sub.groupby("company")[ ["need_total", "noneed_total"] ].mean().sort_values("need_total", ascending=False))
        self.result_text.delete("1.0", tk.END)
        self.result_text.insert(tk.END, "[산업별 AI 필요함 평균 비율]\n\n")
        for idx, row in grouped.iterrows():
            self.result_text.insert(tk.END, f"- {idx}: 필요함 {row['need_total']:.1f}%, 필요하지 않음 {row['noneed_total']:.1f}%\n")
        
        plt.figure(figsize=(8, 5))
        plt.bar(grouped.index, grouped["need_total"], color='#10b981') # 막대 색상 지정
        plt.xticks(rotation=45, ha="right")
        plt.ylabel("AI 필요함 비율(%)")
        plt.title("산업별 AI 필요성 비교")
        plt.ylim(0, 100)
        plt.tight_layout()
        plt.show()

def main():
    try:
        df = load_data()
    except Exception as e:
        print("엑셀 파일 읽기 실패:", e)
        # GUI를 통해 사용자에게 오류 알림
        root = tk.Tk()
        root.withdraw() # 메인 윈도우 숨기기
        messagebox.showerror("오류", f"엑셀 파일 ({EXCEL_FILE}) 읽기 실패: {e}\n파일이 현재 디렉토리에 있는지 확인하세요.")
        root.destroy()
        return
        
    app = AiNeedApp(df)
    app.mainloop()

if __name__ == "__main__":
    main()