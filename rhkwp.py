import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import ttk, messagebox

# ========== matplotlib 한글 폰트 깨짐 방지 (Windows 기준) ==========
plt.rc('font', family='Malgun Gothic')  # 한글 글꼴 설정
plt.rc('axes', unicode_minus=False)     # 음수(-) 깨짐 방지

# ========== 엑셀 파일 이름 ==========
EXCEL_FILE = "인공지능기술_AI필요성.xlsx"

# ========== 분석 대상 산업 목록 ==========
TARGET_COMPANIES = [
    "제조업",
    "건설업",
    "도매및소매업",
    "정보통신업",
    "전문,과학및기술서비스업",
]

# ========== 엑셀 컬럼명이므로 반드시 실제 엑셀과 일치해야 함 ==========
COL_SECTOR1 = "특성별(1)"          # 산업 구분
COL_VERY    = "매우 필요"          # AI 매우 필요 비율
COL_SOME    = "약간 필요"          # AI 약간 필요 비율
COL_LESS    = "별로 필요하지 않음" # AI 별로 필요하지 않음 비율
COL_NEVER   = "전혀 필요하지 않음" # AI 전혀 필요하지 않음 비율


# ==============================================================
#  📌 2. 데이터 불러오기 + 전처리 (비즈니스 로직)
# ==============================================================
def load_data():
    """
    엑셀 데이터를 읽고,
    산업별 AI '필요함'과 '필요하지 않음' 비율을 계산하여 반환
    """
    # Header=2 → 엑셀에서 3번째 줄이 실제 컬럼명
    df = pd.read_excel(EXCEL_FILE, header=2)

    # 필요한 열만 선택하여 새로운 이름으로 지정
    df = df[[COL_SECTOR1, COL_VERY, COL_SOME, COL_LESS, COL_NEVER]].copy()
    df.columns = ["company", "very_need", "some_need", "less_need", "never_need"]

    # 데이터가 숫자가 아니면 NaN 처리 → 계산 가능하게 변환
    for col in ["very_need", "some_need", "less_need", "never_need"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # 산업 이름 없는 행 제거
    df = df.dropna(subset=["company"])

    # 필요함 비율 합(매우필요 + 약간필요)
    df["need_total"] = df["very_need"] + df["some_need"]
    # 필요하지 않음 비율 합(별로 + 전혀)
    df["noneed_total"] = df["less_need"] + df["never_need"]

    return df


# ==============================================================
#  📌 3. Tkinter GUI 클래스 (UI 화면 구성)
# ==============================================================
class AiNeedApp(tk.Tk):
    def __init__(self, df):
        super().__init__()

        # GUI 기본 설정
        self.title("산업별 디지털·AI 활용 필요성 분석")
        self.geometry("620x420")
        self.resizable(False, False)

        self.df = df
        # 분석 대상 산업만 필터링하여 콤보박스에 표시
        self.company_list = sorted(
            [c for c in self.df["company"].unique() if c in TARGET_COMPANIES]
        )

        self.create_widgets()  # 화면 요소 생성

    # ------ 화면 구성 요소 생성 ------
    def create_widgets(self):
        # 제목 라벨
        title = tk.Label(
            self,
            text="산업별 디지털·AI 활용 필요성 분석",
            font=("맑은 고딕", 16, "bold")
        )
        title.pack(pady=10)

        # 설명 문구
        desc = tk.Label(
            self,
            text="산업을 선택하면 해당 산업의 AI 필요성 평균 비율을 알려줍니다.",
            font=("맑은 고딕", 10)
        )
        desc.pack(pady=5)

        # 산업 선택 영역
        top_frame = tk.Frame(self)
        top_frame.pack(pady=10)

        tk.Label(top_frame, text="산업 선택:", font=("맑은 고딕", 11)).grid(row=0, column=0, padx=5, pady=5)

        # 콤보박스(드롭다운)
        self.company_var = tk.StringVar()
        self.company_combo = ttk.Combobox(
            top_frame,
            textvariable=self.company_var,
            values=self.company_list,
            state="readonly",
            width=25
        )
        # 기본 선택 값 설정
        if self.company_list:
            self.company_combo.set(self.company_list[0])
        self.company_combo.grid(row=0, column=1, padx=5, pady=5)

        # 선택 산업 분석 버튼
        analyze_btn = tk.Button(
            top_frame,
            text="선택 산업 분석하기",
            command=self.analyze_selected_company,
            bg="#2563eb",
            fg="white",
            padx=10,
            pady=5
        )
        analyze_btn.grid(row=0, column=2, padx=10, pady=5)

        # 산업 전체 비교 버튼
        compare_btn = tk.Button(
            top_frame,
            text="산업 전체 비교 그래프",
            command=self.show_company_comparison,
            bg="#16a34a",
            fg="white",
            padx=10,
            pady=5
        )
        compare_btn.grid(row=1, column=1, columnspan=2, pady=5)

        # 결과 표시 제목
        result_label = tk.Label(self, text="분석 결과:", font=("맑은 고딕", 12, "bold"))
        result_label.pack(anchor="w", padx=20)

        # 결과 출력 텍스트 박스
        self.result_text = tk.Text(self, height=10, width=75)
        self.result_text.pack(padx=20, pady=5)

        # 초기 안내 메시지
        self.result_text.insert(
            tk.END,
            "1) 산업을 선택하고 [선택 산업 분석하기] 클릭\n"
            "2) 전체 비교 그래프도 확인해보세요!\n"
        )

    # ======================================================
    # 🎯 선택한 산업 분석
    # ======================================================
    def analyze_selected_company(self):
        company = self.company_var.get()

        if not company:
            messagebox.showwarning("주의", "먼저 산업을 선택해주세요.")
            return

        # 선택한 산업 데이터만 추출
        df_sub = self.df[self.df["company"] == company]

        if df_sub.empty:
            messagebox.showinfo("정보", f"{company} 데이터가 없습니다.")
            return

        # 평균 비율 계산
        need = df_sub["need_total"].mean()
        noneed = df_sub["noneed_total"].mean()

        # 결과 출력 영역 초기화
        self.result_text.delete("1.0", tk.END)

        self.result_text.insert(tk.END, f"[선택 산업] {company}\n\n")
        self.result_text.insert(tk.END, f"- AI '필요함' 평균 비율: {need:.1f}%\n")
        self.result_text.insert(tk.END, f"- AI '필요하지 않음' 평균 비율: {noneed:.1f}%\n\n")

        # 간단한 해석 추가
        if need >= 25:
            level = "AI 필요성이 매우 높은 산업"
        elif need >= 15:
            level = "평균보다 다소 높은 산업"
        else:
            level = "상대적으로 낮은 산업"

        self.result_text.insert(tk.END, f"[해석]\n{level}으로 볼 수 있습니다.\n")

        # 그래프 띄우기 여부 확인
        show = messagebox.askyesno("그래프 보기", "해당 산업을 그래프로 볼까요?")
        if show:
            plt.figure()
            plt.bar(["필요함", "필요하지 않음"], [need, noneed])
            plt.ylim(0, 100)
            plt.title(f"{company} - AI 필요성")
            plt.ylabel("비율(%)")
            plt.show()

    # ======================================================
    # 🎯 산업 전체 비교 그래프
    # ======================================================
    def show_company_comparison(self):
        df_sub = self.df[self.df["company"].isin(TARGET_COMPANIES)].copy()

        if df_sub.empty:
            messagebox.showinfo("정보", "비교할 데이터가 없습니다.")
            return

        # 산업별 평균 비교
        grouped = (
            df_sub.groupby("company")[["need_total", "noneed_total"]]
            .mean()
            .sort_values("need_total", ascending=False)
        )

        # 텍스트 박스 초기화 + 결과 표시
        self.result_text.delete("1.0", tk.END)
        self.result_text.insert(tk.END, "[산업별 AI 필요함 평균 비율]\n\n")
        for idx, row in grouped.iterrows():
            self.result_text.insert(
                tk.END,
                f"- {idx}: 필요함 {row['need_total']:.1f}%, "
                f"필요하지 않음 {row['noneed_total']:.1f}%\n"
            )

        # 막대그래프 출력
        plt.figure(figsize=(8, 5))
        plt.bar(grouped.index, grouped["need_total"])
        plt.xticks(rotation=45, ha="right")
        plt.ylabel("AI 필요함 비율(%)")
        plt.title("산업별 AI 필요성 비교")
        plt.ylim(0, 100)
        plt.tight_layout()
        plt.show()


# ==============================================================
#  📌 4. 프로그램 실행 (메인 엔트리 포인트)
# ==============================================================
def main():
    try:
        df = load_data()  # 데이터 불러오기
    except Exception as e:
        print("엑셀 파일 읽기 실패:", e)
        return

    app = AiNeedApp(df)  # GUI 실행
    app.mainloop()


# 실행 시 바로 main 함수 호출
if __name__ == "__main__":
    main()
