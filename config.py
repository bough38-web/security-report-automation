# config.py

import os
import sys
import tkinter as tk
from tkinter import simpledialog, messagebox
import logging

# 1. 설정 및 데이터 매니저 임포트
from config import BASE_DIR, LOG_FILE
from data_manager import load_keywords, save_keywords, setup_environment

# 2. [핵심] 환경 변수 로드 (API Key 세팅)
# core_functions를 불러오기 전에 반드시 실행되어야 합니다.
setup_environment()

# 3. 핵심 기능 임포트
try:
    from core_functions import execute
except Exception as e:
    print(f"[CRITICAL] core_functions 모듈 로드 실패: {e}")
    sys.exit(1)

# =========================
# 로깅 설정
# =========================
def setup_logging():
    log_dir = os.path.dirname(LOG_FILE)
    if log_dir and not os.path.exists(log_dir):
        os.makedirs(log_dir)

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(LOG_FILE, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger("Main")

logger = setup_logging()

# =========================
# GUI 클래스
# =========================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("사내 보안·안전 자동 보고 시스템")
        self.geometry("360x420")
        self.resizable(False, False) 
        
        tk.Label(self, text="현재 검색 키워드 목록:", anchor="w").pack(padx=5, pady=(5,0), fill="x")
        
        frame = tk.Frame(self)
        frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.list = tk.Listbox(frame)
        self.list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = tk.Scrollbar(frame, command=self.list.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.list.config(yscrollcommand=scrollbar.set)
        
        self.refresh()

        button_frame = tk.Frame(self)
        button_frame.pack(fill="x", padx=5, pady=5)

        tk.Button(button_frame, text="키워드 추가", command=self.add).pack(fill="x")
        tk.Button(button_frame, text="키워드 수정", command=self.edit).pack(fill="x")
        tk.Button(button_frame, text="키워드 삭제", command=self.delete).pack(fill="x")
        
        tk.Button(self, text="보고서 생성 실행", bg="green", fg="white",
                  font=("Helvetica", 12, "bold"), command=self._execute_wrapper).pack(fill="x", padx=5, pady=(0, 5))

    def _execute_wrapper(self):
        try:
            execute()
            messagebox.showinfo("완료", "보고서 생성이 완료되었습니다!")
        except Exception as e:
            logger.error(f"실행 오류: {e}")
            messagebox.showerror("오류", f"실행 중 오류 발생: {e}")

    def refresh(self):
        self.list.delete(0, tk.END)
        for k in load_keywords():
            self.list.insert(tk.END, k)

    def add(self):
        v = simpledialog.askstring("키워드 추가", "키워드 입력:")
        if v and v.strip():
            data = load_keywords()
            if v.strip() not in data:
                data.append(v.strip())
                save_keywords(data)
                self.refresh()

    def edit(self):
        sel = self.list.curselection()
        if not sel: return
        data = load_keywords()
        idx = sel[0]
        new = simpledialog.askstring("수정", "수정할 내용:", initialvalue=data[idx])
        if new and new.strip():
            data[idx] = new.strip()
            save_keywords(data)
            self.refresh()

    def delete(self):
        sel = self.list.curselection()
        if not sel: return
        if messagebox.askyesno("삭제", "정말 삭제하시겠습니까?"):
            data = load_keywords()
            del data[sel[0]]
            save_keywords(data)
            self.refresh()

# =========================
# 메인 실행부
# =========================
def start_application():
    # GitHub Actions 환경 감지
    is_github_actions = os.environ.get("GITHUB_ACTIONS") == "true"
    
    if is_github_actions:
        try:
            logger.info("GitHub Actions CLI 모드로 실행합니다.")
            execute()
            logger.info("실행 완료.")
        except Exception as e:
            logger.error(f"GitHub Actions 오류: {e}")
            sys.exit(1) # 오류 발생 시 Action 실패 처리
    else:
        logger.info("로컬 GUI 모드로 실행합니다.")
        app = App()
        app.mainloop()

if __name__ == "__main__":
    start_application()

# =========================
# 시스템 설정
# =========================
PPT_TITLE = "보안·안전 자동 보고"
PPT_FOOTER = (
    "본 보고서는 외부 공개 뉴스 기반 자동 수집 자료이며\n"
    "AI 요약은 참고용으로 활용됩니다."
)
PPT_DATE = datetime.now().strftime("%Y-%m-%d")

# =========================
# 뉴스 및 AI 설정
# =========================
GOOGLE_NEWS_URL = "https://news.google.com/rss/search?q={q}&hl=ko&gl=KR&ceid=KR:ko"
MAX_NEWS_ENTRIES = 100
OPENAI_MODEL = "gpt-4o-mini"
MAX_TOKENS = 300

# =========================
# 리스크 키워드
# =========================
RISK_KEYWORDS = {
    "RED": ["사망", "침입", "해킹", "중대", "폭발", "재난"],
    "AMBER": ["사고", "논란", "장애", "위험", "유출", "화재"]
}

# =========================
# 이메일 설정
# =========================
MAIL_TO = "bough38@gmail.com"
MAIL_CC = "heebon.park@kt.com"
MAIL_SUBJECT = f"[자동] 보안·안전 보고 ({datetime.now().date()})"

# =========================
# 초기 키워드 파일 생성
# =========================
INITIAL_KEYWORDS = ["KT텔레캅","에스원","쉴더스","보안 사고","안전 사고"]
