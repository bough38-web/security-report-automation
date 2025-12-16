# main.py
import os
import sys
import tkinter as tk
from tkinter import simpledialog, messagebox
import logging

# 1. 설정 및 데이터 매니저를 '가장 먼저' 임포트합니다.
from config import BASE_DIR, LOG_FILE
from data_manager import load_keywords, save_keywords, setup_environment

# 2. [중요] 환경 변수 로드를 다른 모든 무거운 모듈 임포트보다 먼저 실행합니다.
# 이렇게 해야 이후 임포트되는 모듈들이 API Key를 인식할 수 있습니다.
setup_environment()

# 3. 환경 설정이 완료된 후 핵심 기능(OpenAI 클라이언트 포함)을 임포트합니다.
try:
    from core_functions import execute, client 
except Exception as e:
    # API 키 오류 등으로 임포트 자체가 실패할 경우를 대비한 예외 처리
    print(f"[CRITICAL] core_functions 모듈 로드 실패: {e}")
    # GitHub Actions 환경이 아니라면 GUI 알림 시도 (Tkinter 초기화 전이라 제한적일 수 있음)
    if os.environ.get("GITHUB_ACTIONS") != "true":
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, f"모듈 로드 실패: {e}", "치명적 오류", 0)
    sys.exit(1)


# =========================
# 로깅 설정 초기화
# =========================
def setup_logging():
    """로그 시스템을 설정합니다."""
    # 로그 디렉토리가 없으면 생성 (안전장치)
    log_dir = os.path.dirname(LOG_FILE)
    if not os.path.exists(log_dir):
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
# GUI (키워드 관리 및 실행)
# =========================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("사내 보안·안전 자동 보고 시스템")
        self.geometry("360x420")
        self.resizable(False, False) 
        
        # UI: Listbox 레이블 추가
        tk.Label(self, text="현재 검색 키워드 목록:", anchor="w").pack(padx=5, pady=(5,0), fill="x")
        
        # Listbox 프레임 및 스크롤바
        frame = tk.Frame(self)
        frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.list = tk.Listbox(frame)
        self.list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = tk.Scrollbar(frame, command=self.list.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.list.config(yscrollcommand=scrollbar.set)
        
        self.refresh()

        # 버튼 프레임
        button_frame = tk.Frame(self)
        button_frame.pack(fill="x", padx=5, pady=5)

        tk.Button(button_frame, text="키워드 추가", command=self.add).pack(fill="x")
        tk.Button(button_frame, text="키워드 수정", command=self.edit).pack(fill="x")
        tk.Button(button_frame, text="키워드 삭제", command=self.delete).pack(fill="x")
        
        # 메인 실행 버튼
        tk.Button(self, text="보고서 자동 생성 및 메일 발송 (실행)", bg="green", fg="white",
                  font=("Helvetica", 12, "bold"), command=self._execute_wrapper).pack(fill="x", padx=5, pady=(0, 5))

    def _execute_wrapper(self):
        """execute 함수를 호출하고 GUI 메시지를 처리하는 래퍼"""
        try:
            execute()
            messagebox.showinfo("완료", "자동 보고 시스템 실행 및 이메일 발송 완료!")
        except Exception as e:
            logger.error(f"실행 오류 발생: {e}")
            messagebox.showerror("오류", f"실행 중 오류 발생: {e}\n자세한 내용은 로그 파일 확인")

    def refresh(self):
        self.list.delete(0, tk.END)
        for k in load_keywords():
            self.list.insert(tk.END, k)

    def add(self):
        v = simpledialog.askstring("키워드 추가", "새로운 키워드를 입력하세요.")
        if v and v.strip():
            data = load_keywords()
            if v.strip() not in data:
                data.append(v.strip())
                save_keywords(data)
                self.refresh()
            else:
                messagebox.showwarning("중복", f"'{v.strip()}' 키워드는 이미 존재합니다.")

    def edit(self):
        sel = self.list.curselection()
        if not sel: 
            messagebox.showwarning("선택 오류", "수정할 키워드를 목록에서 선택해 주세요.")
            return
        idx = sel[0]
        data = load_keywords()
        new = simpledialog.askstring("키워드 수정", "새 키워드를 입력하세요.", initialvalue=data[idx])
        if new and new.strip():
            new_keyword = new.strip()
            if new_keyword != data[idx]:
                if new_keyword in data:
                    messagebox.showwarning("중복", f"'{new_keyword}' 키워드는 이미 존재합니다.")
                else:
                    data[idx] = new_keyword
                    save_keywords(data)
                    self.refresh()

    def delete(self):
        sel = self.list.curselection()
        if not sel: 
            messagebox.showwarning("선택 오류", "삭제할 키워드를 목록에서 선택해 주세요.")
            return
        
        if messagebox.askyesno("삭제 확인", "선택한 키워드를 정말 삭제하시겠습니까?"):
            idx = sel[0]
            data = load_keywords()
            data.pop(idx)
            save_keywords(data)
            self.refresh()


# =========================
# 진입점 (실행 환경 구분)
# =========================
def start_application():
    """실행 환경(CLI vs GUI)을 구분하여 애플리케이션을 시작합니다."""
    # setup_environment()는 이미 상단에서 실행되었으므로 여기서 다시 호출할 필요가 없습니다.
    
    # GITHUB_ACTIONS 환경 변수가 'true'이면 GitHub Actions에서 실행 중
    is_github_actions = os.environ.get("GITHUB_ACTIONS") == "true"
    
    if is_github_actions:
        # CLI 환경: execute() 함수만 호출 (GUI 제외)
        try:
            logger.info("GitHub Actions CLI 환경 실행 시작.")
            execute()
            logger.info("GitHub Actions CLI 환경 실행 완료.")
        except Exception as e:
            logger.error(f"GitHub Actions 실행 중 치명적인 오류 발생: {e}")
            raise # 오류를 다시 발생시켜 GitHub Action이 실패로 표시되게 함
    else:
        # 로컬 환경: GUI 실행
        logger.info("로컬 GUI 환경 실행 시작.")
        try:
            app = App()
            app.mainloop()
        except KeyboardInterrupt:
            logger.info("사용자에 의해 프로그램 종료")

if __name__ == "__main__":
    start_application()
