# main.py
import os, json
import tkinter as tk
from tkinter import simpledialog, messagebox
import logging

# 다른 모듈에서 함수 및 설정 가져오기
from core_functions import execute, client # execute 함수와 OpenAI 클라이언트
from data_manager import load_keywords, save_keywords, setup_environment
from config import BASE_DIR, LOG_FILE

# =========================
# 로깅 설정 초기화
# =========================
def setup_logging():
    """로그 시스템을 설정합니다."""
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

    # (add, edit, delete 함수는 기존과 동일)
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
    setup_environment()
    
    # GITHUB_ACTIONS 환경 변수가 'true'이면 GitHub Actions에서 실행 중
    is_github_actions = os.environ.get("GITHUB_ACTIONS") == "true"
    
    if is_github_actions:
        # CLI 환경: execute() 함수만 호출 (GUI 제외)
        try:
            logger.info("GitHub Actions CLI 환경 실행 시작.")
            execute()
            logger.info("GitHub Actions CLI 환경 실행 완료.")
        except Exception as e:
            # GitHub Actions에서 발생한 오류는 로그로 기록하고 실패 상태를 반환
            logger.error(f"GitHub Actions 실행 중 치명적인 오류 발생: {e}")
            # sys.exit(1)을 사용하여 Action을 실패시킬 수도 있지만,
            # 현재는 로그를 남기고 정상 종료합니다.
            raise # 오류를 다시 발생시켜 Action을 실패시킵니다.
    else:
        # 로컬 환경: GUI 실행
        logger.info("로컬 GUI 환경 실행 시작.")
        App().mainloop()

if __name__ == "__main__":
    start_application()