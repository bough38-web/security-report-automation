import os
from datetime import datetime

# 1. 경로 설정
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = os.path.join(BASE_DIR, "app.log")
DATA_FILE = os.path.join(BASE_DIR, "keywords.json")

# 2. 시스템 설정
PPT_TITLE = "보안·안전 자동 보고"
PPT_FOOTER = (
    "본 보고서는 외부 공개 뉴스 기반 자동 수집 자료이며\n"
    "AI 요약은 참고용으로 활용됩니다."
)
PPT_DATE = datetime.now().strftime("%Y-%m-%d")

# 3. 뉴스 및 AI 설정
GOOGLE_NEWS_URL = "https://news.google.com/rss/search?q={q}&hl=ko&gl=KR&ceid=KR:ko"
MAX_NEWS_ENTRIES = 100
OPENAI_MODEL = "gpt-4o-mini"
MAX_TOKENS = 300

# 4. 리스크 키워드
RISK_KEYWORDS = {
    "RED": ["사망", "침입", "해킹", "중대", "폭발", "재난"],
    "AMBER": ["사고", "논란", "장애", "위험", "유출", "화재"]
}

# 5. 이메일 설정
MAIL_TO = "bough38@gmail.com"
MAIL_CC = "heebon.park@kt.com"
MAIL_SUBJECT = f"[자동] 보안·안전 보고 ({datetime.now().date()})"

# 6. 초기 키워드
INITIAL_KEYWORDS = ["KT텔레캅","에스원","쉴더스","보안 사고","안전 사고"]
