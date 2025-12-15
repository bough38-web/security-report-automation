# core_functions.py
import urllib.parse
import feedparser
import pandas as pd
import logging
from openai import OpenAI
# import win32com.client as win32
from datetime import datetime
from config import (
    GOOGLE_NEWS_URL, MAX_NEWS_ENTRIES, RISK_KEYWORDS,
    OPENAI_MODEL, MAX_TOKENS, MAIL_TO, MAIL_CC, MAIL_SUBJECT
)

# 로깅 설정
logger = logging.getLogger(__name__)

# OpenAI 클라이언트 초기화
try:
    client = OpenAI()
except Exception as e:
    logger.error(f"OpenAI 클라이언트 초기화 오류 (OPENAI_API_KEY 확인): {e}")
    # 실행 엔진에서 이 오류를 처리하도록 예외는 전파
    raise

def crawl_news(keywords: list) -> pd.DataFrame:
    """구글 RSS를 이용하여 뉴스 크롤링"""
    rows = []
    logger.info(f"뉴스 크롤링 시작. 키워드: {keywords}")
    for kw in keywords:
        try:
            q = urllib.parse.quote(kw)
            url = GOOGLE_NEWS_URL.format(q=q)
            feed = feedparser.parse(url)
            
            for e in feed.entries[:MAX_NEWS_ENTRIES]:
                rows.append({
                    "keyword": kw,
                    "title": e.title,
                    "link": e.link,
                    # 뉴스 발행 시간 추가 (없을 경우 현재 시간)
                    "published": getattr(e, 'published', datetime.now().isoformat())
                })
            logger.debug(f"키워드 '{kw}': {len(feed.entries)}개 항목 크롤링 완료.")
        except Exception as e:
            logger.error(f"'{kw}' 키워드 크롤링 오류: {e}")
            
    df = pd.DataFrame(rows)
    df.drop_duplicates(subset=["title", "link"], inplace=True) # 중복 기사 제거
    logger.info(f"전체 뉴스 크롤링 완료. 총 {len(df)}개 기사.")
    return df

def risk_level(title: str) -> str:
    """뉴스 제목 기반 리스크 레벨 분류"""
    for r, words in RISK_KEYWORDS.items():
        if any(w in title for w in words):
            return r
    return "GREEN"

def ai_summary(text: str) -> str:
    """OpenAI API를 이용한 텍스트 요약"""
    prompt = (
        "다음은 보안 및 안전 관련 뉴스 기사 제목들입니다. "
        "이를 종합하여 'OOO 관련' 제목으로 시작하는 3줄 이내의 간결한 핵심 요약본을 작성해 주세요. "
        f"제목: {text}"
    )
    try:
        res = client.chat.completions.create(
            model=OPENAI_MODEL,
            messages=[{"role":"user","content":prompt}],
            max_tokens=MAX_TOKENS,
            temperature=0.3 # 일관성을 위해 낮은 temperature 설정
        )
        summary = res.choices[0].message.content.strip()
        logger.debug(f"AI 요약 성공. 내용: {summary[:50]}...")
        return summary
    except Exception as e:
        logger.error(f"AI 요약 오류: {e}")
        return f"AI 요약 실패: {e}"

def send_mail(zip_path: str, body: str):
    """Outlook을 이용한 이메일 발송"""
    try:
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = MAIL_TO
        mail.CC = MAIL_CC
        mail.Subject = MAIL_SUBJECT
        mail.Body = body
        mail.Attachments.Add(zip_path)
        mail.Send()
        logger.info(f"이메일 발송 완료. To: {MAIL_TO}, Subject: {MAIL_SUBJECT}")
    except Exception as e:
        logger.error(f"Outlook 이메일 발송 오류: {e}. Outlook이 실행 중인지 확인하세요.")
        raise
