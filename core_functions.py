import sys
import json
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.header import Header
from email import encoders
from datetime import datetime
import time

# 한글 출력을 위한 인코딩 설정
sys.stdout.reconfigure(encoding='utf-8')

import requests
import feedparser
import openai

from dotenv import load_dotenv
from config import DATA_FILE, GOOGLE_NEWS_URL, MAX_NEWS_ENTRIES, OPENAI_MODEL, MAX_TOKENS, MAIL_TO, MAIL_CC, MAIL_SUBJECT, PPT_TITLE
from data_manager import load_keywords
from report_generator import make_ppt

def setup_environment():
    """환경 변수(.env)를 로드합니다."""
    # GitHub Actions에서는 Secrets가 자동으로 주입되므로 .env가 없어도 괜찮습니다.
    load_dotenv(override=True) 

def load_keywords():
    """JSON 파일에서 키워드 목록을 불러옵니다."""
    if not os.path.exists(DATA_FILE):
        return []
    try:
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (json.JSONDecodeError, IOError):
        return []

def save_keywords(keywords):
    """키워드 목록을 JSON 파일에 저장합니다."""
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(keywords, f, ensure_ascii=False, indent=4)

import urllib.parse

def crawl_news(keyword):
    """구글 뉴스 RSS를 크롤링하여 뉴스 목록을 반환합니다."""
    encoded_keyword = urllib.parse.quote(keyword)
    url = GOOGLE_NEWS_URL.format(q=encoded_keyword)
    try:
        feed = feedparser.parse(url)
        news_items = []
        for entry in feed.entries[:MAX_NEWS_ENTRIES]:
            news_items.append({
                'title': entry.title,
                'link': entry.link,
                'published': entry.published,
                'summary': entry.description
            })
        return news_items
    except Exception as e:
        print(f"Error crawling news for {keyword}: {e}")
        return []

def summarize_news(keyword, news_items):
    """OpenAI를 사용하여 뉴스 전체를 요약합니다."""
    if not news_items:
        return "뉴스 없음"

    openai_api_key = os.getenv("OPENAI_API_KEY")
    if not openai_api_key:
        print("OPENAI_API_KEY not found. Skipping AI summary.")
        return "AI 요약 실패 (API Key 없음)"

    client = openai.OpenAI(api_key=openai_api_key)

    # 뉴스 제목들만 모아서 요약 요청
    titles = "\n".join([f"- {item['title']}" for item in news_items[:10]]) # 상위 10개만 요약
    prompt = f"다음은 '{keyword}' 관련 주요 뉴스 제목들입니다. 이를 바탕으로 보안/안전 관점에서 핵심 내용을 3~5줄로 요약해줘:\n{titles}"

    try:
        response = client.chat.completions.create(
            model=OPENAI_MODEL,
            messages=[
                {"role": "system", "content": "You are a helpful assistant for security news summarization."},
                {"role": "user", "content": prompt}
            ],
            max_tokens=MAX_TOKENS
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        print(f"OpenAI summary error: {e}")
        return "AI 요약 중 오류 발생"

def send_email(file_path):
    """생성된 PPT 파일을 이메일로 전송합니다."""
    smtp_user = os.getenv("SMTP_USER")
    smtp_password = os.getenv("SMTP_PASSWORD")
    
    if not smtp_user or not smtp_password:
        print("SMTP Credentials not found. Skipping email.")
        return

    msg = MIMEMultipart()
    msg['From'] = smtp_user
    msg['To'] = MAIL_TO
    msg['Cc'] = MAIL_CC
    msg['Subject'] = Header(MAIL_SUBJECT, 'utf-8')

    body = f"안녕하세요,\n\n{datetime.now().date()} 보안·안전 자동 보고서입니다.\n첨부파일을 확인해주세요."
    msg.attach(MIMEText(body, 'plain', 'utf-8'))

    try:
        with open(file_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
        
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {os.path.basename(file_path)}",
        )
        msg.attach(part)

        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(smtp_user, smtp_password)
        text = msg.as_string()
        
        # CC 포함 전송
        recipients = [MAIL_TO]
        if MAIL_CC:
            recipients.append(MAIL_CC)
            
        server.sendmail(smtp_user, recipients, text)
        server.quit()
        print(f"Email sent successfully to {recipients}")
    except Exception as e:
        print(f"Failed to send email: {e}")

def execute():
    """전체 프로세스 실행: 크롤링 -> 요약 -> PPT 생성 -> 이메일 전송"""
    keywords = load_keywords()
    if not keywords:
        print("Keywords list is empty.")
        return

    summary_map = {}
    
    for keyword in keywords:
        print(f"Processing: {keyword}...")
        news_items = crawl_news(keyword)
        summary = summarize_news(keyword, news_items)
        summary_map[keyword] = summary
        time.sleep(1) # 부하 조절

    ppt_path = os.path.join(os.getcwd(), f"security_report_{datetime.now().strftime('%Y%m%d')}.pptx")
    make_ppt(summary_map, ppt_path)
    
    # 이메일 전송 (환경 변수 확인)
    send_email(ppt_path)
