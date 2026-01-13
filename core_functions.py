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
from config import DATA_FILE, GOOGLE_NEWS_URL, MAX_NEWS_ENTRIES, OPENAI_MODEL, MAX_TOKENS, MAIL_TO, MAIL_CC, MAIL_SUBJECT, PPT_TITLE, RISK_KEYWORDS
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

def analyze_risk(title):
    """뉴스 제목을 기반으로 리스크 등급(RED, AMBER, GREEN)을 판별합니다."""
    title_norm = title.replace(" ", "")
    for keyword in RISK_KEYWORDS["RED"]:
        if keyword in title_norm:
            return "RED"
    for keyword in RISK_KEYWORDS["AMBER"]:
        if keyword in title_norm:
            return "AMBER"
    return "GREEN"

def generate_dashboard(news_data, summary_map):
    """수집된 데이터로 index.html 대시보드를 생성합니다."""
    json_data = json.dumps(news_data, ensure_ascii=False)
    
    summary_html = ""
    for keyword, text in summary_map.items():
        summary_html += f"<strong>• {keyword}:</strong> {text}<br>"

    # HTML 템플릿 (축약된 형태가 아닌 전체 포함)
    html_content = f"""
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>보안/안전 뉴스 분석 대시보드</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://unpkg.com/@phosphor-icons/web"></script>
    <style>
        body {{ font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif; background-color: #f3f4f6; }}
        .card {{ background: white; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); transition: transform 0.2s; }}
        .card:hover {{ transform: translateY(-2px); }}
        .risk-badge-RED {{ background-color: #fee2e2; color: #991b1b; border: 1px solid #fecaca; }}
        .risk-badge-AMBER {{ background-color: #ffedd5; color: #9a3412; border: 1px solid #fed7aa; }}
        .risk-badge-GREEN {{ background-color: #dcfce7; color: #166534; border: 1px solid #bbf7d0; }}
    </style>
</head>
<body class="text-gray-800">
    <nav class="bg-slate-900 text-white p-4 sticky top-0 z-50 shadow-lg">
        <div class="container mx-auto flex justify-between items-center">
            <div class="flex items-center gap-3">
                <i class="ph ph-shield-check text-3xl text-blue-400"></i>
                <div>
                    <h1 class="text-xl font-bold">Security Analysis Dashboard</h1>
                    <p class="text-xs text-slate-400">보안/안전 뉴스 모니터링 시스템</p>
                </div>
            </div>
            <div class="hidden md:flex gap-4 text-sm">
                <span class="px-3 py-1 bg-slate-800 rounded-full">데이터 업데이트: {datetime.now().strftime('%Y-%m-%d %H:%M')}</span>
            </div>
        </div>
    </nav>
    <div class="container mx-auto p-4 max-w-7xl">
        <div class="bg-gradient-to-r from-blue-900 to-indigo-900 rounded-2xl p-6 text-white mb-8 shadow-xl">
            <div class="flex items-start gap-4">
                <div class="p-3 bg-white/10 rounded-lg">
                    <i class="ph ph-robot text-3xl text-yellow-300"></i>
                </div>
                <div>
                    <h2 class="text-lg font-bold mb-2 flex items-center gap-2">AI 임원 요약 리포트 <span class="text-xs font-normal bg-blue-600 px-2 py-0.5 rounded">Auto-Generated</span></h2>
                    <p class="text-blue-100 leading-relaxed text-sm md:text-base">{summary_html}</p>
                </div>
            </div>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-4 gap-4 mb-8">
            <div class="card p-5 border-l-4 border-blue-500">
                <p class="text-gray-500 text-sm font-medium">총 분석 기사</p>
                <p class="text-3xl font-bold mt-1" id="total-count">-</p>
            </div>
            <div class="card p-5 border-l-4 border-red-500">
                <p class="text-gray-500 text-sm font-medium">위기(Critical) 감지</p>
                <p class="text-3xl font-bold mt-1 text-red-600" id="critical-count">-</p>
            </div>
            <div class="card p-5 border-l-4 border-yellow-500">
                <p class="text-gray-500 text-sm font-medium">주의(Warning) 감지</p>
                <p class="text-3xl font-bold mt-1 text-yellow-600" id="warning-count">-</p>
            </div>
            <div class="card p-5 border-l-4 border-green-500">
                <p class="text-gray-500 text-sm font-medium">최다 언급 키워드</p>
                <p class="text-2xl font-bold mt-1 text-green-700 truncate" id="top-keyword">-</p>
            </div>
        </div>
        <div class="grid grid-cols-1 lg:grid-cols-3 gap-6 mb-8">
            <div class="card p-6 lg:col-span-2">
                <h3 class="font-bold text-gray-700 mb-4 flex items-center gap-2"><i class="ph ph-trend-up"></i> 일별 뉴스 트렌드</h3>
                <canvas id="trendChart" height="250"></canvas>
            </div>
            <div class="card p-6">
                <h3 class="font-bold text-gray-700 mb-4 flex items-center gap-2"><i class="ph ph-chart-pie-slice"></i> 리스크 분포</h3>
                <div class="relative h-64"><canvas id="riskChart"></canvas></div>
            </div>
        </div>
        <div class="flex flex-col md:flex-row gap-6">
            <div class="w-full md:w-64 shrink-0 space-y-4">
                <div class="card p-5 sticky top-24">
                    <h3 class="font-bold text-gray-700 mb-4 border-b pb-2">필터링 옵션</h3>
                    <div class="mb-4">
                        <label class="block text-sm font-medium text-gray-600 mb-1">키워드 선택</label>
                        <select id="keyword-filter" class="w-full p-2 border rounded-lg bg-gray-50 focus:ring-2 focus:ring-blue-500 outline-none"><option value="all">전체 보기</option></select>
                    </div>
                    <div class="mb-4">
                        <label class="block text-sm font-medium text-gray-600 mb-1">리스크 레벨</label>
                        <select id="risk-filter" class="w-full p-2 border rounded-lg bg-gray-50 focus:ring-2 focus:ring-blue-500 outline-none">
                            <option value="all">전체 등급</option>
                            <option value="RED">🚨 위기 (RED)</option>
                            <option value="AMBER">⚠️ 주의 (AMBER)</option>
                            <option value="GREEN">✅ 양호 (GREEN)</option>
                        </select>
                    </div>
                    <div>
                        <label class="block text-sm font-medium text-gray-600 mb-1">검색어</label>
                        <input type="text" id="search-input" placeholder="제목 검색..." class="w-full p-2 border rounded-lg bg-gray-50 focus:ring-2 focus:ring-blue-500 outline-none">
                    </div>
                </div>
            </div>
            <div class="flex-1">
                <div class="flex justify-between items-center mb-4">
                    <h3 class="font-bold text-xl text-gray-800">상세 뉴스 리스트</h3>
                    <span id="filtered-count" class="text-sm text-gray-500">Total: 0건</span>
                </div>
                <div id="news-container" class="space-y-3"></div>
            </div>
        </div>
    </div>
    <script>
        const rawData = {json_data};
        let currentData = [...rawData];
        const container = document.getElementById('news-container');
        const totalCountEl = document.getElementById('total-count');
        const criticalCountEl = document.getElementById('critical-count');
        const warningCountEl = document.getElementById('warning-count');
        const topKeywordEl = document.getElementById('top-keyword');
        const filteredCountEl = document.getElementById('filtered-count');
        const keywordFilter = document.getElementById('keyword-filter');
        const riskFilter = document.getElementById('risk-filter');
        const searchInput = document.getElementById('search-input');
        function initFilters() {{
            const keywords = [...new Set(rawData.map(item => item.keyword))].filter(k => k);
            keywords.forEach(k => {{
                const option = document.createElement('option');
                option.value = k;
                option.textContent = k.toUpperCase();
                keywordFilter.appendChild(option);
            }});
        }}
        function renderKPIs(data) {{
            totalCountEl.textContent = data.length.toLocaleString();
            criticalCountEl.textContent = data.filter(i => i.risk === 'RED').length.toLocaleString();
            warningCountEl.textContent = data.filter(i => i.risk === 'AMBER').length.toLocaleString();
            if(data.length === 0) {{ topKeywordEl.textContent = "-"; return; }}
            const counts = {{}};
            data.forEach(x => {{ counts[x.keyword] = (counts[x.keyword] || 0) + 1; }});
            const top = Object.keys(counts).reduce((a, b) => counts[a] > counts[b] ? a : b);
            topKeywordEl.textContent = top.toUpperCase();
        }}
        function renderList(data) {{
            container.innerHTML = '';
            filteredCountEl.textContent = `Total: ${{data.length}}건`;
            if (data.length === 0) {{ container.innerHTML = '<div class="p-8 text-center text-gray-400">검색 결과가 없습니다.</div>'; return; }}
            data.forEach(item => {{
                const el = document.createElement('div');
                el.className = 'card p-4 hover:bg-gray-50 cursor-pointer group';
                el.onclick = () => window.open(item.link, '_blank');
                const riskClass = `risk-badge-${{item.risk}}` || 'risk-badge-GREEN';
                el.innerHTML = `
                    <div class="flex justify-between items-start gap-4">
                        <div class="flex-1">
                            <div class="flex items-center gap-2 mb-1">
                                <span class="text-xs font-bold px-2 py-0.5 rounded border uppercase ${{riskClass}}">${{item.risk}}</span>
                                <span class="text-xs font-semibold text-blue-600 bg-blue-50 px-2 py-0.5 rounded border border-blue-100">${{item.keyword.toUpperCase()}}</span>
                                <span class="text-xs text-gray-400">${{item.date}}</span>
                            </div>
                            <h4 class="font-bold text-gray-800 group-hover:text-blue-600 transition-colors leading-snug">${{item.title}}</h4>
                        </div>
                        <i class="ph ph-arrow-square-out text-gray-300 group-hover:text-blue-500"></i>
                    </div>
                `;
                container.appendChild(el);
            }});
        }}
        let trendChartInstance = null;
        let riskChartInstance = null;
        function renderCharts(data) {{
            const dateCounts = {{}};
            const riskCounts = {{ 'RED': 0, 'AMBER': 0, 'GREEN': 0 }};
            data.forEach(item => {{
                const d = item.date;
                dateCounts[d] = (dateCounts[d] || 0) + 1;
                if (riskCounts[item.risk] !== undefined) {{ riskCounts[item.risk]++; }} else {{ riskCounts['GREEN']++; }}
            }});
            const sortedDates = Object.keys(dateCounts).sort();
            const trendData = sortedDates.map(d => dateCounts[d]);
            const ctxTrend = document.getElementById('trendChart').getContext('2d');
            if (trendChartInstance) trendChartInstance.destroy();
            trendChartInstance = new Chart(ctxTrend, {{
                type: 'line',
                data: {{ labels: sortedDates, datasets: [{{ label: '일별 기사량', data: trendData, borderColor: '#3b82f6', backgroundColor: 'rgba(59, 130, 246, 0.1)', tension: 0.3, fill: true }}] }},
                options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ display: false }} }}, scales: {{ y: {{ beginAtZero: true, grid: {{ display: false }} }}, x: {{ grid: {{ display: false }} }} }} }}
            }});
            const ctxRisk = document.getElementById('riskChart').getContext('2d');
            if (riskChartInstance) riskChartInstance.destroy();
            riskChartInstance = new Chart(ctxRisk, {{
                type: 'doughnut',
                data: {{ labels: ['위기 (Red)', '주의 (Amber)', '양호 (Green)'], datasets: [{{ data: [riskCounts['RED'], riskCounts['AMBER'], riskCounts['GREEN']], backgroundColor: ['#ef4444', '#f59e0b', '#22c55e'], borderWidth: 0 }}] }},
                options: {{ responsive: true, maintainAspectRatio: false, cutout: '70%', plugins: {{ legend: {{ position: 'right' }} }} }}
            }});
        }}
        function filterData() {{
            const keyVal = keywordFilter.value;
            const riskVal = riskFilter.value;
            const searchVal = searchInput.value.toLowerCase();
            const filtered = rawData.filter(item => {{
                const matchKey = keyVal === 'all' || item.keyword === keyVal;
                const matchRisk = riskVal === 'all' || item.risk === riskVal;
                const matchSearch = item.title.toLowerCase().includes(searchVal);
                return matchKey && matchRisk && matchSearch;
            }});
            currentData = filtered;
            renderKPIs(currentData);
            renderList(currentData);
            renderCharts(currentData);
        }}
        keywordFilter.addEventListener('change', filterData);
        riskFilter.addEventListener('change', filterData);
        searchInput.addEventListener('input', filterData);
        initFilters();
        filterData();
    </script>
</body>
</html>
"""
    output_path = os.path.join(os.getcwd(), "index.html")
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"Dashboard generated: {output_path}")

def execute():
    """전체 프로세스 실행: 크롤링 -> 요약 -> PPT 생성 -> 이메일 전송 -> 웹 대시보드 갱신"""
    keywords = load_keywords()
    if not keywords:
        print("Keywords list is empty.")
        return

    summary_map = {}
    all_news_data = [] # 대시보드용 전체 데이터
    
    for keyword in keywords:
        print(f"Processing: {keyword}...")
        news_items = crawl_news(keyword)
        
        # 1. 뉴스 데이터 수집 및 리스크 분석
        for item in news_items:
            risk = analyze_risk(item['title'])
            # 날짜 포맷팅
            try:
                # RSS feed published example: "Mon, 06 Jan 2025 10:00:00 GMT"
                dt = datetime.strptime(item['published'], "%a, %d %b %Y %H:%M:%S %Z")
                date_fmt = dt.strftime("%Y-%m-%d")
            except:
                date_fmt = datetime.now().strftime("%Y-%m-%d")

            all_news_data.append({
                "keyword": keyword,
                "title": item['title'],
                "link": item['link'],
                "date": date_fmt,
                "risk": risk
            })

        # 2. AI 요약
        summary = summarize_news(keyword, news_items)
        summary_map[keyword] = summary
        time.sleep(1) # 부하 조절

    # 3. PPT 생성
    ppt_path = os.path.join(os.getcwd(), f"security_report_{datetime.now().strftime('%Y%m%d')}.pptx")
    make_ppt(summary_map, ppt_path)

    # 4. 웹 대시보드(index.html) 생성
    generate_dashboard(all_news_data, summary_map)
    
    # 5. 이메일 전송
    send_email(ppt_path)
