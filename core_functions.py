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
    """수집된 데이터로 index.html 대시보드를 생성합니다. (v2.0: One-Page PPT Style, Smart Filter, PDF Export)"""
    json_data = json.dumps(news_data, ensure_ascii=False)
    
    summary_html = ""
    for keyword, text in summary_map.items():
        summary_html += f"<div class='mb-2'><span class='font-bold text-blue-700'>• {keyword}</span>: <span class='text-gray-700'>{text}</span></div>"

    # HTML 템플릿 (v2.0 Professional PPT Style)
    html_content = f"""
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Security Insight Pro</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://unpkg.com/@phosphor-icons/web"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Pretendard:wght@300;400;500;600;700&display=swap');
        body {{ font-family: 'Pretendard', sans-serif; background-color: #eef2f6; -webkit-print-color-adjust: exact; }}
        .a4-page {{ 
            width: 100%; 
            max-width: 210mm; 
            margin: 0 auto; 
            background: white; 
            min-height: 297mm; 
            box-shadow: 0 10px 30px rgba(0,0,0,0.1); 
            padding: 10mm;
            position: relative;
        }}
        .btn-filter {{ transition: all 0.2s; }}
        .btn-filter.active {{ background-color: #1e3a8a; color: white; border-color: #1e3a8a; }}
        .btn-filter:hover:not(.active) {{ background-color: #eff6ff; }}
        
        /* PDF Generation Overrides */
        /* PDF Generation Overrides */
        body.generating-pdf .no-print {{ display: none !important; }}
        body.generating-pdf .a4-page {{ 
            margin: 0 !important; 
            padding: 10mm !important;
            width: 210mm !important; 
            height: 297mm !important; /* Force A4 Height */
            max-height: 297mm !important;
            overflow: hidden !important; /* Crop overflow */
            box-shadow: none !important; 
            border: none !important; 
        }}
        /* Hide lower part of news list if it overflows */
        body.generating-pdf #news-container {{
            max-height: 110mm !important;
            overflow: hidden !important;
        }}
        body.generating-pdf * {{ transform: none !important; transition: none !important; box-shadow: none !important; }}
        body.generating-pdf ::-webkit-scrollbar {{ display: none; }}
        
        /* Scrollbar mostly hidden for clean look */
        ::-webkit-scrollbar {{ width: 6px; }}
        ::-webkit-scrollbar-track {{ background: transparent; }}
        ::-webkit-scrollbar-thumb {{ background: #cbd5e1; border-radius: 3px; }}
    </style>
</head>
<body class="py-8">
    
    <!-- Control Bar -->
    <div class="fixed top-4 right-4 z-50 flex gap-2 no-print">
        <button onclick="triggerUpdate()" class="bg-slate-800 text-white px-4 py-2 rounded-lg shadow-lg hover:bg-slate-700 flex items-center gap-2 font-medium border border-slate-600 transition-colors">
            <i class="ph ph-arrows-clockwise"></i> Update Data
        </button>
        <button onclick="downloadPDF()" class="bg-blue-900 text-white px-4 py-2 rounded-lg shadow-lg hover:bg-blue-800 flex items-center gap-2 font-medium transition-colors">
            <i class="ph ph-download-simple"></i> Download PDF
        </button>
    </div>

    <!-- Main Dashboard Area (Target for PDF) -->
    <div id="dashboard-content" class="a4-page rounded-xl">
        
        <!-- Header -->
        <header class="flex justify-between items-center border-b-2 border-slate-800 pb-4 mb-6">
            <div class="flex items-center gap-3">
                <div class="bg-slate-900 text-white p-2 rounded-lg">
                    <i class="ph ph-shield-check text-3xl"></i>
                </div>
                <div>
                    <h1 class="text-2xl font-bold text-slate-900 tracking-tight">SECURITY INSIGHT PRO</h1>
                    <p class="text-xs text-slate-500 font-medium tracking-wide">INTELLIGENCE DASHBOARD | {datetime.now().strftime('%Y-%m-%d')}</p>
                </div>
            </div>
            <div class="text-right">
                <div class="flex items-center gap-2 justify-end">
                    <span class="w-2 h-2 rounded-full bg-green-500 animate-pulse"></span>
                    <span class="text-xs font-bold text-green-700 uppercase">Live System</span>
                </div>
                <p class="text-[10px] text-gray-400 mt-1">CONFIDENTIAL REPORT</p>
            </div>
        </header>

        <!-- Executive Summary -->
        <section class="mb-6">
            <div class="bg-slate-50 rounded-lg p-5 border border-slate-200">
                <h3 class="text-sm font-bold text-slate-700 mb-3 flex items-center gap-2">
                    <i class="ph ph-robot text-blue-600"></i> AI Executive Briefing
                    <span class="text-[10px] bg-blue-100 text-blue-700 px-2 py-0.5 rounded-full">Analysis Read</span>
                </h3>
                <div class="text-xs md:text-sm text-slate-700 leading-relaxed">
                    {summary_html}
                </div>
            </div>
        </section>

        <!-- KPIs -->
        <section class="grid grid-cols-4 gap-4 mb-6">
            <div class="p-4 rounded-lg border border-blue-100 bg-blue-50/50">
                <p class="text-xs text-blue-500 font-bold uppercase mb-1">Total News</p>
                <p class="text-3xl font-bold text-slate-800" id="total-count">-</p>
            </div>
            <div class="p-4 rounded-lg border border-red-100 bg-red-50/50">
                <p class="text-xs text-red-500 font-bold uppercase mb-1">Critical Risk</p>
                <p class="text-3xl font-bold text-red-700" id="critical-count">-</p>
            </div>
            <div class="p-4 rounded-lg border border-amber-100 bg-amber-50/50">
                <p class="text-xs text-amber-500 font-bold uppercase mb-1">Warning</p>
                <p class="text-3xl font-bold text-amber-700" id="warning-count">-</p>
            </div>
            <div class="p-4 rounded-lg border border-green-100 bg-green-50/50">
                <p class="text-xs text-green-600 font-bold uppercase mb-1">Top Keyword</p>
                <p class="text-xl font-bold text-green-800 truncate" id="top-keyword">-</p>
            </div>
        </section>

        <!-- Charts Area -->
        <section class="grid grid-cols-3 gap-6 mb-6">
            <div class="col-span-2 border border-gray-100 rounded-xl p-4 shadow-sm bg-white">
                <h3 class="text-xs font-bold text-gray-500 uppercase mb-4">Daily Trend Analysis</h3>
                <div class="relative h-48 w-full">
                    <canvas id="trendChart"></canvas>
                </div>
            </div>
            <div class="border border-gray-100 rounded-xl p-4 shadow-sm bg-white">
                <h3 class="text-xs font-bold text-gray-500 uppercase mb-4">Risk Distribution</h3>
                <div class="relative h-48 w-full flex justify-center">
                    <canvas id="riskChart"></canvas>
                </div>
            </div>
        </section>

        <!-- Smart Filter (Buttons) -->
        <section class="mb-4 no-print">
            <div class="flex flex-wrap gap-2 items-center bg-gray-50 p-3 rounded-lg border border-gray-200">
                <div class="flex items-center gap-2 mr-4">
                    <i class="ph ph-funnel text-gray-400"></i>
                    <span class="text-xs font-bold text-gray-500">SMART FILTER</span>
                </div>
                
                <!-- Date Buttons -->
                <div class="flex gap-1" id="date-filters">
                    <button class="btn-filter px-3 py-1 text-xs rounded-md border border-gray-200 bg-white text-gray-600 font-medium" data-val="today">Today</button>
                    <button class="btn-filter px-3 py-1 text-xs rounded-md border border-gray-200 bg-white text-gray-600 font-medium" data-val="3d">3 Days</button>
                    <button class="btn-filter active px-3 py-1 text-xs rounded-md border border-gray-200 bg-white text-gray-600 font-medium" data-val="7d">7 Days</button>
                    <button class="btn-filter px-3 py-1 text-xs rounded-md border border-gray-200 bg-white text-gray-600 font-medium" data-val="all">All Time</button>
                </div>

                <div class="h-4 w-px bg-gray-300 mx-2"></div>

                <!-- Risk Buttons -->
                <div class="flex gap-1" id="risk-filters">
                    <button class="btn-filter active px-3 py-1 text-xs rounded-md border border-gray-200 bg-white text-gray-600 font-medium" data-val="all">All Risks</button>
                    <button class="btn-filter px-3 py-1 text-xs rounded-md border border-gray-200 bg-white text-gray-600 font-medium" data-val="RED">Critical</button>
                    <button class="btn-filter px-3 py-1 text-xs rounded-md border border-gray-200 bg-white text-gray-600 font-medium" data-val="AMBER">Warning</button>
                </div>
                
                <div class="h-4 w-px bg-gray-300 mx-2"></div>

                <!-- Keyword Buttons (Dynamic) -->
                <div class="flex gap-1 flex-wrap" id="keyword-filters">
                    <button class="btn-filter active px-3 py-1 text-xs rounded-md border border-gray-200 bg-white text-gray-600 font-medium" data-val="all">All Topics</button>
                </div>
            </div>
        </section>

        <!-- Data List -->
        <section>
            <div class="flex justify-between items-end mb-3">
                <h3 class="text-sm font-bold text-slate-800">Intelligence Feed</h3>
                <span class="text-xs text-slate-400" id="filtered-count">Total: 0</span>
            </div>
            <div id="news-container" class="grid grid-cols-2 gap-3">
                <!-- Cards will be injected here -->
            </div>
        </section>

        <!-- Footer -->
        <footer class="mt-8 pt-4 border-t border-gray-100 flex justify-between items-center text-[10px] text-gray-400">
            <p>Generated by Security Report Automation System</p>
            <p>CONFIDENTIAL | INTERNAL USE ONLY</p>
        </footer>
    </div>

    <script>
        const rawData = {json_data};
        let currentData = [...rawData];

        // Chart Instances
        let trendChart = null;
        let riskChart = null;

        // Current Filter State
        let state = {{ risk: 'all', keyword: 'all', date: '7d' }};

        function init() {{
            initFilters();
            renderAll();
        }}

        function initFilters() {{
            // Render Keyword Buttons
            const keywords = [...new Set(rawData.map(d => d.keyword))].filter(k => k);
            const container = document.getElementById('keyword-filters');
            keywords.forEach(k => {{
                const btn = document.createElement('button');
                btn.className = 'btn-filter px-3 py-1 text-xs rounded-md border border-gray-200 bg-white text-gray-600 font-medium';
                btn.textContent = k;
                btn.dataset.val = k;
                btn.onclick = () => setFilter('keyword', k);
                container.appendChild(btn);
            }});

            // Attach Risk Button Events
            document.querySelectorAll('#risk-filters button').forEach(btn => {{
                btn.onclick = () => setFilter('risk', btn.dataset.val);
            }});

            // Attach Date Button Events
            document.querySelectorAll('#date-filters button').forEach(btn => {{
                btn.onclick = () => setFilter('date', btn.dataset.val);
            }});
        }}

        function setFilter(type, val) {{
            state[type] = val;
            
            // Update UI Active State
            let selector = `#${{type}}-filters button`;
            
            document.querySelectorAll(selector).forEach(b => {{
                if(b.dataset.val === val) b.classList.add('active');
                else b.classList.remove('active');
            }});

            filterData();
        }}

        function filterData() {{
            const today = new Date();
            today.setHours(0,0,0,0);

            currentData = rawData.filter(item => {{
                // Keyword & Risk Filter
                const matchKey = state.keyword === 'all' || item.keyword === state.keyword;
                const matchRisk = state.risk === 'all' || item.risk === state.risk;

                // Date Filter
                let matchDate = true;
                if (state.date !== 'all') {{
                    const itemDate = new Date(item.date);
                    itemDate.setHours(0,0,0,0);
                    const diffTime = Math.abs(today - itemDate);
                    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)); 
                    
                    if (state.date === 'today') matchDate = (diffDays === 0);
                    else if (state.date === '3d') matchDate = (diffDays <= 3);
                    else if (state.date === '7d') matchDate = (diffDays <= 7);
                }}

                return matchKey && matchRisk && matchDate;
            }});
            renderAll();
        }}

        function renderAll() {{
            renderKPIs();
            renderCharts();
            renderList();
        }}

        function renderKPIs() {{
            document.getElementById('total-count').innerText = currentData.length;
            document.getElementById('critical-count').innerText = currentData.filter(i => i.risk === 'RED').length;
            document.getElementById('warning-count').innerText = currentData.filter(i => i.risk === 'AMBER').length;
            
            if (currentData.length > 0) {{
                const counts = {{}};
                currentData.forEach(x => counts[x.keyword] = (counts[x.keyword] || 0) + 1);
                const top = Object.keys(counts).reduce((a, b) => counts[a] > counts[b] ? a : b);
                document.getElementById('top-keyword').innerText = top;
            }} else {{
                document.getElementById('top-keyword').innerText = "-";
            }}
            document.getElementById('filtered-count').innerText = `Total: ${{currentData.length}}`;
        }}

        function renderCharts() {{
            // Trend
            const dates = {{}};
            currentData.forEach(i => dates[i.date] = (dates[i.date]||0)+1);
            const sortedDates = Object.keys(dates).sort();
            
            const ctxTrend = document.getElementById('trendChart');
            if(trendChart) trendChart.destroy();
            trendChart = new Chart(ctxTrend, {{
                type: 'line',
                data: {{
                    labels: sortedDates,
                    datasets: [{{
                        data: sortedDates.map(d => dates[d]),
                        borderColor: '#1e3a8a',
                        backgroundColor: 'rgba(30, 58, 138, 0.05)',
                        borderWidth: 2,
                        tension: 0.4,
                        pointRadius: 2,
                        fill: true
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{ legend: {{ display: false }} }},
                    scales: {{ 
                        x: {{ grid: {{ display: false }}, ticks: {{ font: {{ size: 9 }} }} }},
                        y: {{ beginAtZero: true, grid: {{ display: false }}, ticks: {{ display: false }} }}
                    }}
                }}
            }});

            // Risk
            const riskCounts = {{ RED:0, AMBER:0, GREEN:0 }};
            currentData.forEach(i => riskCounts[i.risk] = (riskCounts[i.risk]||0)+1);
            
            const ctxRisk = document.getElementById('riskChart');
            if(riskChart) riskChart.destroy();
            riskChart = new Chart(ctxRisk, {{
                type: 'doughnut',
                data: {{
                    labels: ['Crit', 'Warn', 'Safe'],
                    datasets: [{{
                        data: [riskCounts.RED, riskCounts.AMBER, riskCounts.GREEN],
                        backgroundColor: ['#ef4444', '#f59e0b', '#10b981'],
                        borderWidth: 0,
                        hoverOffset: 4
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    cutout: '75%',
                    plugins: {{ 
                        legend: {{ position: 'right', labels: {{ boxWidth: 10, font: {{ size: 10 }} }} }} 
                    }}
                }}
            }});
        }}

        function renderList() {{
            const div = document.getElementById('news-container');
            div.innerHTML = '';
            
            // Show only top 8 items to fit on page (or scrollable if needed, but intended for single page)
            // For PDF purposes, we allow it to grow, but visually we want compact.
            // Let's list all but styled compactly.
            
            if (currentData.length === 0) {{
                div.innerHTML = '<div class="col-span-2 text-center text-gray-400 py-8 text-sm">No data available</div>';
                return;
            }}

            currentData.forEach(item => {{
                const el = document.createElement('div');
                el.className = 'border border-gray-100 rounded-lg p-3 bg-white hover:border-blue-300 transition-colors cursor-pointer group shadow-sm flex flex-col justify-between h-full';
                el.onclick = () => window.open(item.link, '_blank');
                
                const badgeColor = item.risk === 'RED' ? 'bg-red-50 text-red-700 border-red-100' :
                                   item.risk === 'AMBER' ? 'bg-amber-50 text-amber-700 border-amber-100' :
                                   'bg-green-50 text-green-700 border-green-100';

                el.innerHTML = `
                    <div>
                        <div class="flex justify-between items-start mb-1">
                            <span class="text-[10px] font-bold px-1.5 py-0.5 rounded border ${{badgeColor}}">${{item.risk}}</span>
                            <span class="text-[10px] text-gray-400">${{item.date}}</span>
                        </div>
                        <h4 class="text-sm font-bold text-slate-800 leading-snug group-hover:text-blue-700 mb-1 line-clamp-2">${{item.title}}</h4>
                    </div>
                    <div class="flex justify-between items-end mt-2">
                        <span class="text-[10px] font-semibold text-slate-500 bg-slate-100 px-2 py-0.5 rounded">${{item.keyword}}</span>
                    </div>
                `;
                div.appendChild(el);
            }});
        }}

        function downloadPDF() {{
            window.scrollTo(0, 0);
            document.body.classList.add('generating-pdf');
            
            const element = document.getElementById('dashboard-content');
            const opt = {{
                margin:       0,
                filename:     `Security_Report_${{new Date().toISOString().slice(0,10)}}.pdf`,
                image:        {{ type: 'jpeg', quality: 0.98 }},
                html2canvas:  {{ scale: 1.5, useCORS: true, scrollY: 0 }},
                jsPDF:        {{ unit: 'mm', format: 'a4', orientation: 'portrait' }}
            }};
            
            setTimeout(() => {{
                html2pdf().set(opt).from(element).save().then(() => {{
                    document.body.classList.remove('generating-pdf');
                }});
            }}, 500);
        }}

        async function triggerUpdate() {{
            const pwd = prompt("관리자 암호를 입력하세요:");
            if (pwd !== "3867") {{
                if (pwd) alert("암호가 일치하지 않습니다.");
                return;
            }}

            let token = localStorage.getItem("gh_pat");
            if (!token) {{
                token = prompt("최초 1회 설정: GitHub Personal Access Token을 입력해주세요.\\n(입력된 토큰은 브라우저에만 저장되며 서버로 전송되지 않습니다.)");
                if (!token) return;
                localStorage.setItem("gh_pat", token);
            }}

            const repoOwner = "bough38-web";
            const repoName = "security-report-automation";
            const workflowId = "scheduled_report.yml";

            if (!confirm("지금 즉시 크롤링을 시작하시겠습니까?\\n(약 2~3분 소요됩니다)")) return;

            try {{
                const response = await fetch(`https://api.github.com/repos/${{repoOwner}}/${{repoName}}/actions/workflows/${{workflowId}}/dispatches`, {{
                    method: 'POST',
                    headers: {{
                        'Authorization': `Bearer ${{token}}`,
                        'Accept': 'application/vnd.github.v3+json'
                    }},
                    body: JSON.stringify({{ ref: 'main' }})
                }});

                if (response.ok) {{
                    alert("✅ 업데이트 요청 성공!\\n약 3분 뒤에 새로고침 해주세요.");
                }} else {{
                    const err = await response.json();
                    alert(`❌ 요청 실패: ${{err.message}}\\n토큰 권한(workflow)을 확인해주세요.`);
                    // 토큰 문제일 수 있으므로 초기화 기회 제공
                    if(confirm("토큰을 재설정하시겠습니까?")) localStorage.removeItem("gh_pat");
                }}
            }} catch (error) {{
                alert(`❌ 네트워크 오류: ${{error}}\\n인터넷 연결을 확인해주세요.`);
            }}
        }}

        // Run
        init();

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
