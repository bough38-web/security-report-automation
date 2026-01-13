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
from playwright.sync_api import sync_playwright

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
    """수집된 데이터로 index.html 대시보드를 생성합니다. (v3.0: Left Sidebar Layout)"""
    json_data = json.dumps(news_data, ensure_ascii=False)
    
    summary_html = ""
    for keyword, text in summary_map.items():
        summary_html += f"<div class='mb-2'><span class='font-bold text-blue-700'>• {keyword}</span>: <span class='text-gray-700'>{text}</span></div>"

    # HTML 템플릿 (v3.0 Sidebar Layout)
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
            display: flex; /* Sidebar Layout */
            overflow: hidden;
            position: relative;
        }}

        /* Left Sidebar Styling */
        .sidebar {{
            width: 260px;
            background-color: #0f172a; /* Slate 900 */
            color: white;
            padding: 2rem 1.5rem;
            display: flex;
            flex-direction: column;
            flex-shrink: 0;
        }}

        /* Main Content Styling */
        .main-content {{
            flex: 1;
            padding: 2rem;
            display: flex;
            flex-direction: column;
            background-color: #ffffff;
        }}

        /* Filter Buttons (Vertical) */
        .btn-filter {{ 
            width: 100%; 
            display: flex; 
            align-items: center; 
            justify-content: space-between;
            padding: 0.5rem 0.75rem; 
            margin-bottom: 0.25rem;
            border-radius: 0.5rem; 
            font-size: 0.75rem;
            transition: all 0.2s;
            border: 1px solid #1e293b;
            color: #94a3b8;
            background: rgba(255,255,255,0.05);
        }}
        .btn-filter:hover {{ background: rgba(255,255,255,0.1); color: white; }}
        .btn-filter.active {{ 
            background: #2563eb; 
            color: white; 
            border-color: #2563eb; 
            font-weight: 600; 
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        }}
        
        /* Section Divider in Sidebar */
        .sidebar-divider {{ height: 1px; background: #1e293b; margin: 1.5rem 0; }}
        .sidebar-title {{ font-size: 0.7rem; font-weight: 700; color: #475569; text-transform: uppercase; margin-bottom: 0.5rem; letter-spacing: 0.05em; }}

        /* PDF Generation Overrides */
        body.generating-pdf .no-print {{ display: none !important; }}
        body.generating-pdf .a4-page {{ 
            margin: 0 !important; 
            width: 210mm !important; 
            height: 297mm !important; /* Force A4 Height */
            overflow: hidden !important; 
            box-shadow: none !important; 
            border: none !important; 
        }}
        /* Hide lower part of news list if it overflows */
        body.generating-pdf #news-container {{
            max-height: 100mm !important; /* Adjusted for layout */
            overflow: hidden !important;
        }}
        body.generating-pdf * {{ transform: none !important; transition: none !important; box-shadow: none !important; }}
        body.generating-pdf ::-webkit-scrollbar {{ display: none; }}
    </style>
</head>
<body class="py-8">

    <!-- Password Modal -->
    <div id="pwd-modal" class="fixed inset-0 bg-black/50 z-[60] hidden flex items-center justify-center no-print">
        <div class="bg-white rounded-lg p-6 w-80 shadow-2xl">
            <h3 class="font-bold text-lg mb-4 text-gray-800">Administrator Access</h3>
            <p class="text-sm text-gray-600 mb-2">Enter admin password:</p>
            <input type="password" id="admin-pwd-input" class="w-full p-2 border rounded mb-4 focus:ring-2 focus:ring-blue-500 outline-none" placeholder="Password">
            <div class="flex justify-end gap-2">
                <button onclick="closePwdModal()" class="px-3 py-1 text-sm text-gray-500 hover:bg-gray-100 rounded">Cancel</button>
                <button onclick="submitPwd()" class="px-3 py-1 text-sm bg-blue-600 text-white rounded hover:bg-blue-700">Confirm</button>
            </div>
        </div>
    </div>

    <!-- Main Dashboard Area (Target for PDF) -->
    <div id="dashboard-content" class="a4-page rounded-xl overflow-hidden">
        
        <!-- LEFT SIDEBAR -->
        <aside class="sidebar">
            <!-- Brand -->
            <div class="mb-8">
                <div class="flex items-center gap-2 mb-1">
                    <i class="ph-fill ph-shield-check text-2xl text-blue-400"></i>
                    <h1 class="text-lg font-bold tracking-tight text-white">Insight Pro</h1>
                </div>
                <p class="text-[10px] text-slate-400">Security Intelligence Dashboard</p>
            </div>

            <!-- Controls (Update/PDF) -->
            <div class="mb-6 grid grid-cols-1 gap-2 no-print">
                 <button onclick="triggerUpdate()" class="w-full py-2 bg-slate-800 hover:bg-slate-700 text-white text-xs rounded border border-slate-700 transition-colors flex items-center justify-center gap-2">
                    <i class="ph ph-arrows-clockwise"></i> Update Data
                </button>
                <button onclick="downloadPDF()" class="w-full py-2 bg-blue-600 hover:bg-blue-500 text-white text-xs rounded transition-colors flex items-center justify-center gap-2 font-semibold shadow-lg shadow-blue-900/20">
                    <i class="ph ph-download-simple"></i> Download PDF
                </button>
            </div>
            
            <!-- Filters -->
            <div class="no-print">
                <!-- Date -->
                <div class="mb-4">
                    <h4 class="sidebar-title">Time Range</h4>
                    <div id="date-filters">
                        <button class="btn-filter" data-val="today"><span>Today</span></button>
                        <button class="btn-filter" data-val="3d"><span>3 Days</span></button>
                        <button class="btn-filter active" data-val="7d"><span>7 Days</span></button>
                        <button class="btn-filter" data-val="all"><span>All Time</span></button>
                    </div>
                </div>

                <div class="sidebar-divider"></div>

                <!-- Risk -->
                <div class="mb-4">
                    <h4 class="sidebar-title">Risk Level</h4>
                    <div id="risk-filters">
                        <button class="btn-filter active" data-val="all"><span>All Risks</span> <i class="ph ph-check"></i></button>
                        <button class="btn-filter" data-val="RED"><span class="text-red-400">Critical</span> <i class="ph-fill ph-warning-circle text-red-500"></i></button>
                        <button class="btn-filter" data-val="AMBER"><span class="text-amber-400">Warning</span> <i class="ph-fill ph-warning text-amber-500"></i></button>
                    </div>
                </div>

                <div class="sidebar-divider"></div>

                <!-- Keywords -->
                <div>
                    <h4 class="sidebar-title">Topics</h4>
                    <div id="keyword-filters" class="max-h-[300px] overflow-y-auto pr-1">
                        <button class="btn-filter active" data-val="all"><span>All Topics</span></button>
                        <!-- Dynamic Keywords -->
                    </div>
                </div>
            </div>

            <!-- Footer Info -->
            <div class="mt-auto pt-6 text-[9px] text-slate-500">
                <p>Last Updated:</p>
                <p class="font-mono text-slate-400">{datetime.now().strftime('%Y-%m-%d %H:%M')}</p>
                <p class="mt-2">v3.0.1 Stable</p>
            </div>
        </aside>

        <!-- RIGHT MAIN CONTENT -->
        <main class="main-content">
            
            <!-- Header -->
            <header class="flex justify-between items-end border-b border-gray-100 pb-4 mb-6">
                <div>
                    <h2 class="text-2xl font-bold text-slate-900 leading-none mb-1">Security Executive Report</h2>
                    <p class="text-xs text-gray-500">Daily Intelligence Briefing & Risk Assessment</p>
                </div>
                <div class="flex items-center gap-2">
                    <span class="w-2 h-2 rounded-full bg-green-500 animate-pulse"></span>
                    <span class="text-xs font-bold text-green-700 uppercase">System Live</span>
                </div>
            </header>

            <!-- Executive Summary -->
            <section class="mb-6">
                <div class="bg-slate-50 rounded-xl p-5 border border-slate-200 shadow-sm">
                    <h3 class="text-sm font-bold text-slate-800 mb-3 flex items-center gap-2">
                        <i class="ph-fill ph-robot text-blue-600"></i> AI Executive Briefing
                    </h3>
                    <div class="text-xs md:text-sm text-slate-700 leading-relaxed space-y-2">
                        {summary_html}
                    </div>
                </div>
            </section>

            <!-- KPIs -->
            <section class="grid grid-cols-4 gap-4 mb-6">
                <div class="p-4 rounded-xl border border-gray-100 bg-white shadow-sm flex flex-col justify-between">
                    <p class="text-[10px] text-gray-400 font-bold uppercase tracking-wider">Total Feed</p>
                    <p class="text-3xl font-bold text-slate-800 mt-1" id="total-count">-</p>
                </div>
                <div class="p-4 rounded-xl border-l-4 border-red-500 bg-red-50/30 flex flex-col justify-between">
                    <p class="text-[10px] text-red-500 font-bold uppercase tracking-wider">Critical</p>
                    <p class="text-3xl font-bold text-red-700 mt-1" id="critical-count">-</p>
                </div>
                <div class="p-4 rounded-xl border-l-4 border-amber-500 bg-amber-50/30 flex flex-col justify-between">
                    <p class="text-[10px] text-amber-600 font-bold uppercase tracking-wider">Warning</p>
                    <p class="text-3xl font-bold text-amber-700 mt-1" id="warning-count">-</p>
                </div>
                <div class="p-4 rounded-xl border border-gray-100 bg-white shadow-sm flex flex-col justify-between">
                    <p class="text-[10px] text-green-600 font-bold uppercase tracking-wider">Top Keyword</p>
                    <p class="text-lg font-bold text-green-800 mt-1 truncate" id="top-keyword">-</p>
                </div>
            </section>

            <!-- Charts Area -->
            <section class="grid grid-cols-3 gap-6 mb-6">
                <div class="col-span-2 border border-gray-100 rounded-xl p-4 shadow-sm bg-white">
                    <h3 class="text-[11px] font-bold text-gray-400 uppercase mb-4 tracking-wider">7-Day Incident Trend</h3>
                    <div class="relative h-40 w-full">
                        <canvas id="trendChart"></canvas>
                    </div>
                </div>
                <div class="border border-gray-100 rounded-xl p-4 shadow-sm bg-white">
                    <h3 class="text-[11px] font-bold text-gray-400 uppercase mb-4 tracking-wider">Risk Distribution</h3>
                    <div class="relative h-40 w-full flex justify-center">
                        <canvas id="riskChart"></canvas>
                    </div>
                </div>
            </section>

            <!-- Data List -->
            <section class="flex-1 min-h-0 flex flex-col">
                <div class="flex justify-between items-end mb-3">
                    <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                        <i class="ph-fill ph-list-dashes text-gray-400"></i> Intelligence Feed
                    </h3>
                    <span class="text-xs text-slate-400 font-mono" id="filtered-count">Total: 0</span>
                </div>
                <div id="news-container" class="grid grid-cols-2 gap-3 pb-2">
                    <!-- Cards will be injected here -->
                </div>
            </section>

        </main>
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
                btn.className = 'btn-filter';
                btn.dataset.val = k;
                btn.innerHTML = `<span>${{k}}</span>`; // Simple span for text
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

        let pwdResolve = null;

        function showPwdModal() {{
            return new Promise((resolve) => {{
                pwdResolve = resolve;
                const modal = document.getElementById('pwd-modal');
                const input = document.getElementById('admin-pwd-input');
                modal.classList.remove('hidden');
                input.value = '';
                input.focus();
                
                // Enter key support
                input.onkeyup = (e) => {{ if(e.key === 'Enter') submitPwd(); }};
            }});
        }}

        function closePwdModal() {{
            document.getElementById('pwd-modal').classList.add('hidden');
            if (pwdResolve) pwdResolve(null);
        }}

        function submitPwd() {{
            const val = document.getElementById('admin-pwd-input').value;
            document.getElementById('pwd-modal').classList.add('hidden');
            if (pwdResolve) pwdResolve(val);
        }}

        async function triggerUpdate() {{
            const pwd = await showPwdModal();
            if (pwd !== "3867") {{
                if (pwd !== null) alert("암호가 일치하지 않습니다."); // Only alert if not cancelled
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

            # [Filter] KT텔레캅 Exclusion Rule: Exclude High Risk (RED, AMBER) items
            if keyword == "KT텔레캅" and risk in ["RED", "AMBER"]:
                print(f"  - Skipped (Negative Filter): {item['title']}")
                continue

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

    # 3. 비용이 드는 PPT 생성 대신, 웹 대시보드(index.html) 우선 생성
    generate_dashboard(all_news_data, summary_map)
    html_abs_path = os.path.join(os.getcwd(), "index.html")
    pdf_path = os.path.join(os.getcwd(), f"Security_Report_{datetime.now().strftime('%Y%m%d')}.pdf")

    # 4. Playwright를 이용한 PDF 생성 (Headless Browser)
    try:
        print("Generating PDF from Dashboard via Headless Browser...")
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()
            # 로컬 파일 로드 ('file://' 프로토콜)
            page.goto(f"file://{html_abs_path}")
            
            # 차트 렌더링 대기 (2초)
            page.wait_for_timeout(2000) 
            
            # PDF 스타일 강제 적용 (A4 Single Page)
            page.evaluate("document.body.classList.add('generating-pdf')")
            
            page.pdf(
                path=pdf_path,
                format="A4",
                print_background=True,
                margin={"top": "0mm", "bottom": "0mm", "left": "0mm", "right": "0mm"}
            )
            browser.close()
        print(f"PDF Generated Successfully: {pdf_path}")
    except Exception as e:
        print(f"PDF Generation Failed: {e}")
        pdf_path = None # 실패 시 메일 발송 스킵 혹은 본문만 전송

    # 5. 이메일 전송 (PDF 첨부)
    if pdf_path and os.path.exists(pdf_path):
        send_email(pdf_path)
    else:
        print("PDF 파일이 생성되지 않아 이메일을 발송하지 않습니다.")
