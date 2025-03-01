# bot.py (Core System)
import asyncio
import aiohttp
import re
import sqlite3
import time
from urllib.parse import urlparse, urljoin
from bs4 import BeautifulSoup
from serpapi import GoogleSearch
from aiosqlite import connect
from flask import Flask, render_template_string

# ========== CONFIGURATION ==========
SERPAPI_KEY = "your_serpapi_key"
MAX_CONCURRENT_TASKS = 3           # Conservative for 1GB RAM
REQUEST_DELAY = (1, 3)            # Random delay range
DB_NAME = 'emails.db'
LOG_FILE = 'bot.log'
# ===================================

app = Flask(__name__)

class OptimizedEmailHunter:
    def __init__(self):
        self.loop = asyncio.get_event_loop()
        self.session = None
        self.keyword_queue = asyncio.Queue()
        self.visited_domains = set()
        self.task_count = 0
        self.start_time = time.time()
        
        # Initialize database
        self.loop.run_until_complete(self.init_db())
        
    async def init_db(self):
        async with connect(DB_NAME) as db:
            await db.execute('''CREATE TABLE IF NOT EXISTS emails
                              (email TEXT PRIMARY KEY, 
                               domain TEXT,
                               source_url TEXT,
                               timestamp REAL)''')
            await db.execute('''CREATE TABLE IF NOT EXISTS stats
                              (key TEXT PRIMARY KEY,
                               value TEXT)''')
            await db.commit()

    async def log_error(self, message):
        async with connect(DB_NAME) as db:
            await db.execute('INSERT INTO stats VALUES (?,?)', 
                           (f'error_{time.time()}', message))
            await db.commit()

    async def store_email(self, email, domain, url):
        try:
            async with connect(DB_NAME) as db:
                await db.execute(
                    "INSERT OR IGNORE INTO emails VALUES (?,?,?,?)",
                    (email, domain, url, time.time())
                )
                await db.commit()
        except Exception as e:
            await self.log_error(f"DB Error: {str(e)}")

    async def fetch(self, url):
        try:
            async with self.session.get(url, timeout=10) as response:
                return await response.text()
        except Exception as e:
            await self.log_error(f"Fetch Error {url}: {str(e)}")
            return None

    async def crawl(self, url):
        domain = urlparse(url).netloc
        if domain in self.visited_domains:
            return
        self.visited_domains.add(domain)
        
        content = await self.fetch(url)
        if not content:
            return
            
        # Extract emails
        emails = re.findall(
            r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b", 
            content
        )
        for email in emails:
            await self.store_email(email, domain, url)
            
        # Extract links
        soup = BeautifulSoup(content, 'html.parser')
        for link in soup.find_all('a', href=True):
            new_url = urljoin(url, link['href'])
            if urlparse(new_url).netloc == domain:
                await asyncio.sleep(1)  # Respectful delay
                self.loop.create_task(self.crawl(new_url))

    async def search(self, query):
        try:
            params = {
                "q": f"{query} site:.ae -platform",
                "api_key": SERPAPI_KEY,
                "num": 20,
                "gl": "ae"
            }
            
            results = GoogleSearch(params).get_dict()
            for result in results.get('organic_results', []):
                url = result.get('link')
                if url:
                    self.loop.create_task(self.crawl(url))
                    
        except Exception as e:
            await self.log_error(f"Search Error: {str(e)}")

    async def run(self):
        self.session = aiohttp.ClientSession()
        
        # Initial search terms
        initial_terms = [
            "UAE food manufacturers",
            "Dubai beverage companies",
            "Abu Dhabi food suppliers"
        ]
        
        for term in initial_terms:
            await self.keyword_queue.put(term)
            
        # Main loop
        while True:
            if self.task_count < MAX_CONCURRENT_TASKS:
                term = await self.keyword_queue.get()
                self.task_count += 1
                self.loop.create_task(self.search(term))
                await asyncio.sleep(5)  # Rate limit searches
            else:
                await asyncio.sleep(1)

@app.route('/')
async def dashboard():
    stats = {}
    async with connect(DB_NAME) as db:
        # Get total emails
        cursor = await db.execute("SELECT COUNT(*) FROM emails")
        stats['total_emails'] = (await cursor.fetchone())[0]
        
        # Get latest error
        cursor = await db.execute(
            "SELECT value FROM stats WHERE key LIKE 'error_%' ORDER BY key DESC LIMIT 1"
        )
        latest_error = (await cursor.fetchone() or [None])[0]
        
    return render_template_string('''
        <h1>Email Hunter Dashboard</h1>
        <p>Total Emails: {{ total_emails }}</p>
        <h3>Latest Error:</h3>
        <pre>{{ latest_error }}</pre>
        <h3>System Status:</h3>
        <p>Running since: {{ start_time | ctime }}</p>
    ''', total_emails=stats['total_emails'], 
        latest_error=latest_error,
        start_time=hunter.start_time)

def run_flask():
    app.run(host='0.0.0.0', port=5000)

if __name__ == "__main__":
    hunter = OptimizedEmailHunter()
    
    # Start Flask in background
    from threading import Thread
    Thread(target=run_flask).start()
    
    # Start main bot loop
    asyncio.get_event_loop().run_until_complete(hunter.run())
