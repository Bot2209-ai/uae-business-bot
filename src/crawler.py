import requests
from bs4 import BeautifulSoup
import re
import time
import random
import json
import configparser
from itertools import cycle

config = configparser.ConfigParser()
config.read('config/config.ini')

class BusinessFinder:
    def __init__(self):
        self.proxies = cycle(open(config['PROXY']['proxy_file']).read().splitlines())
        self.emails = set()
        self.status = {'emails': 0, 'last_check': '', 'errors': []}
        
    def get_proxy(self):
        return {'http': next(self.proxies)}
    
    def google_search(self, query):
        try:
            params = {
                'key': config['API']['google_api_key'],
                'cx': config['API']['google_cse_id'],
                'q': f"{query} site:.ae -{config['SEARCH']['exclude_terms']}",
                'num': 10
            }
            response = requests.get(
                'https://www.googleapis.com/customsearch/v1',
                params=params,
                proxies=self.get_proxy(),
                timeout=10
            )
            return [item['link'] for item in response.json().get('items', [])]
        except Exception as e:
            self.status['errors'].append(str(e))
            return []
    
    def extract_emails(self, url):
        try:
            response = requests.get(
                url,
                proxies=self.get_proxy(),
                timeout=10,
                headers={'User-Agent': 'Mozilla/5.0'}
            )
            return re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', response.text)
        except Exception as e:
            self.status['errors'].append(str(e))
            return []
    
    def run(self):
        while True:
            try:
                # 1. Search Google
                query = f"{random.choice(config['SEARCH']['industries'].split(','))} companies in {random.choice(config['SEARCH']['locations'].split(','))}"
                urls = self.google_search(query)
                
                # 2. Crawl Websites
                for url in urls[:int(config['SETTINGS']['max_pages'])]:
                    emails = self.extract_emails(url)
                    self.emails.update(emails)
                    
                    # Update status
                    self.status['emails'] = len(self.emails)
                    self.status['last_check'] = time.ctime()
                    with open('status.json', 'w') as f:
                        json.dump(self.status, f)
                    
                    # Save emails
                    with open(config['SETTINGS']['output_file'], 'w') as f:
                        f.write('\n'.join(self.emails))
                    
                    time.sleep(float(config['SETTINGS']['request_delay']))
                
            except Exception as e:
                self.status['errors'].append(str(e))
                time.sleep(60)

if __name__ == '__main__':
    BusinessFinder().run()
