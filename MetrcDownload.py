import os.path

import aiohttp
import asyncio
import json
import time
from openpyxl import Workbook, load_workbook
from concurrent.futures import ThreadPoolExecutor
from bs4 import BeautifulSoup as bs

LIMIT = 3
TIMEOUT = 600  # seconds
USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 ' \
             '(KHTML, like Gecko) Chrome/97.0.4692.99 Safari/537.36'


class Metrc:

    def __init__(self, start_date, end_date, username, password,email, report_data):
        self.gui_queue = None
        self.username = username
        self.password = password
        self.email = email
        self.start_date = start_date
        self.end_date = end_date
        self.report = report_data
        self.license = []

    async def load_login(self):
        url = 'https://co.metrc.com/log-in?ReturnUrl=%2f'
        headers = {
            'Connection': 'keep-alive',
            'Cache-Control': 'max-age=0',
            'sec-ch-ua': '" Not;A Brand";v="99", "Google Chrome";v="97", "Chromium";v="97"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'Upgrade-Insecure-Requests': '1',
            'User-Agent': USER_AGENT,
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;'
                      'q=0.8,application/signed-exchange;v=b3;q=0.9',
            'Sec-Fetch-Site': 'same-origin',
            'Sec-Fetch-Mode': 'navigate',
            'Sec-Fetch-User': '?1',
            'Sec-Fetch-Dest': 'document',
            'Accept-Language': 'en-GB,en-US;q=0.9,en;q=0.8',
        }
        loop_count = 0
        async with self.sema:
            while loop_count < 2:
                try:
                    async with self.session.get(url, headers=headers) as request:
                        response = await request.content.read()
                        content = response.decode('utf-8')
                        html_content = bs(content, "html.parser")
                        self.req_code = html_content.find(attrs={'name': "__RequestVerificationToken"})['value']
                        title = html_content.find('title').text
                        if title == 'Log in | metrc':
                            return True
                except:
                    loop_count += 1
                    await asyncio.sleep(2)
            return False

    async def login(self):
        url = 'https://co.metrc.com/log-in?ReturnUrl=%2f'
        headers = {
            'Connection': 'keep-alive',
            'Cache-Control': 'max-age=0',
            'sec-ch-ua': '" Not;A Brand";v="99", "Google Chrome";v="97", "Chromium";v="97"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'Upgrade-Insecure-Requests': '1',
            'Origin': 'https://co.metrc.com',
            'User-Agent': USER_AGENT,
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;'
                      'q=0.8,application/signed-exchange;v=b3;q=0.9',
            'Sec-Fetch-Site': 'same-origin',
            'Sec-Fetch-Mode': 'navigate',
            'Sec-Fetch-User': '?1',
            'Sec-Fetch-Dest': 'document',
            'Accept-Language': 'en-GB,en-US;q=0.9,en;q=0.8',
        }

        data = {
            '__RequestVerificationToken': self.req_code,
            'Username': self.username,
            'Password': self.password,
            'Email': self.email,
        }

        loop_count = 0
        async with self.sema:
            while loop_count < 2:
                try:
                    async with self.session.post(url, headers=headers, data=data) as request:
                        response = await request.content.read()
                        content = response.decode('utf-8')
                        html_content = bs(content, "html.parser")
                        title = html_content.find('title').text
                        report_title = html_content.find(attrs={'class': "title"}).text
                        content = content.split('\n')
                        for con in content:
                            if 'ApiVerificationToken' in con:
                                self.api_token = con.strip().replace(',', '').replace("'", '').split(' ')[1]
                        if 'Reports Control Panel' in title and report_title == 'Reports Control Panel':
                            return True
                except:
                    loop_count += 1
                    await asyncio.sleep(3)
            return False

    async def get_reports(self):
        url = 'https://co.metrc.com/api/reports/data/facilities'
        headers = {
            'Connection': 'keep-alive',
            # 'ApiVerificationToken': 'lwK31OysCx_FIQDzAR-2T7zcdOS42jCoYZGXv-yUgGcCqDPGkODNcCmRkl4pzAADnrw-7tfuw8tdUgq0yKqhQf4JpOrRcxdXG0RMMfL4nKqzDAaecLzz3eWXoSJdxCnb3BFRgPVnbc2dM8khYQtfYQ2:jDyBe0DH53rbVu7-QGH14rVFdyPpPgLUjVuFVuEtDPKavJJnpUS5NKHEmVaMCt3_mmghLn3Me9X2THhGbULkvX8QVWOc7f8iDYgV17DY16u3oiYTjyCwOUqDQj29fAWnr_S03Xj7dGnLy_nrhTZskA2',
            'ApiVerificationToken': self.api_token,
            # 'X-NewRelic-ID': 'VgACWF9aDRADVVhRDgUCVFM=',
            # 'tracestate': '2659995@nr=0-1-2659995-320096468-d961ddb365e5ff0e----1654080887208',
            # 'traceparent': '00-f9437e45bb7804422187a231b1addb47-d961ddb365e5ff0e-01',
            'sec-ch-ua-mobile': '?0',
            'User-Agent': USER_AGENT,
            # 'newrelic': 'eyJ2IjpbMCwxXSwiZCI6eyJ0eSI6IkJyb3dzZXIiLCJhYyI6IjI2NTk5OTUiLCJhcCI6IjMyMDA5NjQ2OCIsImlkIjoiZDk2MWRkYjM2NWU1ZmYwZSIsInRyIjoiZjk0MzdlNDViYjc4MDQ0MjIxODdhMjMxYjFhZGRiNDciLCJ0aSI6MTY1NDA4MDg4NzIwOH19',
            'Accept': '*/*',
            'X-Requested-With': 'XMLHttpRequest',
            'sec-ch-ua': '" Not;A Brand";v="99", "Google Chrome";v="97", "Chromium";v="97"',
            'sec-ch-ua-platform': '"Windows"',
            'Sec-Fetch-Site': 'same-origin',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Dest': 'empty',
            'Accept-Language': 'en-GB,en-US;q=0.9,en;q=0.8',
        }

        loop_count = 0
        async with self.sema:
            while loop_count < 2:
                try:
                    async with self.session.get(url, headers=headers) as request:
                        response = await request.content.read()
                        content = response.decode('utf-8')
                        html_content = json.loads(content)
                        for report in self.report[1:]:
                            for con in html_content:
                                if report[0] == con['LicenseNumber']:
                                    self.license.append([con['Id'], con['LicenseNumber']])
                        if self.license:
                            return True
                        else:
                            return False
                except:
                    loop_count += 1
                    await asyncio.sleep(3)
            return False

    async def download_report(self, license):
        url = f"https://co.metrc.com/reports/transfers?id={license[0]}&start={self.start_date}&end={self.end_date}&format=excel"

        headers = {
            'sec-ch-ua': '" Not;A Brand";v="99", "Google Chrome";v="97", "Chromium";v="97"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'Upgrade-Insecure-Requests': '1',
            'User-Agent': USER_AGENT,
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;'
                      'q=0.8,application/signed-exchange;v=b3;q=0.9',
            'host': 'co.metrc.com',
        }

        loop_count = 0
        async with self.sema:
            while loop_count < 2:
                try:
                    async with self.session.get(url, headers=headers) as request:
                        response = await request.content.read()
                        start = self.start_date.replace('/', '-')
                        end = self.end_date.replace('/', '-')
                        file_path = os.path.join(os.getcwd(),'Downloads',f'{start}-{end}')
                        file = os.path.join(file_path, f'{license[1]}.xls')
                        os.makedirs(file_path, exist_ok=True)
                        with open(file, 'wb') as f:
                            f.write(response)
                        await asyncio.sleep(1)
                        self.gui_queue.put({"status": f"Report downloaded for {license[1]}"}) if self.gui_queue else None
                        return True
                except:
                    loop_count += 1
                    await asyncio.sleep(3)
            return False

    async def start_process(self, executor):
        timeout = aiohttp.ClientTimeout(total=TIMEOUT)
        conn = aiohttp.TCPConnector(limit=5, limit_per_host=5)
        self.sema = asyncio.Semaphore(LIMIT)
        async with aiohttp.ClientSession(connector=conn, timeout=timeout) as self.session:
            load_login = await self.load_login()
            if not load_login:
                self.gui_queue.put({"status": "Error while loading login page"}) if self.gui_queue else None
                return False

            login = await self.login()
            if not login:
                self.gui_queue.put({"status": "Error while login into portal"}) if self.gui_queue else None
                return False

            reports = await self.get_reports()
            if not reports:
                self.gui_queue.put({"status": f"Error while fetching report id {self.report} from portal"}) if self.gui_queue else None
                return False

            tasks = []
            for license in self.license:
                tasks.append(self.download_report(license))
            if tasks:
                loop = asyncio.get_event_loop()
                for future in asyncio.as_completed(tasks, loop=loop):
                    download = await future
                    if not download:
                        self.gui_queue.put(
                            {
                                'status': f'Unable to download report of {license[1]}'})
                        continue
        return True

    def download_process(self):
        loop = asyncio.new_event_loop()
        executor = ThreadPoolExecutor(max_workers=3)
        future = asyncio.ensure_future(self.start_process(executor), loop=loop)
        loop.run_until_complete(future)
        return future.result()

class RunMetrc:

    def __init__(self):
        self.gui_queue = None

    def run(self, start_date, end_date, username, password, email, report_data):
        start_time = time.perf_counter()
        self.gui_queue.put({'status': 'Metrc Reconciliation Report Download Processing....'}) if self.gui_queue else None
        # for report in report_data[1:]:
        mertc = Metrc(start_date, end_date, username, password, email, report_data)
        download_status = mertc.download_process()
        if not download_status:
            self.gui_queue.put(
                {'status': f'Error while downloading report.'}) if self.gui_queue else None
            return False

        self.gui_queue.put({'status': 'Metrc Reconciliation Report Downloading Processed.'}) if self.gui_queue else None
        end_time = time.perf_counter()
        time_taken = time.strftime("%H:%M:%S", time.gmtime(int(end_time - start_time)))
        self.gui_queue.put({'status': f'\nTime Taken : {time_taken}'}) if self.gui_queue else None
        return False
