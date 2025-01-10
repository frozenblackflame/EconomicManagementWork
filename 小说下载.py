import requests
from bs4 import BeautifulSoup
import json
import os
from concurrent.futures import ThreadPoolExecutor
import re
import concurrent.futures

def get_headers():
    """返回请求头"""
    return {
        'authority': 'www.tkxyk.cc',
        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
        'accept-language': 'zh-CN,zh-TW;q=0.9,zh;q=0.8,ja;q=0.7,en;q=0.6',
        'cache-control': 'max-age=0',
        'dnt': '1',
        'sec-ch-ua': '"Google Chrome";v="131", "Chromium";v="131", "Not_A Brand";v="24"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
        'sec-fetch-dest': 'document',
        'sec-fetch-mode': 'navigate',
        'sec-fetch-site': 'same-origin',
        'sec-fetch-user': '?1',
        'upgrade-insecure-requests': '1',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'
    }

def get_page_urls(html):
    """从选择器中提取所有目录页面的URL"""
    soup = BeautifulSoup(html, 'html.parser')
    select = soup.find('select', id='indexselect')
    if not select:
        print("警告: 未找到目录选择器")
        return []
    
    urls = []
    base_url = "https://www.tkxyk.cc"
    for option in select.find_all('option'):
        url = base_url + option['value']
        urls.append(url)
    print(f"找到 {len(urls)} 个目录页面")
    return urls

def get_chapter_urls(page_url):
    """获取每个目录页面中的章节URL"""
    try:
        print(f"正在获取页面: {page_url}")
        response = requests.get(page_url, headers=get_headers(), timeout=10)
        response.encoding = 'utf-8'
        soup = BeautifulSoup(response.text, 'html.parser')
        chapters = []
        
        # 只获取chapters类下的链接
        chapter_list = soup.find('ul', class_='chapters')
        if not chapter_list:
            print(f"警告: 在页面 {page_url} 中未找到章节列表")
            return []
            
        chapter_links = chapter_list.find_all('a')
        if not chapter_links:
            print(f"警告: 在页面 {page_url} 中未找到章节链接")
            return []
            
        for link in chapter_links:
            href = link.get('href', '')
            if 'javascript:;' not in href:
                chapter_url = f"https://www.tkxyk.cc{href}"
                chapter_title = link.text.strip()
                chapters.append({
                    'title': chapter_title,
                    'url': chapter_url
                })
        print(f"从 {page_url} 获取到 {len(chapters)} 个章节")
        return chapters
    except Exception as e:
        print(f"获取页面 {page_url} 失败: {str(e)}")
        return []

def get_novel_info(html):
    """从第一页获取小说名称和所有目录页URLs"""
    soup = BeautifulSoup(html, 'html.parser')
    novel_title = soup.find('h1').find('a').text.strip()
    page_urls = get_page_urls(html)
    return novel_title, page_urls

def get_chapter_content(url, chapter_title, folder_path):
    """获取章节内容并保存到单独文件"""
    try:
        print(f"开始下载章节: {chapter_title}")
        print(f"正在访问URL: {url}")
        
        all_content = []
        current_url = url
        
        while True:
            print(f"正在获取页面: {current_url}")
            response = requests.get(current_url, headers=get_headers(), timeout=10)
            response.encoding = 'utf-8'
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # 获取内容
            content_div = soup.find('div', class_='articlecontent')
            if content_div:
                paragraphs = content_div.find_all('p')
                content = '\n'.join(p.text.strip() for p in paragraphs if p.text.strip())
                all_content.append(content)
            
            # 检查是否为最后一页
            if "下一章" in soup.text:
                break
                
            # 获取下一页URL
            next_link = soup.find('a', id='next_url')
            if not next_link:
                break
                
            current_url = f"https://www.tkxyk.cc{next_link['href']}"
        
        # 保存到单独的文件
        chapter_file = os.path.join(folder_path, f"{chapter_title}.txt")
        with open(chapter_file, 'w', encoding='utf-8') as f:
            f.write(f"{chapter_title}\n\n")
            f.write('\n'.join(all_content))
        
        print(f"章节下载完成: {chapter_title}")
        return chapter_title
        
    except Exception as e:
        print(f"获取章节内容失败 {chapter_title}: {str(e)}")
        return None

def merge_chapters(folder_path, output_file, chapters):
    """按顺序合并所有章节文件"""
    try:
        print("开始合并章节文件...")
        with open(output_file, 'w', encoding='utf-8') as outfile:
            for chapter in chapters:
                chapter_file = os.path.join(folder_path, f"{chapter['title']}.txt")
                if os.path.exists(chapter_file):
                    with open(chapter_file, 'r', encoding='utf-8') as infile:
                        outfile.write(infile.read())
                        outfile.write("\n\n")
        print(f"合并完成，已保存到: {output_file}")
    except Exception as e:
        print(f"合并章节失败: {str(e)}")

def download_novel():
    """下载小说内容"""
    try:
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        json_file = os.path.join(desktop_path, f"{novel_title}.json")
        
        # 创建小说专用文件夹
        novel_folder = os.path.join(desktop_path, novel_title)
        os.makedirs(novel_folder, exist_ok=True)
        print(f"创建文件夹: {novel_folder}")
        
        # 读取JSON文件
        with open(json_file, 'r', encoding='utf-8') as f:
            chapters = json.load(f)
        
        # 使用多线程下载内容
        with ThreadPoolExecutor(max_workers=5) as executor:
            future_to_chapter = {
                executor.submit(get_chapter_content, 
                              chapter['url'], 
                              chapter['title'],
                              novel_folder): chapter
                for chapter in chapters
            }
            
            # 等待所有下载完成
            completed_chapters = []
            for future in concurrent.futures.as_completed(future_to_chapter):
                chapter_title = future.result()
                if chapter_title:
                    completed_chapters.append(chapter_title)
        
        # 合并所有章节
        output_file = os.path.join(desktop_path, f"{novel_title}.txt")
        merge_chapters(novel_folder, output_file, chapters)
        
    except Exception as e:
        print(f"下载小说失败: {str(e)}")

def main():
    try:
        # 获取第一页
        first_page_url = "https://www.tkxyk.cc/indexlist/342265/1.html"
        print(f"正在获取首页: {first_page_url}")
        response = requests.get(first_page_url, headers=get_headers(), timeout=10)
        response.encoding = 'utf-8'
        
        # 从第一页获取小说名称和所有目录页URLs
        global novel_title  # 使其成为全局变量以便download_novel使用
        novel_title, page_urls = get_novel_info(response.text)
        if not page_urls:
            print("错误: 未能获取到任何目录页面URL")
            return
        
        # 使用多线程获取所有章节
        all_chapters = []
        with ThreadPoolExecutor(max_workers=5) as executor:
            results = list(executor.map(
                lambda url: get_chapter_urls(url),
                page_urls
            ))
            for chapters in results:
                all_chapters.extend(chapters)
        
        if not all_chapters:
            print("错误: 未能获取到任何章节")
            return
            
        # 保存到桌面，使用小说名称作为文件名
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        json_path = os.path.join(desktop_path, f"{novel_title}.json")
        
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(all_chapters, f, ensure_ascii=False, indent=2)
        
        print(f"成功保存 {len(all_chapters)} 个章节到 {json_path}")
        
        # 开始下载小说内容
        print("开始下载小说内容...")
        download_novel()
        
    except Exception as e:
        print(f"程序执行出错: {str(e)}")

if __name__ == "__main__":
    main()
