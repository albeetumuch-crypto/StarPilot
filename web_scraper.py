#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
網頁爬蟲程式 - 抓取 i23.uk 前三篇文章
"""

import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
import time
import json
import re

def scrape_article(url):
    """抓取單篇文章的完整內容"""
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        response = requests.get(url, headers=headers, timeout=10)
        response.encoding = 'utf-8'
        
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # 抓取文章標題
        title_elem = soup.find('h1', class_='post-title')
        title = title_elem.get_text(strip=True) if title_elem else "未找到標題"
        
        # 從 JSON-LD 結構化數據中抓取內容
        content = ""
        json_ld = soup.find('script', {'type': 'application/ld+json'})
        
        if json_ld:
            try:
                # 查找包含 articleBody 的 JSON-LD
                json_content = json_ld.string
                # 嘗試找到所有 JSON-LD 資料
                for script in soup.find_all('script', {'type': 'application/ld+json'}):
                    try:
                        data = json.loads(script.string)
                        if isinstance(data, dict) and 'articleBody' in data:
                            content = data['articleBody']
                            # 清理特殊字符
                            content = content.replace('\\n', '\n').replace('\\u0026', '&')
                            content = re.sub(r'\\u[0-9a-fA-F]{4}', '', content)
                            break
                    except:
                        continue
            except Exception as e:
                print(f"  解析 JSON-LD 時出錯: {e}")
        
        # 如果從 JSON-LD 沒有找到，嘗試從 HTML 中提取
        if not content:
            post_content = soup.find('div', class_='post-content')
            if post_content:
                # 移除導航和其他元素，只保留段落
                for nav in post_content.find_all('nav'):
                    nav.decompose()
                for hr in post_content.find_all('hr'):
                    hr.decompose()
                    
                paragraphs = post_content.find_all('p')
                content = '\n'.join([p.get_text(strip=True) for p in paragraphs if p.get_text(strip=True)])
        
        if not content:
            content = "未找到文章內容"
        
        # 抓取發佈日期
        date_elem = soup.find('span', class_='post-meta')
        if date_elem:
            # 提取第一個包含日期的 span
            span = date_elem.find('span')
            if span and span.has_attr('title'):
                date = span['title']
            else:
                date = date_elem.get_text(strip=True).split('·')[0]
        else:
            date = "未找到日期"
        
        return {
            'title': title,
            'content': content,
            'date': date,
            'url': url
        }
    except Exception as e:
        print(f"抓取 {url} 時出錯: {e}")
        return None

def scrape_homepage():
    """抓取首頁並獲取前三篇文章的連結"""
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        response = requests.get('https://www.i23.uk/', headers=headers, timeout=10)
        response.encoding = 'utf-8'
        
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # 找到所有文章連結
        article_links = []
        articles = soup.find_all('article', class_='post-entry')
        
        for idx, article in enumerate(articles[:3]):  # 只取前3篇
            link_elem = article.find('a', class_='entry-link')
            title_elem = article.find('h2', class_='entry-hint-parent')
            
            if link_elem and title_elem:
                href = link_elem.get('href')
                title = title_elem.get_text(strip=True)
                
                # 轉換相對 URL 為絕對 URL
                full_url = urljoin('https://www.i23.uk/', href)
                article_links.append({
                    'title': title,
                    'url': full_url
                })
                print(f"找到文章 {idx+1}: {title}")
        
        return article_links
    except Exception as e:
        print(f"抓取首頁時出錯: {e}")
        return []

def main():
    print("開始抓取 i23.uk 網站...\n")
    
    # 第一步：抓取首頁獲取文章連結
    print("步驟 1: 掃描首頁...")
    article_links = scrape_homepage()
    
    if not article_links:
        print("無法找到文章連結")
        return
    
    print(f"\n找到 {len(article_links)} 篇文章\n")
    
    # 第二步：抓取每篇文章的詳細內容
    articles = []
    for idx, article_info in enumerate(article_links, 1):
        print(f"步驟 2.{idx}: 抓取文章 - {article_info['title']}")
        article_data = scrape_article(article_info['url'])
        if article_data:
            articles.append(article_data)
        time.sleep(1)  # 禮貌性延遲
    
    # 第三步：將內容存成 TXT 檔
    print("\n步驟 3: 保存到 TXT 檔...")
    
    output_filename = '/workspaces/StarPilot/i23_articles.txt'
    
    with open(output_filename, 'w', encoding='utf-8') as f:
        f.write("=" * 80 + "\n")
        f.write("i23.uk - 小明的 AI 科技觀點 - 前三篇文章\n")
        f.write("=" * 80 + "\n\n")
        
        for idx, article in enumerate(articles, 1):
            f.write(f"\n【文章 {idx}】\n")
            f.write("-" * 80 + "\n")
            f.write(f"標題: {article['title']}\n")
            f.write(f"日期: {article['date']}\n")
            f.write(f"URL: {article['url']}\n")
            f.write("-" * 80 + "\n\n")
            f.write(f"內容:\n{article['content']}\n")
            f.write("\n" + "=" * 80 + "\n")
    
    print(f"\n✓ 成功！文章已保存到: {output_filename}")
    print(f"✓ 共抓取 {len(articles)} 篇文章")

if __name__ == '__main__':
    main()
