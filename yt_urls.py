#!/usr/bin/env python3
import sys
from bs4 import BeautifulSoup
import html
import pandas as pd


def load_html(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()
    return BeautifulSoup(content, 'html.parser')


def list_ytd_rich_item_renderer(soup):
    return soup.find_all('ytd-rich-item-renderer', recursive=True)


def extract_video_data(videos):
    video_data = []
    for video in reversed(videos):
        a = video.find('a', id='video-title-link')
        url = a['href']
        title = html.unescape(a['title'])
        celltext = f'=HYPERLINK("{url}", "{title}")'
        video_data.append([celltext])
    return video_data


def save_to_excel(data, channel_name, file_name='urls.xlsx'):
    df = pd.DataFrame(data, columns=[channel_name])
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='w') as writer:
        df.to_excel(writer, sheet_name=channel_name, index=False)


def main():
    try:
        input('あらかじめ「ウェブページ、完全」でsrc.htmlとして保存してください。Enterを押すと処理を開始します。')
    except KeyboardInterrupt:
        print('\n処理を中断しました。')
        sys.exit()
    file_path = sys.argv[1] if len(sys.argv) > 1 else 'src.html'
    soup = load_html(file_path)
    channel_name = soup.title.string.split(' - ')[0]
    videos = list_ytd_rich_item_renderer(soup)
    video_data = extract_video_data(videos)
    save_to_excel(video_data, channel_name)


if __name__ == '__main__':
    main()
