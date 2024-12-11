import os
import requests
from lxml import etree
import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk

# 支持中文显示
plt.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
plt.rcParams['axes.unicode_minus'] = False  # 用来正常显示负号

# 全局变量定义
excelFileName = "豆瓣电影top250数据.xlsx"
csvFileName = "豆瓣电影top250数据.csv"

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36 Edg/108.0.1462.46'
}

def get_first_text(list):
    try:
        return list[0].strip()
    except:
        return ""

def check() -> bool:
    if os.path.exists(csvFileName) and os.path.exists(excelFileName):
        print("文件已存在:", excelFileName, "无需重复爬取")
        return True
    else:
        return False

def getMovieData():
    df = pd.DataFrame(columns=["序号", "标题", "链接", "导演", "评分", "评价人数", "简介", "年份", "地区", "类型"])

    urls = ['https://movie.douban.com/top250?start={}&filter='.format(str(i * 25)) for i in range(10)]
    count = 1
    for url in urls:
        print(f"正在爬取: {url}")
        res = requests.get(url=url, headers=headers)
        html = etree.HTML(res.text)
        lis = html.xpath('//*[@id="content"]/div/div[1]/ol/li')
        for li in lis:
            title = get_first_text(li.xpath('./div/div[2]/div[1]/a/span[1]/text()'))
            src = get_first_text(li.xpath('./div/div[2]/div[1]/a/@href'))
            director_info = li.xpath('./div/div[2]/div[2]/p[1]/text()')
            director = director_info[0].strip().replace("\xa0\xa0\xa0", "\t").split("\t")[0].replace('导演: ', '')
            info_text = director_info[1].strip().replace('\xa0', '').split('/')
            year = info_text[0].strip()
            region = info_text[1].strip()
            genre = info_text[2].strip()

            score = get_first_text(li.xpath('./div/div[2]/div[2]/div/span[2]/text()'))
            comment = get_first_text(li.xpath('./div/div[2]/div[2]/div/span[4]/text()'))
            summary = get_first_text(li.xpath('./div/div[2]/div[2]/p[2]/span/text()'))

            df.loc[len(df.index)] = [count, title, src, director, score, comment, summary, year, region, genre]

            count += 1
            print(f"爬取到电影: {title}")

    df.to_excel(excelFileName, sheet_name="豆瓣电影top250数据", na_rep="")
    df.to_csv(csvFileName)
    print("文件已经生成！")

def plot_year_distribution():
    df = pd.read_csv(csvFileName)
    year_counts = df['年份'].value_counts().sort_index()
    plt.figure(figsize=(14, 8))  # 进一步增加图表的宽度和高度
    bars = plt.bar(year_counts.index, year_counts.values, color='skyblue', width=0.8)  # 增加条形的宽度
    plt.xlabel('年份')
    plt.ylabel('电影数量')
    plt.title('豆瓣电影Top250年份分布')
    # 旋转X轴标签并调整字体大小
    plt.xticks(rotation=45, ha='right', fontsize=10)  # 调整字体大小
    # 显示每个条形的数值
    for bar in bars:
        yval = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2, yval, int(yval), va='bottom', ha='center', fontsize=10)
    plt.tight_layout()  # 调整布局以适应标签
    plt.savefig("year_distribution.png")
    plt.show()

def plot_rating_distribution():
    df = pd.read_csv(csvFileName)
    rating_counts = df['评分'].value_counts().sort_index()
    
    plt.figure(figsize=(8, 8))
    patches, texts, autotexts = plt.pie(rating_counts, labels=None, autopct='%1.1f%%', startangle=90, pctdistance=0.85)
    
    plt.title('豆瓣电影Top 250评分分布')
    
    # 将所有评分及其占比放在图例上
    labels = [f'{index} ({value}次)' for index, value in zip(rating_counts.index, rating_counts.values)]
    plt.legend(patches, labels, title="评分", loc="best", fontsize='small')  # 使用plt.legend()添加图例
    
    plt.savefig("rating_distribution.png")
    plt.show()

def plot_comment_count_by_rating():
    df = pd.read_csv(csvFileName)
    df['评价人数'] = df['评价人数'].str.replace('人评价', '').astype(int)
    plt.figure(figsize=(12, 6))
    plt.scatter(df.index, df['评分'], alpha=0.5)
    plt.xlabel('电影索引')
    plt.ylabel('评分')
    plt.title('评价人数随评分变化')
    plt.gca().get_yaxis().get_major_formatter().set_scientific(False)  # 避免科学计数法
    plt.savefig("comment_count_by_rating.png")
    plt.show()

def plot_top10_comment_counts():
    df = pd.read_csv(csvFileName)
    df['评价人数'] = df['评价人数'].str.replace('人评价', '').astype(int)
    top10_comments = df.nlargest(10, '评价人数')
    
    plt.figure(figsize=(10, 6))
    # 使用horizontal bar chart (水平条形图)
    bars = plt.barh(top10_comments['标题'], top10_comments['评价人数'], color="lightgreen")
    
    plt.xlabel('评价人数')
    plt.ylabel('电影名称')
    plt.title('评论人数TOP10电影')
    
    # 为每个条形添加评价人数标签
    for bar in bars:
        width = bar.get_width()
        plt.text(width, bar.get_y() + bar.get_height()/2, f'{int(width)}', va='center', ha='left', fontsize=8)
    
    plt.tight_layout()
    plt.savefig("top10_comment_counts.png")
    plt.show()

def plot_genre_wordcloud():
    df = pd.read_csv(csvFileName)
    genre_text = ' '.join(df['类型'].dropna().astype(str).tolist())
    
    # 指定支持中文的字体路径
    font_path = 'C:\\Windows\\Fonts\\msyh.ttc'  # 例如使用微软雅黑字体
    
    from wordcloud import WordCloud
    wordcloud = WordCloud(font_path=font_path, width=800, height=400, background_color='white').generate(genre_text)
    plt.figure(figsize=(10, 5))
    plt.imshow(wordcloud, interpolation='bilinear')
    plt.axis('off')  # 关闭坐标轴
    plt.title('电影类型词云图')
    plt.savefig("genre_wordcloud.png")
    plt.show()

def analyze_data(choice):
    if choice == '1':
        plot_year_distribution()
    elif choice == '2':
        plot_rating_distribution()
    elif choice == '3':
        plot_comment_count_by_rating()
    elif choice == '4':
        plot_top10_comment_counts()
    elif choice == '5':
        plot_genre_wordcloud()

def create_gui():
    window = tk.Tk()
    window.title("豆瓣电影Top 250数据分析")
    window.geometry('400x300')

    def show_image(image_path):
        image = Image.open(image_path)
        image = ImageTk.PhotoImage(image)
        label = tk.Label(window, image=image)
        label.image = image
        label.pack()

    def analyze_data_and_show(choice, image_path):
        analyze_data(choice)
        show_image(image_path)

    tk.Button(window, text="Top250年份分布条状图", command=lambda: analyze_data_and_show('1', "year_distribution.png")).pack(fill=tk.X)
    tk.Button(window, text="豆瓣电影Top 250评分分布饼图", command=lambda: analyze_data_and_show('2', "rating_distribution.png")).pack(fill=tk.X)
    tk.Button(window, text="评价人数随评分变化散点图", command=lambda: analyze_data_and_show('3', "comment_count_by_rating.png")).pack(fill=tk.X)
    tk.Button(window, text="评论人数TOP10条状图", command=lambda: analyze_data_and_show('4', "top10_comment_counts.png")).pack(fill=tk.X)
    tk.Button(window, text="电影类型词云图", command=lambda: analyze_data_and_show('5', "genre_wordcloud.png")).pack(fill=tk.X)

    window.mainloop()

def main():
    if not check():
        getMovieData()
    create_gui()

if __name__ == '__main__':
    main()