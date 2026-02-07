# -*- coding: utf-8 -*-
"""
小红书数据看板 - 自动更新脚本
============================

这个脚本做的事情：
1. 读取你从小红书下载的 Excel 文件
2. 自动处理数据、分析关键词
3. 更新网页看板
4. 推送到 GitHub（让网页链接自动更新）

使用方法：
1. 从小红书创作者中心下载数据 Excel
2. 双击运行这个脚本（或在命令行运行 python update_dashboard.py）
3. 脚本会自动找到最新的 Excel 文件并更新

作者：Claude Code
"""

import os
import re
import json
import glob
from datetime import datetime
from collections import Counter

# ========== 配置区域 ==========
# 你可以修改这些设置

# Excel 文件所在的文件夹（下载文件夹）
DOWNLOADS_FOLDER = r"C:\Users\90543\Downloads"

# 项目文件夹（看板代码所在位置）
PROJECT_FOLDER = r"C:\Users\90543\Projects\xiaohongshu-dashboard"

# Excel 文件名匹配规则（小红书导出的文件通常包含这些关键词）
EXCEL_PATTERNS = [
    "*数据*分析*.xlsx",
    "*小红书*.xlsx",
    "*笔记*数据*.xlsx",
    "*内容*数据*.xlsx",
    "*.xlsx"  # 最后兜底：任何 Excel 文件
]

# Python 路径（Windows Store 的 Python 可能有问题，所以指定完整路径）
PYTHON_PATH = r"C:\Users\90543\AppData\Local\Programs\Python\Python312\python.exe"


# ========== 工具函数 ==========

def print_step(step_num, message):
    """打印步骤信息，让用户知道进度"""
    print(f"\n{'='*50}")
    print(f"  步骤 {step_num}: {message}")
    print(f"{'='*50}")


def find_latest_excel():
    """
    在下载文件夹中找到最新的 Excel 文件

    原理：
    - 按照预设的文件名模式搜索
    - 如果找到多个文件，选择最新修改的那个
    """
    print_step(1, "寻找 Excel 数据文件")

    all_excel_files = []

    # 按照不同的模式搜索文件
    for pattern in EXCEL_PATTERNS:
        search_path = os.path.join(DOWNLOADS_FOLDER, pattern)
        found_files = glob.glob(search_path)

        # 过滤掉临时文件（以 ~ 开头的文件）
        found_files = [f for f in found_files if not os.path.basename(f).startswith('~')]

        if found_files:
            print(f"  找到匹配 '{pattern}' 的文件: {len(found_files)} 个")
            all_excel_files.extend(found_files)
            break  # 找到就停，不继续用更宽泛的模式

    if not all_excel_files:
        print("\n  ❌ 没有找到 Excel 文件！")
        print(f"  请确认下载文件夹路径正确: {DOWNLOADS_FOLDER}")
        print("  请确认已从小红书下载了数据文件")
        return None

    # 按修改时间排序，选最新的
    all_excel_files.sort(key=os.path.getmtime, reverse=True)
    latest_file = all_excel_files[0]

    # 显示文件信息
    file_time = datetime.fromtimestamp(os.path.getmtime(latest_file))
    print(f"\n  ✅ 选择文件: {os.path.basename(latest_file)}")
    print(f"  📅 修改时间: {file_time.strftime('%Y-%m-%d %H:%M:%S')}")

    return latest_file


def read_excel_data(excel_path):
    """
    读取 Excel 文件中的数据

    小红书导出的 Excel 通常包含这些列：
    - 笔记标题
    - 发布时间
    - 笔记类型（图文/视频）
    - 曝光量
    - 阅读量
    - 点击率
    - 点赞数
    - 评论数
    - 收藏数
    - 涨粉数
    - 分享数
    - 平均阅读时长
    """
    print_step(2, "读取 Excel 数据")

    try:
        # 尝试导入 pandas（处理 Excel 的库）
        import pandas as pd
    except ImportError:
        print("\n  ❌ 需要安装 pandas 库来读取 Excel")
        print("  请运行以下命令安装：")
        print(f'  "{PYTHON_PATH}" -m pip install pandas openpyxl')
        return None

    try:
        # 读取 Excel 文件
        df = pd.read_excel(excel_path)
        print(f"  ✅ 成功读取 {len(df)} 行数据")
        print(f"  📊 列名: {list(df.columns)}")
        return df
    except Exception as e:
        print(f"\n  ❌ 读取 Excel 失败: {e}")
        return None


def process_data(df):
    """
    处理数据，转换成看板需要的格式

    这个函数做的事：
    1. 识别 Excel 中的列名（小红书可能用不同的名称）
    2. 提取需要的数据
    3. 计算月份信息
    4. 分析标题关键词
    """
    print_step(3, "处理数据")

    # ===== 列名映射 =====
    # 小红书导出的列名可能不一样，这里列出可能的名称
    column_mappings = {
        'title': ['笔记标题', '标题', 'title', '内容标题'],
        'date': ['发布时间', '发布日期', 'date', '时间', '创建时间'],
        'type': ['笔记类型', '类型', 'type', '内容类型'],
        'impressions': ['曝光量', '曝光', 'impressions', '展现量', '展现'],
        'views': ['阅读量', '观看量', 'views', '播放量', '点击量'],
        'ctr': ['点击率', 'ctr', '点击转化率'],
        'likes': ['点赞数', '点赞', 'likes', '赞'],
        'comments': ['评论数', '评论', 'comments'],
        'collects': ['收藏数', '收藏', 'collects', '收藏量'],
        'followers': ['涨粉数', '涨粉', 'followers', '新增粉丝', '粉丝增长'],
        'shares': ['分享数', '分享', 'shares', '分享量'],
        'avgViewTime': ['平均阅读时长', '平均观看时长', 'avgViewTime', '平均时长']
    }

    # 找到实际的列名
    actual_columns = {}
    df_columns_lower = {col.lower().strip(): col for col in df.columns}

    for key, possible_names in column_mappings.items():
        for name in possible_names:
            # 尝试精确匹配
            if name in df.columns:
                actual_columns[key] = name
                break
            # 尝试小写匹配
            if name.lower() in df_columns_lower:
                actual_columns[key] = df_columns_lower[name.lower()]
                break
            # 尝试包含匹配
            for col in df.columns:
                if name in col or name.lower() in col.lower():
                    actual_columns[key] = col
                    break
            if key in actual_columns:
                break

    print(f"  识别到的列: {actual_columns}")

    # ===== 转换数据 =====
    notes = []

    for _, row in df.iterrows():
        note = {}

        # 标题（必需）
        if 'title' in actual_columns:
            note['title'] = str(row[actual_columns['title']]).strip()
        else:
            continue  # 没有标题就跳过

        # 日期
        if 'date' in actual_columns:
            date_val = row[actual_columns['date']]
            if pd.notna(date_val):
                # 处理不同的日期格式
                if hasattr(date_val, 'strftime'):
                    note['date'] = date_val.strftime('%Y-%m-%d')
                else:
                    # 尝试解析字符串日期
                    try:
                        parsed_date = pd.to_datetime(str(date_val))
                        note['date'] = parsed_date.strftime('%Y-%m-%d')
                    except:
                        note['date'] = str(date_val)[:10]
                note['month'] = note['date'][:7]  # 提取 YYYY-MM

        # 类型
        if 'type' in actual_columns:
            type_val = str(row[actual_columns['type']]).strip()
            note['type'] = '视频' if '视频' in type_val else '图文'
        else:
            note['type'] = '图文'

        # 数值字段
        numeric_fields = ['impressions', 'views', 'likes', 'comments',
                         'collects', 'followers', 'shares', 'avgViewTime']

        for field in numeric_fields:
            if field in actual_columns:
                val = row[actual_columns[field]]
                # 处理可能的非数字情况
                try:
                    if pd.isna(val):
                        note[field] = 0
                    else:
                        note[field] = int(float(str(val).replace(',', '').replace('%', '')))
                except:
                    note[field] = 0
            else:
                note[field] = 0

        # 点击率特殊处理（可能是百分比格式）
        if 'ctr' in actual_columns:
            ctr_val = row[actual_columns['ctr']]
            try:
                if pd.isna(ctr_val):
                    note['ctr'] = 0
                elif '%' in str(ctr_val):
                    note['ctr'] = float(str(ctr_val).replace('%', '')) / 100
                elif float(ctr_val) > 1:
                    note['ctr'] = float(ctr_val) / 100
                else:
                    note['ctr'] = float(ctr_val)
            except:
                note['ctr'] = 0
        else:
            # 如果没有点击率，尝试计算
            if note.get('impressions', 0) > 0:
                note['ctr'] = round(note.get('views', 0) / note['impressions'], 3)
            else:
                note['ctr'] = 0

        notes.append(note)

    print(f"  ✅ 成功处理 {len(notes)} 篇笔记")

    # ===== 提取月份列表 =====
    months = sorted(list(set(n.get('month', '') for n in notes if n.get('month'))))
    print(f"  📅 时间范围: {months[0] if months else 'N/A'} 到 {months[-1] if months else 'N/A'}")

    # ===== 分析关键词 =====
    high_likes_keywords = analyze_keywords(notes, 'likes')
    high_followers_keywords = analyze_keywords(notes, 'followers')

    return {
        'notes': notes,
        'months': months,
        'highLikesKeywords': high_likes_keywords,
        'highFollowersKeywords': high_followers_keywords
    }


def analyze_keywords(notes, metric, top_n=30):
    """
    分析标题中的关键词

    原理：
    1. 取表现最好的前 20 篇笔记
    2. 提取标题中的关键词
    3. 统计出现频率
    """
    import re

    # 按指标排序，取前 20
    sorted_notes = sorted(notes, key=lambda x: x.get(metric, 0), reverse=True)[:20]

    # 提取所有标题
    titles = ' '.join(n.get('title', '') for n in sorted_notes)

    # 简单的中文分词（按标点和常见词分割）
    # 移除表情符号和特殊字符
    titles = re.sub(r'[\U0001F600-\U0001F64F\U0001F300-\U0001F5FF\U0001F680-\U0001F6FF\U0001F1E0-\U0001F1FF]', '', titles)
    titles = re.sub(r'[^\w\s\u4e00-\u9fff]', ' ', titles)

    # 分词（简单的按空格和长度切分）
    words = []
    for word in titles.split():
        word = word.strip().lower()
        if len(word) >= 2:
            words.append(word)

    # 对中文进行简单的 n-gram 切分
    chinese_text = ''.join(re.findall(r'[\u4e00-\u9fff]+', titles))
    for i in range(len(chinese_text) - 1):
        words.append(chinese_text[i:i+2])
    for i in range(len(chinese_text) - 2):
        words.append(chinese_text[i:i+3])

    # 停用词（常见但无意义的词）
    stopwords = {'的', '了', '是', '在', '我', '有', '和', '就', '不', '人', '都',
                 '一', '个', '上', '这', '为', '吗', '你', '到', '说', '要', '会',
                 '来', '对', '可以', '什么', '没有', '怎么', '那么', '这个', '一个'}

    # 统计词频
    word_counts = Counter(w for w in words if w not in stopwords and len(w) >= 2)

    # 返回前 N 个
    return word_counts.most_common(top_n)


def update_html(data):
    """
    更新 HTML 文件中的数据

    原理：
    - 找到 index.html 中的 rawData 变量
    - 用新数据替换掉旧数据
    """
    print_step(4, "更新看板文件")

    html_path = os.path.join(PROJECT_FOLDER, "index.html")

    try:
        # 读取当前 HTML
        with open(html_path, 'r', encoding='utf-8') as f:
            html_content = f.read()

        # 把数据转成 JSON 格式
        data_json = json.dumps(data, ensure_ascii=False, indent=2)

        # 用正则表达式替换 rawData
        # 匹配 const rawData = {...}; 这一段
        pattern = r'const rawData = \{[\s\S]*?\};'
        replacement = f'const rawData = {data_json};'

        new_html = re.sub(pattern, replacement, html_content, count=1)

        # 写回文件
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(new_html)

        print(f"  ✅ 已更新 {html_path}")
        return True

    except Exception as e:
        print(f"  ❌ 更新 HTML 失败: {e}")
        return False


def push_to_github():
    """
    推送更新到 GitHub

    这样你的网页链接就会自动更新
    """
    print_step(5, "推送到 GitHub")

    try:
        import subprocess

        os.chdir(PROJECT_FOLDER)

        # Git 命令序列
        commands = [
            ['git', 'add', 'index.html'],
            ['git', 'commit', '-m', f'更新数据 {datetime.now().strftime("%Y-%m-%d %H:%M")}'],
            ['git', 'push']
        ]

        for cmd in commands:
            print(f"  执行: {' '.join(cmd)}")
            result = subprocess.run(cmd, capture_output=True, text=True)
            if result.returncode != 0 and 'nothing to commit' not in result.stdout + result.stderr:
                print(f"  ⚠️ 命令输出: {result.stderr or result.stdout}")

        print("  ✅ 已推送到 GitHub")
        print("  🌐 几分钟后访问你的看板链接查看更新")
        return True

    except Exception as e:
        print(f"  ⚠️ 推送失败: {e}")
        print("  你可以稍后手动推送：")
        print("  1. 打开命令行")
        print(f"  2. cd {PROJECT_FOLDER}")
        print("  3. git add . && git commit -m '更新数据' && git push")
        return False


def main():
    """
    主函数 - 串联所有步骤
    """
    print("\n" + "="*60)
    print("   🔴 小红书数据看板 - 自动更新工具")
    print("="*60)
    print(f"\n⏰ 开始时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # 步骤 1: 找到 Excel 文件
    excel_path = find_latest_excel()
    if not excel_path:
        input("\n按回车键退出...")
        return

    # 步骤 2: 读取数据
    df = read_excel_data(excel_path)
    if df is None:
        input("\n按回车键退出...")
        return

    # 步骤 3: 处理数据
    try:
        import pandas as pd
        data = process_data(df)
    except Exception as e:
        print(f"\n❌ 处理数据时出错: {e}")
        input("\n按回车键退出...")
        return

    # 步骤 4: 更新 HTML
    if not update_html(data):
        input("\n按回车键退出...")
        return

    # 步骤 5: 推送到 GitHub
    push_to_github()

    # 完成
    print("\n" + "="*60)
    print("   ✅ 更新完成！")
    print("="*60)
    print(f"\n📊 共更新 {len(data['notes'])} 篇笔记数据")
    print(f"📅 时间范围: {data['months'][0]} 到 {data['months'][-1]}")
    print(f"\n🌐 看板链接: https://belljia95.github.io/xiaohongshu-dashboard/")

    input("\n按回车键退出...")


if __name__ == "__main__":
    main()
