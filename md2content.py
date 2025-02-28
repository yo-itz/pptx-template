from pptx import Presentation
from pptx.util import Inches
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import os
import markdown
from collections import defaultdict

# プレゼンテーションオブジェクトの作成
prs = Presentation()
# アスペクト比を16:9に設定
prs.slide_width = Inches(16)
prs.slide_height = Inches(9)

# マークダウン内容を構造化したデータ

def parse_markdown(md_file_path):
    """
    Markdownファイルを読み込み、スライドの内容を構造化して辞書形式で返す
    """
    content = defaultdict(list)  # 各スライド内容を格納する辞書
    
    if not os.path.exists(md_file_path):
        raise FileNotFoundError(f"Markdownファイルが見つかりません: {md_file_path}")
    
    with open(md_file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    current_slide = None
    current_section = None
    for line in lines:
        line = line.strip()

        if line.startswith('# '):  # 見出し1はスライドのタイトル
            current_slide = line[2:]  # '# ' を削除
            content[current_slide]  # 新しいスライドが始まる
            current_section = None
        elif line.startswith('## '):  # 見出し2はセクションタイトル
            current_section = line[3:]  # '## ' を削除
        elif line.startswith('- '):  # リスト項目
            if current_slide:
                if current_section:
                    content[current_slide].append(f"{current_section} - {line[2:]}")
                else:
                    content[current_slide].append(line[2:])
        elif line == '':  # 空行はスライド区切り
            current_section = None

    return dict(content)

# Markdownファイルの読み込みと解析
md_file_path = 'report.md'  # 解析するMarkdownファイル
content = parse_markdown(md_file_path)

# 変換結果の表示（内容が正しく変換されたか確認）
for slide, sections in content.items():
    print(f"スライド: {slide}")
    for section in sections:
        print(f"  - {section}")

# スライドの追加
def add_slide(title, content=None, img_path=None, img_width=None, img_height=None, video_path=None):
    slide_layout = prs.slide_layouts[5]  # シンプルなレイアウトを使用
    slide = prs.slides.add_slide(slide_layout)
    
    # タイトルの追加
    title_placeholder = slide.shapes.title
    if title_placeholder:
        title_placeholder.text = title
    
    # コンテンツテキストがあれば追加
    if content:
        text_box = slide.shapes.add_textbox(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.5), Inches(9), Inches(4))
        text_frame = text_box.text_frame
        p = text_frame.add_paragraph()
        p.text = content
    

# タイトルスライド
slide_layout = prs.slide_layouts[0]  # タイトルスライド
slide = prs.slides.add_slide(slide_layout)
title = slide.shapes.title
subtitle = slide.placeholders[1]
title.text = content["title"][0] if "title" in content else "タイトル"
subtitle.text = content["subtitle"][0] if "subtitle" in content else "サブタイトル"

# 目次スライド
add_slide("目次", 
          content="\n".join(content.get("目次", [])))

# プロジェクト概要スライド
add_slide("プロジェクト概要", 
          content="\n".join(content.get("プロジェクト概要", [])))

# 開発の目的と目標スライド
add_slide("開発の目的と目標", 
          content="\n".join(content.get("開発の目的と目標", [])))

# 開発手法と技術スライド
add_slide("開発手法と技術", 
          content="\n".join(content.get("開発手法と技術", [])), 
          img_path="images/model_architecture.png", img_width=Inches(8), img_height=Inches(4))

# 結果と成果スライド
add_slide("結果と成果", 
          content="\n".join(content.get("結果と成果", [])), 
          img_path="images/performance_graph.png", img_width=Inches(8), img_height=Inches(4))

# デモスライド
add_slide("デモ", 
          content="\n".join(content.get("デモ", [])), 
          video_path="images/demo_video.mp4")

# 認識結果スライド
add_slide("認識結果", 
          content="\n".join(content.get("認識結果", [])), 
          img_path="images/result_image.png", img_width=Inches(8), img_height=Inches(4))

# 今後の展開スライド
add_slide("今後の展開", 
          content="\n".join(content.get("今後の展開", [])))

# 結論と提案スライド
add_slide("結論と提案", 
          content="\n".join(content.get("結論と提案", [])))

# Q&Aスライド
add_slide("Q&A", content="\n".join(content.get("Q&A", [])))

# プレゼンテーションファイルの保存
prs.save("画像認識AI開発レポート.pptx")

print("PowerPointファイルが作成されました。")