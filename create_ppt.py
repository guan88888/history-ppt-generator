import os
import sys
from pptx import Presentation
from pptx.util import Inches
from typing import List, Dict, Any

# --- 常量与配置 ---

# 建议：将配置与代码分离，未来可移至JSON或YAML文件
# 建议：使用常量代替“魔术数字”，增强可读性
TITLE_SLIDE_LAYOUT = 5
TITLE_AND_CONTENT_LAYOUT = 1

# 脚本所在的目录，确保图片路径正确
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILENAME = os.path.join(BASE_DIR, '古代人的日常生活.pptx')

SLIDE_DATA: List[Dict[str, Any]] = [
    {
        "layout": TITLE_SLIDE_LAYOUT,
        "title": "古代女子的梳妆台",
        "image": "image_0.jpg",
        "is_cover": True,
    },
    {
        "layout": TITLE_AND_CONTENT_LAYOUT,
        "title": "前言",
        "content": '《古代人的日常生活》一书，为我们揭示了古人衣食住行的方方面面。\n爱美之心，人皆有之。那么，古代女子是如何“妆”点自己的呢？',
        "image": "image_1.jpg",
    },
    # ... (其他幻灯片数据与之前相同) ...
    {
        "layout": TITLE_AND_CONTENT_LAYOUT,
        "title": "粉底——米粉与铅粉",
        "content": '最早的粉底是用米做的米粉，但容易脱落。\n后来出现了铅粉，美白效果更好，但长期使用有毒。',
        "image": "image_2.jpg",
    },
    {
        "layout": TITLE_AND_CONTENT_LAYOUT,
        "title": "腮红与口红——朱砂与胭脂",
        "content": '朱砂是早期的腮红和口红原料。\n后来，张骞出使西域带回胭脂，成为主流。',
        "image": "image_3.jpg",
    },
    {
        "layout": TITLE_AND_CONTENT_LAYOUT,
        "title": "眉妆——黛与柳枝",
        "content": '“粉黛”中的“黛”指的就是画眉的黑色矿物。\n在黛出现前，人们用烧焦的柳枝画眉。',
        "image": "image_4.jpg",
    },
    {
        "layout": TITLE_AND_CONTENT_LAYOUT,
        "title": "总结与思考",
        "content": '古代化妆品多取自天然，体现了古人的智慧。\n但部分原料有毒，也反映了当时的科技局限。',
    },
    {
        "layout": TITLE_SLIDE_LAYOUT,
        "title": "感谢观看",
    },
]

# 建议：添加类型提示，提高代码清晰度和健壮性
def create_slide(prs: Presentation, slide_info: Dict[str, Any]) -> None:
    """根据提供的信息创建一张幻灯片。"""
    slide_layout = prs.slide_layouts[slide_info["layout"]]
    slide = prs.slides.add_slide(slide_layout)

    if "title" in slide_info:
        slide.shapes.title.text = slide_info["title"]

    if "content" in slide_info:
        slide.placeholders[1].text = slide_info["content"]

    image_name = slide_info.get("image")
    if image_name:
        # 建议：路径拼接更可靠，并增加明确的错误处理
        image_path = os.path.join(BASE_DIR, image_name)
        if not os.path.exists(image_path):
            print(f"警告：图片文件未找到，已跳过：{image_path}", file=sys.stderr)
            return

        try:
            if slide_info.get("is_cover"):
                slide.shapes.add_picture(
                    image_path, Inches(0), Inches(0),
                    width=prs.slide_width, height=prs.slide_height
                )
            else:
                slide.shapes.add_picture(image_path, Inches(8), Inches(2), width=Inches(6))
        except Exception as e:
            print(f"错误：添加图片 '{image_path}' 时失败：{e}", file=sys.stderr)

def main() -> None:
    """主函数：校验、创建并保存演示文稿。"""
    print("开始生成PPT...")
    prs = Presentation()
    prs.slide_width = Inches(16)
    prs.slide_height = Inches(9)

    for slide_info in SLIDE_DATA:
        create_slide(prs, slide_info)

    try:
        prs.save(OUTPUT_FILENAME)
        print(f"PPT已成功生成：{OUTPUT_FILENAME}")
    except Exception as e:
        print(f"错误：保存PPT失败：{e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()