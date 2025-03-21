import json
import argparse
import sys
import logging
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.xmlchemy import OxmlElement
from pptx.enum.dml import MSO_THEME_COLOR
import locale
import requests
from io import BytesIO
from datetime import datetime

# Set locale to handle Chinese characters
try:
    locale.setlocale(locale.LC_ALL, 'zh_TW.UTF-8')
except locale.Error:
    try:
        locale.setlocale(locale.LC_ALL, 'zh_TW')
    except locale.Error:
        try:
            locale.setlocale(locale.LC_ALL, '')  # Use system default locale
        except locale.Error:
            logger.warning("Could not set Chinese locale. Some characters might not display correctly.")

# Setup logging with UTF-8 encoding
logging.basicConfig(level=logging.INFO, encoding='utf-8')
logger = logging.getLogger(__name__)

def download_image(url):
    try:
        response = requests.get(url)
        response.raise_for_status()
        return BytesIO(response.content)
    except Exception as e:
        logger.error(f"Error downloading image from {url}: {str(e)}")
        return None

def set_slide_size_to_16x9(prs):
    prs.slide_width = Inches(16)
    prs.slide_height = Inches(9)

def add_title_slide(prs, title, subtitle):
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    
    # Add background shape
    background = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height
    )
    background.fill.solid()
    background.fill.fore_color.rgb = RGBColor(240, 248, 255)  # Light blue background
    background.line.fill.background()

    # Add decorative elements
    left_bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, Inches(1), prs.slide_height
    )
    left_bar.fill.solid()
    left_bar.fill.fore_color.rgb = RGBColor(30, 144, 255)
    left_bar.line.fill.background()

    # Add title
    title_box = slide.shapes.add_textbox(
        Inches(2), Inches(3), Inches(12), Inches(1.5)
    )
    title_frame = title_box.text_frame
    title_para = title_frame.add_paragraph()
    title_para.text = title
    title_para.font.size = Pt(54)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(30, 144, 255)
    title_para.alignment = PP_ALIGN.LEFT

    # Add subtitle (date)
    subtitle_box = slide.shapes.add_textbox(
        Inches(2), Inches(4.5), Inches(12), Inches(1)
    )
    subtitle_frame = subtitle_box.text_frame
    subtitle_para = subtitle_frame.add_paragraph()
    subtitle_para.text = subtitle
    subtitle_para.font.size = Pt(32)
    subtitle_para.font.color.rgb = RGBColor(100, 100, 100)
    subtitle_para.alignment = PP_ALIGN.LEFT

def add_section_slide(prs, title):
    slide = prs.slides.add_slide(prs.slide_layouts[2])
    
    # Add gradient background
    background = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height
    )
    background.fill.solid()
    background.fill.fore_color.rgb = RGBColor(245, 245, 245)
    background.line.fill.background()

    # Add title
    title_box = slide.shapes.add_textbox(
        Inches(2), Inches(3), Inches(12), Inches(1.5)
    )
    title_frame = title_box.text_frame
    title_para = title_frame.add_paragraph()
    title_para.text = title
    title_para.font.size = Pt(48)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(30, 144, 255)
    title_para.alignment = PP_ALIGN.LEFT
    
    # Add decorative line
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(2),
        Inches(4.5),
        Inches(12),
        Inches(0.05)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = RGBColor(30, 144, 255)
    line.line.fill.background()

def add_content_slide(prs, title, content, app_icon_url=None):
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    shapes = slide.shapes
    
    # Add subtle background
    background = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height
    )
    background.fill.solid()
    background.fill.fore_color.rgb = RGBColor(255, 255, 255)
    background.line.fill.background()
    
    # Add title with app icon if available
    title_left = Inches(1)
    if app_icon_url:
        try:
            icon_image = download_image(app_icon_url)
            if icon_image:
                icon = slide.shapes.add_picture(
                    icon_image,
                    Inches(1),
                    Inches(0.5),
                    height=Inches(1)
                )
                title_left = Inches(2.2)
        except Exception as e:
            logger.error(f"Error adding app icon: {str(e)}")
    
    title_box = shapes.add_textbox(
        title_left, Inches(0.5),
        Inches(13), Inches(1)
    )
    title_frame = title_box.text_frame
    title_para = title_frame.add_paragraph()
    title_para.text = title
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(30, 144, 255)
    
    content_left = Inches(1)
    content_top = Inches(1.8)
    content_width = Inches(14)
    content_height = Inches(6)
    
    content_box = shapes.add_textbox(content_left, content_top, content_width, content_height)
    tf = content_box.text_frame
    tf.word_wrap = True
    
    for key, value in content.items():
        p = tf.add_paragraph()
        p.text = f"{key}:"
        p.font.bold = True
        p.font.size = Pt(24)
        p.font.color.rgb = RGBColor(60, 60, 60)
        
        if isinstance(value, list):
            for item in value:
                p = tf.add_paragraph()
                p.text = f"• {item}"
                p.font.size = Pt(18)
                p.level = 1
                p.font.color.rgb = RGBColor(80, 80, 80)
        elif isinstance(value, dict):
            for sub_key, sub_value in value.items():
                p = tf.add_paragraph()
                p.text = f"• {sub_key}: {sub_value}"
                p.font.size = Pt(18)
                p.level = 1
                p.font.color.rgb = RGBColor(80, 80, 80)
        else:
            p = tf.add_paragraph()
            p.text = f"• {value}"
            p.font.size = Pt(18)
            p.level = 1
            p.font.color.rgb = RGBColor(80, 80, 80)

def add_comparison_slide(prs, title, apps_data):
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    shapes = slide.shapes
    
    # Add background
    background = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height
    )
    background.fill.solid()
    background.fill.fore_color.rgb = RGBColor(255, 255, 255)
    background.line.fill.background()
    
    title_shape = shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.size = Pt(36)
    title_shape.text_frame.paragraphs[0].font.color.rgb = RGBColor(30, 144, 255)
    
    # Calculate layout for comparison boxes
    box_width = Inches(4.5)
    box_height = Inches(6)
    gap = Inches(0.5)
    start_left = (prs.slide_width - (box_width * len(apps_data) + gap * (len(apps_data) - 1))) / 2
    top = Inches(2)
    
    for i, app in enumerate(apps_data):
        left = start_left + (box_width + gap) * i
        box = shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            left, top, box_width, box_height
        )
        box.fill.solid()
        box.fill.fore_color.rgb = RGBColor(245, 245, 245)
        box.line.color.rgb = RGBColor(200, 200, 200)
        
        tf = box.text_frame
        tf.word_wrap = True
        
        p = tf.add_paragraph()
        p.text = app["name"]
        p.font.bold = True
        p.font.size = Pt(20)
        p.alignment = PP_ALIGN.CENTER
        
        # Add content specific to comparison type
        add_comparison_content(tf, app)

def add_comparison_content(text_frame, app_data):
    # Add specific comparison content based on the data
    pass  # Implementation will be added based on specific comparison needs

def add_chapter_slide(prs, chapter_data):
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    shapes = slide.shapes
    
    # Add background
    background = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height
    )
    background.fill.solid()
    background.fill.fore_color.rgb = RGBColor(255, 255, 255)
    background.line.fill.background()
    
    # Add title
    title_box = shapes.add_textbox(
        Inches(1), Inches(0.5),
        Inches(14), Inches(1)
    )
    title_frame = title_box.text_frame
    title_para = title_frame.add_paragraph()
    title_para.text = chapter_data["title"]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(30, 144, 255)
    
    # Add content sections
    sections = [
        ("關鍵發現", chapter_data["keyFindings"]),
        ("數據支持", chapter_data["dataSupport"]),
        ("建議", chapter_data["recommendations"])
    ]
    
    current_top = Inches(2)
    for section_title, section_content in sections:
        # Add section title
        section_title_box = shapes.add_textbox(
            Inches(1), current_top,
            Inches(14), Inches(0.5)
        )
        section_title_frame = section_title_box.text_frame
        section_title_para = section_title_frame.add_paragraph()
        section_title_para.text = section_title
        section_title_para.font.size = Pt(24)
        section_title_para.font.bold = True
        section_title_para.font.color.rgb = RGBColor(60, 60, 60)
        
        # Add section content
        content_box = shapes.add_textbox(
            Inches(1), current_top + Inches(0.7),
            Inches(14), Inches(1.8)
        )
        content_frame = content_box.text_frame
        content_frame.word_wrap = True
        
        for item in section_content:
            p = content_frame.add_paragraph()
            p.text = f"• {item}"
            p.font.size = Pt(18)
            p.font.color.rgb = RGBColor(80, 80, 80)
        
        current_top += Inches(2.5)

def add_ending_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    
    # Add background shape
    background = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height
    )
    background.fill.solid()
    background.fill.fore_color.rgb = RGBColor(240, 248, 255)  # Light blue background
    background.line.fill.background()

    # Add decorative elements
    left_bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, Inches(1), prs.slide_height
    )
    left_bar.fill.solid()
    left_bar.fill.fore_color.rgb = RGBColor(30, 144, 255)
    left_bar.line.fill.background()

    # Add title
    title_box = slide.shapes.add_textbox(
        Inches(2), Inches(3), Inches(12), Inches(1.5)
    )
    title_frame = title_box.text_frame
    title_para = title_frame.add_paragraph()
    title_para.text = "謝謝聆聽"
    title_para.font.size = Pt(54)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(30, 144, 255)
    title_para.alignment = PP_ALIGN.LEFT

    # Add subtitle
    subtitle_box = slide.shapes.add_textbox(
        Inches(2), Inches(4.5), Inches(12), Inches(1)
    )
    subtitle_frame = subtitle_box.text_frame
    subtitle_para = subtitle_frame.add_paragraph()
    subtitle_para.text = "如有任何問題，歡迎提出討論"
    subtitle_para.font.size = Pt(32)
    subtitle_para.font.color.rgb = RGBColor(100, 100, 100)
    subtitle_para.alignment = PP_ALIGN.LEFT

def generate_competitive_analysis_ppt(input_file: str, output_file: str):
    try:
        # Load input data
        logger.info("Loading input data...")
        with open(input_file, 'r', encoding='utf-8') as f:
            data = json.load(f)

        # Create presentation with 16:9 aspect ratio
        logger.info("Creating presentation...")
        prs = Presentation()
        set_slide_size_to_16x9(prs)
        
        # Add title slide
        add_title_slide(prs, data["title"], data["date"])
        
        # Process each app
        for app in data["apps"]:
            logger.info(f"Processing app: {app['name']}")
            
            # Add section slide for app
            add_section_slide(prs, app["name"])
            
            # App Overview
            add_content_slide(
                prs,
                f"{app['name']} 概覽",
                {
                    "評分": {
                        "iOS": app["ratings"]["ios"],
                        "Android": app["ratings"]["android"]
                    },
                    "評論統計": {
                        "正面評價": f"{app['reviews']['stats']['positive']}%",
                        "負面評價": f"{app['reviews']['stats']['negative']}%"
                    },
                    "核心功能": app["features"]["core"],
                    "優勢": app["features"]["advantages"],
                    "待改進": app["features"]["improvements"]
                }
            )

            # UX Analysis
            add_content_slide(
                prs,
                f"{app['name']} 用戶體驗分析",
                {
                    "用戶體驗評分": {
                        "會員登入": f"{app['uxScores']['memberlogin']}%",
                        "搜尋功能": f"{app['uxScores']['search']}%",
                        "商品相關": f"{app['uxScores']['product']}%",
                        "結帳付款": f"{app['uxScores']['checkout']}%",
                        "客戶服務": f"{app['uxScores']['service']}%",
                        "其他": f"{app['uxScores']['other']}%"
                    },
                    "優勢": app["uxAnalysis"]["strengths"],
                    "待改進": app["uxAnalysis"]["improvements"],
                    "總結": app["uxAnalysis"]["summary"]
                }
            )

            # Review Analysis
            add_content_slide(
                prs,
                f"{app['name']} 評論分析",
                {
                    "優勢": app["reviews"]["analysis"]["advantages"],
                    "待改進": app["reviews"]["analysis"]["improvements"],
                    "總結": app["reviews"]["analysis"]["summary"]
                }
            )
        
        # Add ending slide
        add_ending_slide(prs)
        
        # Save presentation
        logger.info(f"Saving presentation to {output_file}")
        prs.save(output_file)
        logger.info("Presentation generated successfully")
        
    except Exception as e:
        logger.error(f"Error generating PPT: {str(e)}")
        raise 