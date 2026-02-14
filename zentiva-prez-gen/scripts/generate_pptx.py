import sys
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# Zentiva Brand Colors
ZENTIVA_BLUE = RGBColor(12, 27, 125)   # #0C1B7D
ZENTIVA_TEAL = RGBColor(0, 169, 143)   # #00A98F

def get_layout_by_name(prs, name):
    for layout in prs.slide_layouts:
        if layout.name == name:
            return layout
    return prs.slide_layouts[1]

def remove_all_slides(prs):
    """Removes all slides from the presentation object."""
    rIds = [prs.slides._sldIdLst[i].rId for i in range(len(prs.slides))]
    for rId in rIds:
        prs.part.drop_rel(rId)
    prs.slides._sldIdLst.clear()

def fill_text_frame(shape, content):
    """Fills a shape's text frame, ensuring the placeholder text is completely replaced."""
    tf = shape.text_frame
    tf.text = "" # Clear ghost text
    
    if isinstance(content, str):
        tf.paragraphs[0].text = content
    elif isinstance(content, list):
        if not content:
            return
            
        for i, item in enumerate(content):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
                
            if isinstance(item, str):
                p.text = item
                p.level = 0
            elif isinstance(item, (list, tuple)) and len(item) >= 2:
                p.text = item[0]
                p.level = item[1]

def create_presentation(title, slides_content, output_path, template_path=None, brand_image=None):
    if template_path and os.path.exists(template_path):
        prs = Presentation(template_path)
        remove_all_slides(prs)
    else:
        prs = Presentation()

    # Title Slide
    title_layout = get_layout_by_name(prs, "3_Title Slide")
    slide = prs.slides.add_slide(title_layout)
    
    for shape in slide.placeholders:
        ph = shape.placeholder_format
        if ph.type == 3: # CENTER_TITLE
            shape.text = title
        elif ph.type == 4: # SUBTITLE
            shape.text = "Strategic Market Insights v008"
        elif ph.idx == 13 and brand_image and os.path.exists(brand_image):
            shape.insert_picture(brand_image)

    # Content Slides
    for slide_data in slides_content:
        layout_name = slide_data.get('layout', "2_Title and Content")
        layout = get_layout_by_name(prs, layout_name)
        slide = prs.slides.add_slide(layout)
        
        if slide.shapes.title:
            slide.shapes.title.text = slide_data['title']
        
        # We need to be very careful: many placeholders (idx 13, 14, 15) might exist
        # We only want to fill ONE for 'content' and ONE for 'content2'
        filled_indices = set()
        
        # First, clear EVERYTHING that has a text frame to remove "Click to add text"
        for shape in slide.placeholders:
            if shape.has_text_frame:
                shape.text_frame.text = ""

        # Now fill only the specific ones we want
        for shape in slide.placeholders:
            ph = shape.placeholder_format
            
            # Standard single column content layout usually uses idx 14 for body
            if ph.idx == 14 and 'content' in slide_data and 14 not in filled_indices:
                fill_text_frame(shape, slide_data['content'])
                filled_indices.add(14)
            # Two column content usually uses 14 and 15
            elif ph.idx == 15 and 'content2' in slide_data and 15 not in filled_indices:
                fill_text_frame(shape, slide_data['content2'])
                filled_indices.add(15)
            # Layout with picture
            elif ph.type == 18 and 'image' in slide_data:
                if os.path.exists(slide_data['image']):
                    shape.insert_picture(slide_data['image'])

    prs.save(output_path)
    print(f"Presentation saved to {output_path}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python generate_pptx.py <title> <output_path>")
        sys.exit(1)

    title = sys.argv[1]
    output_path = sys.argv[2]
    
    script_dir = os.path.dirname(os.path.abspath(__file__))
    assets_dir = os.path.join(script_dir, "..", "assets")
    title_bg = os.path.join(assets_dir, "title_bg.png")

    slides = [
        {
            "title": "Market Overview & Growth",
            "content": [
                "The global biosimilars market is projected to reach ~$49B by 2026.",
                ("17% Compound Annual Growth Rate (CAGR).", 1),
                ("Primary driver: Patent cliffs of high-value biologics.", 1),
                "Key growth engine for Zentiva's portfolio.",
                ("Targeting oncology and immunology segments.", 1)
            ],
            "layout": "2_Title and Content"
        },
        {
            "title": "Strategic Impact",
            "content": [
                "Healthcare cost containment:",
                ("Projected savings exceeding $30B annually.", 1),
                ("Improved accessibility for chronic disease patients.", 1)
            ],
            "content2": [
                "Market Challenges:",
                ("Rebate walls from PBMs.", 1),
                ("Biosimilar voids in specialized therapeutic areas.", 1)
            ],
            "layout": "17_Title and Content" 
        }
    ]
    
    template = os.path.join(assets_dir, "Brand_v001.pptx")
    create_presentation(title, slides, output_path, template if os.path.exists(template) else None, title_bg)
