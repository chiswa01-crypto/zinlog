import os
from PIL import Image, ImageDraw, ImageFont

img_dir = r"d:\new project\export_tools_app\app\static\images"

# Create a clean raster PNG for zinlog_logo_horizontal.png and zinlog_official_logo.png
width, height = 460, 110
img_hz = Image.new("RGBA", (width, height), (255, 255, 255, 0))
draw = ImageDraw.Draw(img_hz)

# Try font or default
try:
    font_bold = ImageFont.truetype("arialbd.ttf", 46)
    font_sub = ImageFont.truetype("arial.ttf", 11)
except:
    font_bold = ImageFont.load_default()
    font_sub = ImageFont.load_default()

# Save PNGs
img_hz.save(os.path.join(img_dir, "zinlog_logo_horizontal.png"))
img_hz.save(os.path.join(img_dir, "zinlog_official_logo.png"))
print("PNG stubs created!")
