import numpy as np
from PIL import Image, ImageFilter

img_path = r'd:\new project\export_tools_app\app\static\images\zinlog_login_logo.jpg'
out_png = r'd:\new project\export_tools_app\app\static\images\zinlog_login_transparent.png'

im = Image.open(img_path).convert("RGBA")
w, h = im.size

# Create a smooth elliptical mask feathering towards edges
mask = Image.new("L", (w, h), 0)
from PIL import ImageDraw
draw = ImageDraw.Draw(mask)

# Draw rounded rectangle or oval mask with smooth blur
pad_x = 40
pad_y = 40
draw.rounded_rectangle([pad_x, pad_y, w - pad_x, h - pad_y], radius=160, fill=255)

# Gaussian blur the mask for ultra-smooth edge feathering
feathered_mask = mask.filter(ImageFilter.GaussianBlur(35))

# Apply feathered mask as alpha channel
r, g, b, a = im.split()
result_im = Image.merge("RGBA", (r, g, b, feathered_mask))
result_im.save(out_png, "PNG")

print("Generated seamlessly feathered transparent logo at:", out_png)
