import os
import numpy as np
from PIL import Image, ImageFilter, ImageDraw

src_jpg = r'C:\Users\laptop\.gemini\antigravity-ide\brain\8c0f60c8-9c12-4e82-b66e-f1312ab064ed\.user_uploaded\media_1786862680773.jpg'
dest_dir = r'd:\new project\export_tools_app\app\static\images'

im = Image.open(src_jpg).convert("RGBA")
w, h = im.size
arr = np.array(im, dtype=np.float32)

r = arr[:, :, 0]
g = arr[:, :, 1]
b = arr[:, :, 2]
brightness = 0.299 * r + 0.587 * g + 0.114 * b

# 1. Determine geometry:
cx = w / 2.0
cy = h * 0.38
radius = w * 0.44  # Radius covering the outer gold petals

print(f"Image size: {w}x{h}, center: ({cx}, {cy}), radius: {radius}")

# Create geometric boundary mask (Mandala circle + Text rectangle at bottom)
geo_mask = Image.new("L", (w, h), 0)
draw = ImageDraw.Draw(geo_mask)

# Circle for mandala
draw.ellipse([cx - radius, cy - radius, cx + radius, cy + radius], fill=255)

# Rectangle for text at the bottom (ZINLOG and SMART LOGISTICS)
draw.rectangle([w * 0.05, h * 0.72, w * 0.95, h * 0.98], fill=255)

geo_arr = np.array(geo_mask, dtype=np.float32) / 255.0

# 2. Foreground element detection
# - ZINLOG text & subtitle:
is_text = (brightness > 80) & (g > 68)
# - Gold lines & arrows & box:
is_warm_gold = (r > g + 4) & (r > 55)
# - Metallic silver Z:
is_silver = (brightness > 75) & (np.abs(r - g) < 35) & (np.abs(g - b) < 35)
# - Mandala gold petals:
is_gold = (r > 52) & (g > 48) & (r > b + 6)
# - Green seal:
is_seal = (g > 115) & (g > r + 15)

fg_mask = (is_text | is_warm_gold | is_silver | is_gold | is_seal)

# 3. Alpha calculation:
alpha = np.zeros_like(brightness)
alpha[fg_mask & (geo_arr > 0.5)] = 255.0

# Soft antialiasing for edges
alpha_pil = Image.fromarray(alpha.astype(np.uint8))
alpha_smooth = alpha_pil.filter(ImageFilter.GaussianBlur(1.0))
alpha_smooth_arr = np.array(alpha_smooth, dtype=np.float32)

# Strict multiplication with geo_mask so absolutely NO corner pixels exist
final_alpha = np.clip(alpha_smooth_arr * geo_arr, 0, 255).astype(np.uint8)

final_arr = np.array(im)
final_arr[:, :, 3] = final_alpha

# Save
out_png = os.path.join(dest_dir, 'zinlog_logo_no_bg.png')
out_png2 = os.path.join(dest_dir, 'zinlog_logo_transparent.png')

res_img = Image.fromarray(final_arr)
res_img.save(out_png, "PNG")
res_img.save(out_png2, "PNG")

print("Saved 100% clean cutout logo with zero corner boxes!")
