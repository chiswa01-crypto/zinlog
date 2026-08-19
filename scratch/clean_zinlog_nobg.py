import numpy as np
from PIL import Image, ImageFilter, ImageOps
import shutil
import os

src_jpg = r'C:\Users\laptop\.gemini\antigravity-ide\brain\8c0f60c8-9c12-4e82-b66e-f1312ab064ed\.user_uploaded\media_1786862680773.jpg'
dest_dir = r'd:\new project\export_tools_app\app\static\images'

# 1. Copy as master JPG
shutil.copy2(src_jpg, os.path.join(dest_dir, 'zinlog_master.jpg'))

# 2. Open image and process alpha matting
im = Image.open(src_jpg).convert("RGBA")
arr = np.array(im, dtype=np.float32)

r = arr[:, :, 0]
g = arr[:, :, 1]
b = arr[:, :, 2]

# Measure brightness
brightness = 0.299 * r + 0.587 * g + 0.114 * b

# Background characteristics:
# In the outer areas and between mandala spaces, the background is dark green
# (g is slightly higher than r and b, but overall intensity is low, brightness < 75)
# In the center background behind Z, brightness is around 40-70 and green hue dominant

# Foreground elements:
# - ZINLOG text at bottom: Sage green text with high brightness (g > 110, r > 80, b > 90) or brightness > 90
# - SMART LOGISTICS text: gold/cream text (r > 90, g > 85)
# - Gold mandala frame: gold lines (r > 65, g > 60, r > b + 10)
# - Silver 3D 'Z': metallic gradient (brightness > 80, neutral hue)
# - Delivery box: warm cardboard (r > g + 10, r > 70)
# - Arrow: gold (r > g, r > 90)
# - Seal: bright emerald (g > 120)

# Build a continuous alpha mask based on color distance from background
# Background color model:
# Dark green hue: normalized vector ~ [0.15, 0.45, 0.25]
bg_color = np.array([10.0, 38.0, 24.0]) # dark forest green

# Calculate Euclidean distance from dark background
diff = np.sqrt((r - bg_color[0])**2 + (g - bg_color[1])**2 + (b - bg_color[2])**2)

# Specific foreground detections:
is_text_zinlog = (brightness > 85) & (g > 75) # catches ZINLOG letters & SMART LOGISTICS
is_mandala_gold = (r > 60) & (g > 55) & (r > b + 5)
is_box_arrow = (r > g + 5) & (r > 60)
is_silver_z = (brightness > 75) & (np.abs(r - g) < 40) & (np.abs(g - b) < 40)
is_seal = (g > 115) & (g > r + 15)

fg_mask = is_text_zinlog | is_mandala_gold | is_box_arrow | is_silver_z | is_seal

# Compute smooth alpha:
# For pixels matching foreground signals, alpha is high (200-255)
# For background pixels, alpha scales down smoothly
alpha = np.zeros_like(brightness)

# Threshold mapping
alpha[fg_mask] = 255.0

# Intermediate threshold for soft antialiasing
inter = (~fg_mask) & (diff > 45) & (brightness > 55)
alpha[inter] = np.clip((diff[inter] - 45) * 6.0, 0, 220)

# Smooth edges with Gaussian filter
alpha_pil = Image.fromarray(alpha.astype(np.uint8))
alpha_smooth = alpha_pil.filter(ImageFilter.GaussianBlur(0.8))

# Assemble final image
final_arr = np.array(im)
final_arr[:, :, 3] = np.array(alpha_smooth)

out_png = os.path.join(dest_dir, 'zinlog_logo_transparent.png')
out_login_png = os.path.join(dest_dir, 'zinlog_logo_no_bg.png')

result_img = Image.fromarray(final_arr)
result_img.save(out_png, "PNG")
result_img.save(out_login_png, "PNG")

print("Saved clean ZINLOG transparent logo at:", out_png)
