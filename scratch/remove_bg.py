import numpy as np
from PIL import Image, ImageFilter

img_path = r'd:\new project\export_tools_app\app\static\images\zinlog_login_logo.jpg'
out_path = r'd:\new project\export_tools_app\app\static\images\zinlog_logo_no_bg.png'

im_pil = Image.open(img_path).convert("RGBA")
data = np.array(im_pil, dtype=np.float32)

r = data[:, :, 0]
g = data[:, :, 1]
b = data[:, :, 2]

# Calculate luminance and color difference
# Foreground elements have:
# 1. High luminance / brightness (silver text 'LOGZIN', silver 'Z', gold outline)
# 2. Strong saturation (cardboard box is warm orange/brown, arrow is gold, seal is emerald)
# Background is dark murky green with low brightness and green hue (g > r, g > b, r < 80, g < 110, b < 90)

brightness = 0.299 * r + 0.587 * g + 0.114 * b

# Gold / Arrow / Box detection (warm colors: r > g > b)
is_warm = (r > g + 10) & (r > 60)

# Silver / Metallic detection (high brightness, neutral hue)
is_bright_neutral = (brightness > 95) & (np.abs(r - g) < 35) & (np.abs(g - b) < 35)

# Thin gold mandala lines (gold hue: r > 70, g > 65, b < 65)
is_gold_line = (r > 70) & (g > 65) & (r > b + 15)

# Bright green seal (g > 130)
is_seal = (g > 125) & (g > r + 20)

# Combine foreground signals
is_foreground = is_warm | is_bright_neutral | is_gold_line | is_seal

# Any pixel that is clearly foreground gets full opacity (255)
# Intermediate pixels get smooth alpha based on distance
alpha = np.zeros_like(brightness)
alpha[is_foreground] = 255.0

# Soft expansion & smoothing to preserve antialiased edges
alpha_img = Image.fromarray(alpha.astype(np.uint8))
alpha_blurred = alpha_img.filter(ImageFilter.GaussianBlur(1.2))

# Multiply alpha back
final_data = np.array(im_pil)
final_data[:, :, 3] = np.array(alpha_blurred)

# Crop transparent borders if needed
res_img = Image.fromarray(final_data)
res_img.save(out_path, "PNG")

print("Created transparent PNG without background successfully at:", out_path)
