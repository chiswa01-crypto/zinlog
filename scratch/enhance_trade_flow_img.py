import os
from PIL import Image, ImageFilter, ImageEnhance

src_img = r'C:\Users\laptop\.gemini\antigravity-ide\brain\8c0f60c8-9c12-4e82-b66e-f1312ab064ed\.user_uploaded\media_1786876614400.png'
im = Image.open(src_img).convert('RGBA')
print('Raw upload size:', im.size)

# Crop border/card if needed, or if it's already the card:
# Let's inspect borders:
w, h = im.size

# Check if there are outer background margins
# Crop inner content cleanly (remove any outer browser frame padding if present)
# Let's crop slightly inside the card border
crop_box = (int(w * 0.015), int(h * 0.02), int(w * 0.985), int(h * 0.98))
cropped = im.crop(crop_box)

# High Quality 2x Upscaling with Lanczos filter for crisp rendering
target_w = cropped.width * 2
target_h = cropped.height * 2
upscaled = cropped.resize((target_w, target_h), Image.Resampling.LANCZOS)

# Apply Unsharp Mask & Sharpness enhancement for crystal clear text and 3D edges
sharpened = upscaled.filter(ImageFilter.UnsharpMask(radius=2, percent=130, threshold=3))

# Subtle contrast enhancement
enhancer = ImageEnhance.Contrast(sharpened)
enhanced = enhancer.enhance(1.05)

# Subtle color boost for vibrant logistics theme
color_enhancer = ImageEnhance.Color(enhanced)
final_img = color_enhancer.enhance(1.04)

dest_path = r'd:\new project\export_tools_app\app\static\images\zinlog_trade_flow_illustration.png'
final_img.save(dest_path, 'PNG', optimize=True)
print(f'Enhanced HD Image saved to {dest_path} with resolution: {final_img.size}')
