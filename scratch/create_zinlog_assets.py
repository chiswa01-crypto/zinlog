import os
from PIL import Image, ImageDraw, ImageFont

img_dir = r"d:\new project\export_tools_app\app\static\images"
os.makedirs(img_dir, exist_ok=True)

# 1. Create horizontal SVG logo for Sidebar & Header (Both dark & light mode friendly)
svg_horizontal = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 460 110" width="100%" height="100%">
  <defs>
    <linearGradient id="goldGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#dfbc7a" />
      <stop offset="50%" stop-color="#f5d799" />
      <stop offset="100%" stop-color="#b88b4a" />
    </linearGradient>
    <linearGradient id="silverGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#ffffff" />
      <stop offset="50%" stop-color="#cbd5e1" />
      <stop offset="100%" stop-color="#64748b" />
    </linearGradient>
    <linearGradient id="emeraldDark" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#0a2318" />
      <stop offset="100%" stop-color="#061710" />
    </linearGradient>
    <linearGradient id="sageText" x1="0%" y1="0%" x2="100%" y2="0%">
      <stop offset="0%" stop-color="#164e32" />
      <stop offset="100%" stop-color="#0d2818" />
    </linearGradient>
    <filter id="subtleGlow" x="-20%" y="-20%" width="140%" height="140%">
      <feDropShadow dx="0" dy="2" stdDeviation="3" flood-color="#000000" flood-opacity="0.15"/>
    </filter>
  </defs>

  <!-- ICON EMBLEM -->
  <g transform="translate(10, 5)">
    <!-- Outer Mandala Circle Pattern -->
    <circle cx="50" cy="50" r="45" fill="none" stroke="url(#goldGrad)" stroke-width="1.2" opacity="0.6"/>
    <circle cx="50" cy="50" r="38" fill="none" stroke="url(#goldGrad)" stroke-width="0.9" opacity="0.45"/>
    <!-- Mandala Petals -->
    <g stroke="url(#goldGrad)" stroke-width="1" fill="none" opacity="0.55">
      <path d="M50,8 C62,25 62,75 50,92 C38,75 38,25 50,8 Z"/>
      <path d="M8,50 C25,62 75,62 92,50 C75,38 25,38 8,50 Z"/>
      <path d="M20,20 C42,32 68,68 80,80 C68,68 32,42 20,20 Z"/>
      <path d="M80,20 C68,42 32,68 20,80 C32,68 68,32 80,20 Z"/>
    </g>

    <!-- Stylized Metallic 'Z' Symbol -->
    <!-- Top horizontal bar -->
    <path d="M28,30 L74,30 C76,30 78,32 76,34 L40,70 L72,70 C74,70 75,72 73,74 L26,74 C24,74 23,72 25,70 L60,34 L28,34 C26,34 26,30 28,30 Z" 
          fill="url(#silverGrad)" filter="url(#subtleGlow)" stroke="#0d2818" stroke-width="0.8"/>
    <path d="M30,32 L72,32 L36,68 L70,68" fill="none" stroke="#ffffff" stroke-width="1.2" opacity="0.8"/>

    <!-- Package / Box in center -->
    <g transform="translate(24, 38)">
      <polygon points="12,0 24,6 12,12 0,6" fill="#c49a6c" stroke="#8a6336" stroke-width="0.8"/>
      <polygon points="0,6 12,12 12,24 0,18" fill="#a87f52" stroke="#8a6336" stroke-width="0.8"/>
      <polygon points="12,12 24,6 24,18 12,24" fill="#b88c5d" stroke="#8a6336" stroke-width="0.8"/>
      <!-- Seal -->
      <circle cx="6" cy="15" r="3.2" fill="#80b996" stroke="#4d7c60" stroke-width="0.6"/>
      <text x="6" y="16.5" font-family="'Plus Jakarta Sans', Arial" font-size="3.5" font-weight="900" fill="#ffffff" text-anchor="middle">z</text>
    </g>

    <!-- Upward Golden Trade Arrow -->
    <path d="M58,68 Q68,54 75,42 L79,48 L78,36 L66,37 L71,42 Q64,52 56,66 Z" 
          fill="url(#goldGrad)" filter="url(#subtleGlow)"/>
  </g>

  <!-- TEXT TYPOGRAPHY -->
  <g transform="translate(125, 20)">
    <!-- Brand Wordmark -->
    <text x="0" y="46" font-family="'Outfit', 'Plus Jakarta Sans', sans-serif" font-size="52" font-weight="900" 
          letter-spacing="2.5" fill="#0d2818">ZINLOG</text>
    
    <!-- Slogan / Subtitle -->
    <text x="2" y="70" font-family="'Plus Jakarta Sans', 'Outfit', sans-serif" font-size="11.5" font-weight="700" 
          letter-spacing="3.2" fill="#2d6a4f">SMART LOGISTICS FOR GLOBAL TRADE</text>
  </g>
</svg>
"""

# 2. Create full dark-mode hero SVG logo for Login & Welcome Banners
svg_dark_hero = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 650 160" width="100%" height="100%">
  <defs>
    <linearGradient id="goldGradHero" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#dfbc7a" />
      <stop offset="50%" stop-color="#fef0cd" />
      <stop offset="100%" stop-color="#b88b4a" />
    </linearGradient>
    <linearGradient id="silverGradHero" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#ffffff" />
      <stop offset="40%" stop-color="#d1fae5" />
      <stop offset="100%" stop-color="#94a3b8" />
    </linearGradient>
    <linearGradient id="sageTextHero" x1="0%" y1="0%" x2="100%" y2="0%">
      <stop offset="0%" stop-color="#a3cfbb" />
      <stop offset="100%" stop-color="#d1fae5" />
    </linearGradient>
    <filter id="heroGlow" x="-20%" y="-20%" width="140%" height="140%">
      <feDropShadow dx="0" dy="4" stdDeviation="6" flood-color="#000000" flood-opacity="0.4"/>
      <feDropShadow dx="0" dy="0" stdDeviation="10" flood-color="#34d399" flood-opacity="0.2"/>
    </filter>
  </defs>

  <!-- ICON EMBLEM -->
  <g transform="translate(20, 10)" filter="url(#heroGlow)">
    <!-- Outer Mandala Circle Pattern -->
    <circle cx="70" cy="70" r="64" fill="none" stroke="url(#goldGradHero)" stroke-width="1.8" opacity="0.85"/>
    <circle cx="70" cy="70" r="54" fill="none" stroke="url(#goldGradHero)" stroke-width="1.2" opacity="0.65"/>
    <!-- Mandala Petals -->
    <g stroke="url(#goldGradHero)" stroke-width="1.4" fill="none" opacity="0.75">
      <path d="M70,12 C86,36 86,104 70,128 C54,104 54,36 70,12 Z"/>
      <path d="M12,70 C36,86 104,86 128,70 C104,54 36,54 12,70 Z"/>
      <path d="M28,28 C58,45 95,95 112,112 C95,95 45,58 28,28 Z"/>
      <path d="M112,28 C95,58 45,95 28,112 C45,95 95,45 112,28 Z"/>
    </g>

    <!-- Stylized Metallic 'Z' Symbol -->
    <path d="M40,42 L102,42 C105,42 107,45 105,48 L56,98 L100,98 C103,98 105,101 102,104 L38,104 C35,104 33,101 35,98 L84,48 L40,48 C37,48 37,42 40,42 Z" 
          fill="url(#silverGradHero)" stroke="#061710" stroke-width="1.2"/>
    <path d="M42,45 L100,45 L50,95 L98,95" fill="none" stroke="#ffffff" stroke-width="1.8" opacity="0.9"/>

    <!-- Package / Box in center -->
    <g transform="translate(34, 54)">
      <polygon points="17,0 34,8.5 17,17 0,8.5" fill="#d4a373" stroke="#8a6336" stroke-width="1"/>
      <polygon points="0,8.5 17,17 17,34 0,25.5" fill="#b08968" stroke="#8a6336" stroke-width="1"/>
      <polygon points="17,17 34,8.5 34,25.5 17,34" fill="#c59b76" stroke="#8a6336" stroke-width="1"/>
      <!-- Seal -->
      <circle cx="8.5" cy="21" r="4.5" fill="#34d399" stroke="#059669" stroke-width="0.8"/>
      <text x="8.5" y="23.5" font-family="'Plus Jakarta Sans', Arial" font-size="5" font-weight="900" fill="#ffffff" text-anchor="middle">z</text>
    </g>

    <!-- Upward Golden Trade Arrow -->
    <path d="M82,95 Q96,76 106,60 L112,68 L110,50 L94,52 L100,60 Q90,74 78,92 Z" 
          fill="url(#goldGradHero)"/>
  </g>

  <!-- TEXT TYPOGRAPHY -->
  <g transform="translate(180, 25)">
    <!-- Brand Wordmark -->
    <text x="0" y="65" font-family="'Outfit', 'Plus Jakarta Sans', sans-serif" font-size="74" font-weight="900" 
          letter-spacing="4" fill="url(#sageTextHero)" filter="url(#heroGlow)">ZINLOG</text>
    
    <!-- Slogan / Subtitle -->
    <text x="3" y="100" font-family="'Plus Jakarta Sans', 'Outfit', sans-serif" font-size="16.5" font-weight="700" 
          letter-spacing="4.5" fill="#a7f3d0" opacity="0.95">SMART LOGISTICS FOR GLOBAL TRADE</text>
  </g>
</svg>
"""

# Save SVG files
with open(os.path.join(img_dir, "zinlog_logo_horizontal.svg"), "w", encoding="utf-8") as f:
    f.write(svg_horizontal)

with open(os.path.join(img_dir, "zinlog_official_logo.svg"), "w", encoding="utf-8") as f:
    f.write(svg_dark_hero)

with open(os.path.join(img_dir, "zinlog_icon.svg"), "w", encoding="utf-8") as f:
    f.write(svg_horizontal)

print("SVG files generated successfully!")
