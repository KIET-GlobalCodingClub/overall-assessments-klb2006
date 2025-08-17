from PIL import Image, ImageOps

# ---------------- LOGO PATHS ---------------- #
KIET_LOGO = r"C:\Users\leela\Downloads\kiet logo.png"

# ---------------- CONFIG ---------------- #
output_logo = r"C:\Users\leela\Desktop\cert-sender\kiet_logo_auto_cropped.png"
resize_width = 300
resize_height = 150

# ---------------- PROCESSING ---------------- #
# Open the KIET logo
logo = Image.open(KIET_LOGO).convert("RGBA")

# Automatically trim any white/transparent borders
logo_cropped = ImageOps.crop(logo, border=0)

# Resize the logo
logo_resized = logo_cropped.resize((resize_width, resize_height), Image.Resampling.LANCZOS)

# Save the final logo
logo_resized.save(output_logo)

print(f"✅ KIET logo cropped and resized successfully at:\n{output_logo}")
