from PIL import Image, ImageChops

# Load the logo
logo_path = r"C:\Users\leela\Downloads\iiit logo.jpeg"
logo = Image.open(logo_path)

# Convert to RGBA if not already, to handle transparency
logo = logo.convert("RGBA")

# Automatically crop out empty/transparent background
bg = Image.new(logo.mode, logo.size, (255, 255, 255, 0))  # transparent background
diff = ImageChops.difference(logo, bg)
bbox = diff.getbbox()  # get bounding box of non-background content

if bbox:
    cropped_logo = logo.crop(bbox)
else:
    cropped_logo = logo  # if no bbox found, keep original

# Optional: Resize the logo
resized_logo = cropped_logo.resize((300, 150))  # adjust size as needed

# Save the processed logo
resized_logo.save("iiit_logo_processed.png")

print("✅ Logo cropped and saved successfully!")
