from PIL import Image, ImageDraw, ImageFont

W, H = 413, 531  # ~35x45mm at 300dpi
img = Image.new("RGB", (W, H), (232, 238, 249))
d = ImageDraw.Draw(img)
d.rectangle([2, 2, W - 3, H - 3], outline=(11, 61, 145), width=3)
try:
    f = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 42)
    f2 = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 20)
except Exception:
    f = ImageFont.load_default()
    f2 = f
d.text((W / 2, H / 2 - 30), "PHOTO", fill=(11, 61, 145), font=f, anchor="mm")
d.text((W / 2, H / 2 + 30), "replace photo.jpg", fill=(90, 90, 90), font=f2, anchor="mm")
img.save("photo_placeholder.jpg", quality=92)
print("placeholder created")
