"""Render cover_1280x720.png. Numbers shown are the SAMPLE data (labelled as such)."""
from PIL import Image, ImageDraw, ImageFont
import sys

out = sys.argv[1]
W, H = 1280, 720
NAVY, CREAM, GREEN, GREY, WHITE = (31, 58, 95), (255, 248, 220), (46, 125, 50), (90, 90, 90), (255, 255, 255)
B = "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"
R = "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"
f = lambda p, s: ImageFont.truetype(p, s)

img = Image.new("RGB", (W, H), WHITE)
d = ImageDraw.Draw(img)
d.rectangle([0, 0, 540, H], fill=NAVY)

# left panel
d.text((50, 90), "FREELANCER", font=f(B, 30), fill=(170, 200, 235))
d.text((50, 140), "Income &", font=f(B, 64), fill=WHITE)
d.text((50, 220), "Tax Set-Aside", font=f(B, 64), fill=WHITE)
d.text((50, 300), "Tracker", font=f(B, 64), fill=WHITE)
d.text((50, 410), "Know what's really yours", font=f(R, 30), fill=CREAM)
d.text((50, 450), "after expenses and taxes.", font=f(R, 30), fill=CREAM)
for i, t in enumerate(["Excel  •  Google Sheets  •  LibreOffice", "No macros  •  No subscription"]):
    d.text((50, 560 + i * 36), t, font=f(R, 22), fill=(200, 215, 235))

# right: mock summary table (sample data)
x0, y0 = 580, 110
d.text((x0, 60), "Monthly Summary", font=f(B, 34), fill=NAVY)
cols = [("Month", 0), ("Income", 120), ("Expenses", 270), ("Profit", 440), ("Set aside", 570)]
d.rectangle([x0, y0, x0 + 680, y0 + 48], fill=NAVY)
for name, dx in cols:
    d.text((x0 + 12 + dx * 0.93, y0 + 11), name, font=f(B, 24), fill=WHITE)
rows = [("Jan", "1,976", "60", "1,916", "479"), ("Feb", "2,082", "90", "1,992", "498"),
        ("Mar", "1,940", "150", "1,790", "447.5")]
for r, row in enumerate(rows):
    y = y0 + 48 + r * 62
    d.rectangle([x0, y, x0 + 680, y + 62], fill=CREAM if r % 2 == 0 else WHITE, outline=(200, 200, 200))
    for (name, dx), v in zip(cols, row):
        d.text((x0 + 12 + dx * 0.93, y + 16), v, font=f(R, 26), fill=(30, 30, 30))
y = y0 + 48 + 3 * 62
d.rectangle([x0, y, x0 + 680, y + 62], fill=(242, 242, 242), outline=(200, 200, 200))
tot = ("Total", "5,998", "300", "5,698", "1,424.5")
for (name, dx), v in zip(cols, tot):
    d.text((x0 + 12 + dx * 0.93, y + 16), v, font=f(B, 26), fill=NAVY)

# feature bullets
feats = ["Monthly income, expenses & profit", "Set-aside % you choose", "Editable expense categories", "500 rows each, reusable every year"]
for i, t in enumerate(feats):
    yy = 420 + i * 48
    d.ellipse([x0, yy + 8, x0 + 20, yy + 28], fill=GREEN)
    d.text((x0 + 36, yy + 2), t, font=f(R, 27), fill=(30, 30, 30))

d.text((x0, 650), "Sample data shown. Not tax advice; set-aside rate is your own planning figure.", font=f(R, 17), fill=GREY)
img.save(out)
print("saved", out, img.size)
