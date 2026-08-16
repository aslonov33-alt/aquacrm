from PIL import Image, ImageDraw

BLUE = (0x1e, 0x88, 0xe5)
WHITE = (255,255,255)

def make_drop_image(size, outpath):
    scale = 4
    W = size*scale
    img = Image.new('RGBA', (W,W), BLUE)
    mask = Image.new('RGBA', (W,W), (0,0,0,0))
    d = ImageDraw.Draw(mask)
    # large bulb ellipse
    d.ellipse([W*0.15, W*0.28, W*0.85, W*0.95], fill=WHITE)
    # small top ellipse
    d.ellipse([W*0.38, W*0.02, W*0.62, W*0.42], fill=WHITE)
    # Optional: a small inner lighter highlight
    hi = Image.new('RGBA', (W,W), (0,0,0,0))
    dh = ImageDraw.Draw(hi)
    dh.ellipse([W*0.45, W*0.08, W*0.65, W*0.28], fill=(255,255,255,60))
    # Composite mask onto background
    img.paste(mask, (0,0), mask)
    img.paste(hi, (0,0), hi)
    # Downscale for smoothing
    img = img.resize((size,size), Image.LANCZOS)
    img.save(outpath, format='PNG')

if __name__ == '__main__':
    make_drop_image(192, 'icon-192.png')
    make_drop_image(512, 'icon-512.png')
    print('Generated icon-192.png and icon-512.png')
