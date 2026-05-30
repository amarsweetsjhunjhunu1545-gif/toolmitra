import sys
from PIL import Image

def process_logo(img_path, full_logo_path, favicon_path):
    img = Image.open(img_path).convert("RGBA")
    data = img.getdata()
    
    new_data = []
    for item in data:
        # Background is white or close to white
        if item[0] > 230 and item[1] > 230 and item[2] > 230:
            # Smooth anti-aliasing to transparent
            avg = (item[0] + item[1] + item[2]) / 3
            alpha = int(max(0, min(255, (255 - avg) * (255 / 25.0))))
            if alpha == 0:
                new_data.append((255, 255, 255, 0))
            else:
                new_data.append((item[0], item[1], item[2], alpha))
        else:
            new_data.append(item)
            
    img.putdata(new_data)
    
    # Save the transparent full logo
    img.save(full_logo_path, "PNG")
    
    # Create favicon by cropping the left part (the icon)
    # The image is 448x297. The icon is usually square on the left.
    # Let's crop a 297x297 square from the left.
    fav = img.crop((0, 0, 297, 297))
    
    # Trim transparent borders from favicon
    bbox = fav.getbbox()
    if bbox:
        fav = fav.crop(bbox)
        
    fav = fav.resize((512, 512), Image.Resampling.LANCZOS)
    fav.save(favicon_path, "PNG")

if __name__ == "__main__":
    in_img = r"orbixapdftool-logo.png"
    out1 = r"orbixa-logo.png"
    out2 = r"favicon.png"
    process_logo(in_img, out1, out2)
    print("Done processing user logo")
