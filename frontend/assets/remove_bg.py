import sys
try:
    from PIL import Image, ImageFilter
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "Pillow"])
    from PIL import Image, ImageFilter

def remove_white_bg(img_path, out_path1, out_path2):
    img = Image.open(img_path).convert("RGBA")
    data = img.getdata()
    
    new_data = []
    for item in data:
        # White background removal with simple anti-aliasing approximation
        # Background in generated image is #FFFFFF
        if item[0] > 230 and item[1] > 230 and item[2] > 230:
            avg = (item[0] + item[1] + item[2]) / 3
            # smooth alpha falloff between 230 and 255
            # 255 -> 0 alpha, 230 -> 255 alpha
            alpha = int(max(0, min(255, (255 - avg) * (255 / 25.0))))
            if alpha == 0:
                new_data.append((255, 255, 255, 0))
            else:
                new_data.append((item[0], item[1], item[2], alpha))
        else:
            new_data.append(item)
            
    img.putdata(new_data)
    
    # Crop to bounding box
    bbox = img.getbbox()
    if bbox:
        # Add a little padding
        padding = 10
        bbox = (max(0, bbox[0]-padding), max(0, bbox[1]-padding), min(img.width, bbox[2]+padding), min(img.height, bbox[3]+padding))
        img = img.crop(bbox)
        
    img.save(out_path1, "PNG")
    
    # Resize for favicon (optional but good practice)
    fav = img.resize((512, 512), Image.Resampling.LANCZOS)
    fav.save(out_path2, "PNG")

if __name__ == "__main__":
    in_img = r"C:\Users\risha\.gemini\antigravity\brain\6a40041c-22d1-4317-b3de-cdaf1c3a5b36\3d_logo_1779635311207.png"
    out1 = r"c:\Users\risha\OneDrive\Desktop\full_website_orbixa_logo_updated\full website\frontend\assets\orbixa-logo.png"
    out2 = r"c:\Users\risha\OneDrive\Desktop\full_website_orbixa_logo_updated\full website\frontend\assets\favicon.png"
    remove_white_bg(in_img, out1, out2)
    print("Done")
