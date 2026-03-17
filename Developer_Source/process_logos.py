import os
from PIL import Image

try:
    from rembg import remove
except ImportError:
    print("rembg not installed yet")
    exit(1)

def process_logo(input_path, output_path, max_size=140):
    try:
        print(f"Processing {input_path}...")
        with open(input_path, 'rb') as i:
            with open(output_path, 'wb') as o:
                input_data = i.read()
                output_data = remove(input_data)
                o.write(output_data)
        
        # Open and resize to maintain quality
        img = Image.open(output_path)
        img.thumbnail((max_size, max_size), Image.Resampling.LANCZOS)
        img.save(output_path, "PNG")
        print(f"Saved {output_path} successfully.")
    except Exception as e:
        print(f"Error processing {input_path}: {e}")

if __name__ == "__main__":
    # Image 1 -> Right
    img1_path = r"C:\Users\gitga\.gemini\antigravity\brain\db6f4385-cb82-49e6-abf6-051b6d9c02ca\media__1772607323029.jpg"
    out1_path = "logo_right.png"
    
    # Image 2 -> Left
    img2_path = r"C:\Users\gitga\.gemini\antigravity\brain\db6f4385-cb82-49e6-abf6-051b6d9c02ca\media__1772607323148.jpg"
    out2_path = "logo_left.png"

    process_logo(img2_path, out2_path)
    process_logo(img1_path, out1_path)
