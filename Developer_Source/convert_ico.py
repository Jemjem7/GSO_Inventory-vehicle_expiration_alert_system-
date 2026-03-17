import sys
from PIL import Image
import os

def convert_png_to_ico(png_path, ico_path):
    try:
        img = Image.open(png_path)
        # Ensure RGBA if it has transparency
        if img.mode != 'RGBA':
            img = img.convert('RGBA')
        
        # Resize to standard icon sizes
        icon_sizes = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (255, 255)]
        img.save(ico_path, format='ICO', sizes=icon_sizes)
        print(f"Successfully converted {png_path} to {ico_path}")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python convert_ico.py <input_png> <output_ico>")
    else:
        convert_png_to_ico(sys.argv[1], sys.argv[2])
