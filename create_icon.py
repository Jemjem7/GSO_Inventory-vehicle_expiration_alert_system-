from PIL import Image, ImageDraw

def create_image():
    # Create a simple icon image for the system tray
    width = 64
    height = 64
    image = Image.new('RGB', (width, height), (255, 255, 255))
    dc = ImageDraw.Draw(image)
    dc.rectangle((width // 4, height // 4, width * 3 // 4, height * 3 // 4), fill=(0, 120, 215)) # Blue square
    dc.text((width // 3 + 2, height // 3 + 5), "V", fill=(255, 255, 255))
    return image

if __name__ == '__main__':
    img = create_image()
    img.save('app_icon.ico', format='ICO', sizes=[(64, 64)])
    print("app_icon.ico created successfully.")
