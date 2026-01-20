from PIL import Image, ImageDraw

def create_icon(filename, color, text):
    # Create a 256x256 image with RGBA mode
    img = Image.new('RGBA', (256, 256), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    
    # Draw a rounded rectangle or circle as background
    # Using a simple rectangle for reliability
    draw.rounded_rectangle([(10, 10), (246, 246)], radius=40, fill=color, outline="white", width=5)
    
    # Simple text drawing - simulating "icon" look without font dependencies
    # We'll just draw some shapes to represent the text/function
    
    if text == "M": # Monitor
        # Draw an 'eye' or screen shape
        draw.rectangle([(50, 50), (206, 180)], outline="white", width=8) # Screen
        draw.polygon([(100, 180), (156, 180), (176, 220), (80, 220)], fill="white") # Stand
        draw.ellipse([(110, 95), (146, 135)], fill="white") # "Eye" pupil
        
    elif text == "H": # Hyperlink
        # Draw a "link" chain
        draw.ellipse([(50, 80), (130, 160)], outline="white", width=12) # Left link
        draw.ellipse([(126, 80), (206, 160)], outline="white", width=12) # Right link
    
    # Save as .ico
    img.save(filename, format='ICO', sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])
    print(f"Created {filename}")

if __name__ == "__main__":
    # Create Monitor Icon (Blue-ish)
    create_icon("monitor.ico", (30, 144, 255), "M")
    
    # Create Worker Icon (Green-ish)
    create_icon("link.ico", (46, 139, 87), "H")
