from PIL import Image
import os

def create_icns():
    if not os.path.exists('icon.png'):
        print("icon.png nicht gefunden!")
        return
    
    # Erstelle temporäres Verzeichnis
    if not os.path.exists('icon.iconset'):
        os.makedirs('icon.iconset')
    
    # Größen für macOS Icons
    sizes = [16, 32, 64, 128, 256, 512, 1024]
    
    # Öffne das Original-Bild
    img = Image.open('icon.png')
    
    # Erstelle die verschiedenen Größen
    for size in sizes:
        icon = img.resize((size, size), Image.Resampling.LANCZOS)
        icon.save(f'icon.iconset/icon_{size}x{size}.png')
        # Erstelle auch @2x Version
        if size <= 512:
            icon = img.resize((size * 2, size * 2), Image.Resampling.LANCZOS)
            icon.save(f'icon.iconset/icon_{size}x{size}@2x.png')
    
    # Konvertiere zu .icns mit iconutil
    os.system('iconutil -c icns icon.iconset')
    
    # Aufräumen
    os.system('rm -rf icon.iconset')

if __name__ == '__main__':
    create_icns() 