from pathlib import Path
from PIL import Image, ImageDraw, ImageFont


def create_app_icon() -> Path:
    """Génère une icône moderne haute résolution (.ico et .png) pour l'application."""
    app_dir = Path(__file__).resolve().parent
    assets_dir = app_dir / "assets"
    assets_dir.mkdir(parents=True, exist_ok=True)

    size = 256
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Fond arrondi dégradé rouge/bordeaux (style Backblaze B2) avec bordure subtile
    margin = 12
    draw.rounded_rectangle(
        [margin, margin, size - margin, size - margin],
        radius=52,
        fill=(210, 32, 47, 255),
        outline=(255, 110, 110, 255),
        width=4,
    )

    # Ombre douce intérieure / cercle central lumineux
    draw.rounded_rectangle(
        [margin + 8, margin + 8, size - margin - 8, size - margin - 8],
        radius=44,
        fill=(228, 45, 60, 255),
    )

    # Dessin d'un nuage blanc épuré au centre
    white = (255, 255, 255, 255)
    # Cercles composant le nuage
    draw.ellipse([52, 108, 124, 180], fill=white)
    draw.ellipse([88, 74, 176, 162], fill=white)
    draw.ellipse([140, 102, 208, 174], fill=white)
    draw.rounded_rectangle([68, 128, 194, 180], radius=22, fill=white)

    # Flèche de téléversement (Upload) rouge foncé au cœur du nuage
    arrow_color = (185, 24, 38, 255)
    # Tige de la flèche
    draw.rounded_rectangle([118, 122, 142, 170], radius=6, fill=arrow_color)
    # Pointe de la flèche
    draw.polygon([(130, 92), (98, 128), (162, 128)], fill=arrow_color)

    # Petit badge "B2" en bas à droite
    badge_bg = (24, 52, 94, 255)
    draw.rounded_rectangle([148, 168, 226, 220], radius=14, fill=badge_bg, outline=white, width=3)

    try:
        font = ImageFont.truetype("arialbd.ttf", 30)
    except Exception:
        font = ImageFont.load_default()

    bbox = draw.textbbox((0, 0), "B2", font=font)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    draw.text(
        (148 + (78 - tw) / 2, 168 + (52 - th) / 2 - 3),
        "B2",
        fill=white,
        font=font,
    )

    png_path = assets_dir / "app_icon.png"
    ico_path = assets_dir / "app_icon.ico"

    img.save(png_path, format="PNG")
    img.save(
        ico_path,
        format="ICO",
        sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)],
    )
    print(f"Icône générée : {ico_path}")
    return ico_path


if __name__ == "__main__":
    create_app_icon()
