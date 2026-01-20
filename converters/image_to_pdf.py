from PIL import Image

def images_to_pdf(images_list, output_path):
    """Convert a list of images to a single PDF."""
    if not images_list:
        return

    imgs = []
    for img_path in images_list:
        img = Image.open(img_path)
        if img.mode != "RGB":
            img = img.convert("RGB")
        imgs.append(img)

    # Save all images as a single PDF
    imgs[0].save(output_path, save_all=True, append_images=imgs[1:])

def compress_image(input_path, output_path, quality=50):
    """
    Compresses an image to reduce size.
    Args:
        input_path: Path to original image.
        output_path: Path to save compressed image.
        quality: JPEG quality (1-95). Lower = more compression.
    """
    img = Image.open(input_path)
    img = img.convert("RGB")  # Ensure compatibility
    img.save(output_path, "JPEG", quality=quality)
