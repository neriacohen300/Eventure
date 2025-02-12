import os
from PIL import Image, ImageFilter, ImageDraw

def process_images(image_paths, output_folder, progress_callback=None):
    try:
        background_folder = os.path.join(output_folder, "01_תמונות", "רקעים")
        images_folder = os.path.join(output_folder, "01_תמונות", "תמונות")
        os.makedirs(background_folder, exist_ok=True)
        os.makedirs(images_folder, exist_ok=True)

        total_images = len(image_paths)
        for i, image_path in enumerate(image_paths, start=1):
            rotation = image_path['rotation']
            original_image = Image.open(image_path['path']).convert("RGBA")

            # Rotate the original image if rotation is specified
            if rotation:
                original_image = original_image.rotate(rotation, expand=True)

            # Resize
            original_aspect = original_image.width / original_image.height
            target_aspect = 1920 / 1080
            if original_aspect > target_aspect:
                new_width = 1920
                new_height = int(1920 / original_aspect)
            else:
                new_height = 1080
                new_width = int(1080 * original_aspect)
            resized_image = original_image.resize((new_width, new_height), Image.Resampling.LANCZOS)

            # Create background and final image
            final_background = Image.new("RGB", (1920, 1080))
            blurred_image = resized_image.resize((1920, 1080), Image.Resampling.LANCZOS).filter(ImageFilter.GaussianBlur(radius=5))
            final_background.paste(blurred_image, (0, 0))

            scaled_width, scaled_height = int(new_width * 0.9), int(new_height * 0.9)
            original_image_scaled = resized_image.resize((scaled_width, scaled_height), Image.Resampling.LANCZOS)
            x_offset, y_offset = (1920 - scaled_width) // 2, (1080 - scaled_height) // 2

            final_image = Image.new("RGBA", (1920, 1080))
            final_image.paste(original_image_scaled, (x_offset, y_offset), original_image_scaled)

            # Save
            final_background.save(os.path.join(background_folder, f"background_img{i}.png"), quality=85)
            final_image.save(os.path.join(images_folder, f"img{i}.png"), quality=85)

            if progress_callback:
                progress = int((i / total_images) * 100)
                progress_callback(progress)
        return True
    except Exception as e:
        print(f"Error processing images: {e}")
        return False