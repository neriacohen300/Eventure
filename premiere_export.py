import os
from PIL import Image, ImageFilter, ImageDraw
import concurrent.futures


def process_images(image_paths, output_folder, progress_callback=None):
    try:
        background_folder = os.path.join(output_folder, "01_תמונות", "רקעים")
        images_folder = os.path.join(output_folder, "01_תמונות", "תמונות")
        os.makedirs(background_folder, exist_ok=True)
        os.makedirs(images_folder, exist_ok=True)

        total_images = len(image_paths)
        
        # Use ProcessPoolExecutor for CPU-bound tasks
        with concurrent.futures.ProcessPoolExecutor() as executor:
            futures = [executor.submit(process_single_image, i, img_data, background_folder, images_folder)
                       for i, img_data in enumerate(image_paths, 1)]
            
            for future in concurrent.futures.as_completed(futures):
                if progress_callback:
                    progress = int((len([f for f in futures if f.done()]) / total_images) * 100)
                    progress_callback(progress)
        
        return True
    except Exception as e:
        print(f"Error processing images: {e}")
        return False

def process_single_image(index, img_data, bg_folder, img_folder):
    rotation = img_data['rotation']
    original_image = Image.open(img_data['path']).convert("RGBA")
    
    if rotation:
        original_image = original_image.rotate(rotation, expand=True)
    
    # Resize logic (same as before)
    original_aspect = original_image.width / original_image.height
    target_aspect = 1920 / 1080
    if original_aspect > target_aspect:
        new_width = 1920
        new_height = int(1920 / original_aspect)
    else:
        new_height = 1080
        new_width = int(1080 * original_aspect)
    resized_image = original_image.resize((new_width, new_height), Image.Resampling.BILINEAR)  # Faster resampling
    
    # Background
    background_image = resized_image.resize((1920, 1080), Image.Resampling.LANCZOS)
    final_background = Image.new("RGB", (1920, 1080))
    final_background.paste(background_image, (0, 0))
    
    # Foreground
    scaled_width, scaled_height = int(new_width * 0.9), int(new_height * 0.9)
    original_image_scaled = resized_image.resize((scaled_width, scaled_height), Image.Resampling.BILINEAR)
    x_offset, y_offset = (1920 - scaled_width) // 2, (1080 - scaled_height) // 2
    final_image = Image.new("RGBA", (1920, 1080))
    final_image.paste(original_image_scaled, (x_offset, y_offset), original_image_scaled)
    
    # Save with optimized formats
    final_background.save(os.path.join(bg_folder, f"background_img{index}.jpg"), quality=85, optimize=True, subsampling=0)  # JPEG for background
    final_image.save(os.path.join(img_folder, f"img{index}.png"), optimize=True)  # Keep PNG for transparency
    
    return True