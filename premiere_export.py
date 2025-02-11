import os
from PIL import Image, ImageFilter, ImageDraw

def process_images(image_paths, output_folder, progress_callback=None):
    try:
        # Create output directories if they don't exist
        background_folder = os.path.join(output_folder, "01_תמונות", "רקעים")
        images_folder = os.path.join(output_folder, "01_תמונות", "תמונות")
        os.makedirs(background_folder, exist_ok=True)
        os.makedirs(images_folder, exist_ok=True)

        total_images = len(image_paths)
        for i, image_path in enumerate(image_paths, start=1):
            # Open the image
            original_image = Image.open(image_path['path']).convert("RGBA")

            # Calculate the new size while maintaining aspect ratio
            original_aspect = original_image.width / original_image.height
            target_aspect = 1920 / 1080

            if original_aspect > target_aspect:
                new_width = 1920
                new_height = int(1920 / original_aspect)
            else:
                new_height = 1080
                new_width = int(1080 * original_aspect)

            # Resize the original image to fit within 1920x1080
            resized_image = original_image.resize((new_width, new_height), Image.Resampling.LANCZOS)

            # Create a new blank image with size 1920x1080
            final_background = Image.new("RGB", (1920, 1080))
            final_image = Image.new("RGBA", (1920, 1080))

            # Blur the resized image with a reduced blur radius for faster processing
            blurred_image = resized_image.resize((1920, 1080), Image.Resampling.LANCZOS).filter(ImageFilter.GaussianBlur(radius=7))

            # Paste the blurred image onto the blank image
            final_background.paste(blurred_image, (0, 0))

            # Calculate the size for the original image to be 90% of the resized image size
            scaled_width = int(new_width * 0.9)
            scaled_height = int(new_height * 0.9)

            # Resize the original image to 90% of its size
            original_image_scaled = resized_image.resize((scaled_width, scaled_height), Image.Resampling.LANCZOS)

            # Calculate the position to place the resized original image in the middle
            x_offset = (1920 - scaled_width) // 2
            y_offset = (1080 - scaled_height) // 2

            # Paste the scaled original image onto the blank image at calculated position
            final_image.paste(original_image_scaled, (x_offset, y_offset), original_image_scaled)

            # Save the blurred background image
            background_output_path = os.path.join(background_folder, f"background_img{i}.png")
            final_background.save(background_output_path, quality=100)

            # Save the centered image with transparent background
            image_output_path = os.path.join(images_folder, f"img{i}.png")
            final_image.save(image_output_path, quality=100)

            # Call the progress callback if provided
            if progress_callback:
                progress = int((i / total_images) * 100)  # Calculate progress percentage
                progress_callback(progress)  # Emit progress update

        return True  # Return True if all images are processed successfully
    except Exception as e:
        print(f"Error processing images: {e}")
        return False