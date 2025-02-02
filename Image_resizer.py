import os
from PIL import Image, ImageFilter

def process_image(image_path, output_folder):
    try:
        # Open the image
        original_image = Image.open(image_path)

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
        original_image = original_image.resize((new_width, new_height), Image.Resampling.LANCZOS)

        # Create a new blank image with size 1920x1080
        final_image = Image.new("RGB", (1920, 1080))

        # Blur the resized image with increased blur radius for more blur
        blurred_image = original_image.resize((1920, 1080), Image.Resampling.LANCZOS).filter(ImageFilter.GaussianBlur(radius=10))

        # Paste the blurred image onto the blank image
        final_image.paste(blurred_image, (0, 0))

        # Calculate the size for the original image to be 90% of the blurred image size
        scaled_width = int(new_width * 0.9)
        scaled_height = int(new_height * 0.9)

        # Resize the original image to 90% of its size
        original_image_scaled = original_image.resize((scaled_width, scaled_height), Image.Resampling.LANCZOS)

        # Calculate the position to place the resized original image in the middle
        x_offset = (1920 - scaled_width) // 2
        y_offset = (1080 - scaled_height) // 2

        # Paste the scaled original image onto the blank image at calculated position
        final_image.paste(original_image_scaled, (x_offset, y_offset))

        # Create the output folder if it doesn't exist
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        # Save the final image to the output folder
        filename = os.path.basename(image_path)
        output_path = os.path.join(output_folder, filename)
        final_image.save(output_path)

        return output_path  # Return the new image path
    except Exception as e:
        print(f"Error processing image {image_path}: {e}")
        return None
