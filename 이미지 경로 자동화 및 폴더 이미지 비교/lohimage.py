from PIL import Image
import imagehash
import os

# Define the paths
folder_a = r"C:\Users\simoony\auto\LOH\a"
folder_b = r"C:\Users\simoony\auto\LOH\b"
output_file = r"C:\Users\simoony\auto\LOH\LOH_image.txt"

# Get list of files in both directories
files_a = set(os.listdir(folder_a))
files_b = set(os.listdir(folder_b))

# Compute intersection of both sets to find common file names
common_files = files_a.intersection(files_b)

# Open output file in write mode
with open(output_file, 'w') as f:
    # Iterate over common file names and compare their images
    for file_name in common_files:
        image_a_path = os.path.join(folder_a, file_name)
        image_b_path = os.path.join(folder_b, file_name)

        # Calculate image hash
        hash_a = imagehash.average_hash(Image.open(image_a_path))
        hash_b = imagehash.average_hash(Image.open(image_b_path))

        # Compare hashes and log the result
        if hash_a == hash_b:
            f.write(f"{file_name} 이미지가 동일함\n")
        else:
            f.write(f"{file_name} 이미지가 동일하지 않음\n")

print("비교가 완료되었습니다.")
