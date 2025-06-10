import pandas as pd
from PIL import Image
import pytesseract
import os

# Set the full path to tesseract.exe
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Path to the folder containing screenshots
screenshot_folder = "C:\\Users\\Darren\\Desktop\\School\\Screen Shot Extractor\\screenshots"

# Excel output file
output_file = "video_info.xlsx"

# Initialize data list
data = []

# Track processed filenames to avoid duplicates
processed_files = set()

# Check if the output Excel file exists and is valid
if os.path.exists(output_file) and os.path.getsize(output_file) > 0:
    try:
        existing_df = pd.read_excel(output_file, engine="openpyxl")
        data.extend(existing_df.to_dict('records'))
        processed_files.update(existing_df["PXL Number"].tolist())
    except Exception as e:
        print(f"Warning: Could not read existing Excel file: {e}")
        print("Creating a new one from scratch.")
else:
    # Create file with initial entry on first run
    data.append({
        "Number": 1,
        "Info": "Top 10 bump set.:08-end",
        "PXL Number": "PXL_20250524_160854612.mp4"
    })
    processed_files.add("PXL_20250524_160854612.mp4")

# Process each screenshot in the folder
for filename in os.listdir(screenshot_folder):
    if filename.endswith((".png", ".jpg", ".jpeg")) and filename not in processed_files:
        img_path = os.path.join(screenshot_folder, filename)
        img = Image.open(img_path)
        text = pytesseract.image_to_string(img)
        lines = text.splitlines()

        # Extract Info (title) and PXL Number
        info = ""
        pxl = ""
        for line in lines:
            if any(keyword in line.lower() for keyword in ["top 10", "defense", "serve", "set"]):
                if not info:
                    info = line.strip()
            if "PXL_" in line and ".mp4" in line:
                pxl = line.strip().split()[-1]

        # Append to data if both fields are found and not a duplicate
        if info and pxl and pxl not in processed_files:
            data.append({
                "Number": len(data) + 1,
                "Info": info,
                "PXL Number": pxl
            })
            processed_files.add(pxl)

# Create DataFrame and save to Excel
df = pd.DataFrame(data)
df.to_excel(output_file, index=False, engine="openpyxl")
print(f"✅ Table saved to {output_file}")
