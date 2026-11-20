import os
import openpyxl
import json
from copy import copy
from google import genai
from dotenv import load_dotenv

load_dotenv()
gemini_api_key = os.getenv("GEMINI_API_KEY")
gemini_ai_model = os.getenv("GEMINI_MODEL", "gemini-3.5-flash-lite")

if not gemini_api_key:
    raise RuntimeError("GEMINI_API_KEY is missing. Add it to .env or export it in your environment.")
print("Gemini API Key loaded successfully.")

# Initialize the new Google GenAI client
client = genai.Client(api_key=gemini_api_key)


def extract_parcel_as_json(image_path):
    """
    Extracts 6 fields from a single parcel image and returns a Python dictionary.
    """
    # 1. Upload the image file using the latest Google GenAI SDK
    uploaded_file = client.files.upload(file=image_path)
    
    prompt = """
    You are an expert data extraction assistant for India Post shipments.
    Analyze this parcel slip image and extract these 6 fields precisely:
    1. Barcode No: Full tracking number printed below the barcode (e.g., starting with CA or CX).
    2. Physical weight: The handwritten weight number written on the side (e.g., 406, 520). Leave blank if missing.
    3. Receiver pincode: The 6-digit Indian PIN code.
    4. Receiver name: Recipient's name written next to "To".
    5. Receiver Mobile No: Primary recipient phone number.
    
    Return the result STRICTLY as a valid JSON object with these exact keys:
    {
      "Barcode": "...",
      "Weight": "...",
      "Pincode": "...",
      "Name": "...",
      "Mobile": "..."    
    }
    """
    
    # 2. Call Gemini Flash with enforced JSON output
    response = client.models.generate_content(
        model=gemini_ai_model,
        contents=[uploaded_file, prompt],
        config={
            'response_mime_type': 'application/json'
        }
    )
    
    # 3. Parse JSON text into a Python dictionary
    return json.loads(response.text)

def process_photo_batch(booking_date_str, image_paths_list):
    """
    Processes a list of image paths one by one, handles errors gracefully, 
    and updates the date-wise Excel file in a single batch.
    """
    extracted_data_list = []
    
    print(f"📦 Starting batch processing for {len(image_paths_list)} photos...")
    
    for idx, img_path in enumerate(image_paths_list, start=1):
        print(f"Processing photo {idx}/{len(image_paths_list)}: {img_path}")
        try:
            # Extract data for this specific photo
            parcel_data = extract_parcel_as_json(img_path)
            extracted_data_list.append(parcel_data)
            print(f"  -> Extracted Barcode: {parcel_data.get('Barcode')} | Name: {parcel_data.get('Name')}")
        except Exception as e:
            print(f"  ❌ Error processing {img_path}: {e}")
            # Optional: Add a flagged/blank record or log for manual review
            continue
            
    # Once all photos are extracted, update the Excel register once
    if extracted_data_list:
        update_excel_register(booking_date_str, extracted_data_list)
        print(f"✅ Batch successfully compiled into vani_bulk_file_{booking_date_str}.xlsx!")
    else:
        print("⚠️ No valid data extracted from this batch.")

def update_excel_register(booking_date_str, extracted_data_list, template_path="./vani_bulk_template.xlsx"):
    """
    Updates or creates the date-wise Excel file (vani_bulk_file_DD.MM.YY.xlsx),
    applies consecutive row placement, auto-increments serial numbers, 
    and retains row 2 template defaults and formatting.
    """
    output_filename = f"vani_bulk_file_{booking_date_str}.xlsx"
    
    # Load existing file or base template if it doesn't exist
    if os.path.exists(output_filename):
        wb = openpyxl.load_workbook(output_filename)
    else:
        wb = openpyxl.load_workbook(template_path)
        
    ws = wb["ArticleDetails"]
    
    
    # Map header positions dynamically
    headers = {cell.value: cell.column for cell in ws[1] if cell.value}
    col_barcode = headers.get("BARCODE NO")
    col_weight = headers.get("PHYSICAL WEIGHT")
    col_pincode = headers.get("RECEIVER PINCODE")
    col_name = headers.get("RECEIVER NAME")
    col_mobile = headers.get("RECEIVER MOBILE NO")
    col_serial = headers.get("SERIAL NUMBER")
    
    # Scan existing records to find max serial and next available row
    existing_barcodes = {}
    max_serial = 0
    next_available_row = 2
    
    for row_idx in range(2, ws.max_row + 1):
        barcode_val = ws.cell(row=row_idx, column=col_barcode).value
        if barcode_val:
            existing_barcodes[barcode_val] = row_idx
            next_available_row = max(next_available_row, row_idx + 1)
            serial_val = ws.cell(row=row_idx, column=col_serial).value
            try:
                if serial_val: 
                    max_serial = max(max_serial, int(serial_val))
            except ValueError:
                pass

    # Process each extracted parcel using UPSERT logic
    for data in extracted_data_list:
        barcode = data.get("Barcode")
        weight = data.get("Weight")
        pincode = data.get("Pincode")
        name = data.get("Name")
        mobile = data.get("Mobile")
        
        if barcode in existing_barcodes:
            # Update existing row
            target_row = existing_barcodes[barcode]
        else:
            # Insert new row consecutively
            target_row = next_available_row
            existing_barcodes[barcode] = target_row
            next_available_row += 1
            max_serial += 1
            
            # Replicate Row 2 defaults, formulas, and formatting for new rows
            for col in range(1, ws.max_column + 1):
                source_cell = ws.cell(row=2, column=col)
                target_cell = ws.cell(row=target_row, column=col)
                target_cell.value = source_cell.value
                if source_cell.has_style:
                    target_cell.font = copy(source_cell.font)
                    target_cell.border = copy(source_cell.border)
                    target_cell.fill = copy(source_cell.fill)
                    target_cell.number_format = copy(source_cell.number_format)
                    target_cell.protection = copy(source_cell.protection)
                    target_cell.alignment = copy(source_cell.alignment)
            
            # Assign auto-incremented serial number
            if col_serial:
                ws.cell(row=target_row, column=col_serial, value=max_serial)

        # Populate the 6 parcel fields
        if col_barcode and barcode: ws.cell(row=target_row, column=col_barcode, value=barcode)
        if col_weight and weight: ws.cell(row=target_row, column=col_weight, value=int(weight) if str(weight).isdigit() else weight)
        if col_pincode and pincode: ws.cell(row=target_row, column=col_pincode, value=str(pincode))
        if col_name and name: ws.cell(row=target_row, column=col_name, value=name)
        if col_mobile and mobile: ws.cell(row=target_row, column=col_mobile, value=str(mobile))

    wb.save(output_filename)
    print(f"Successfully updated {output_filename}")

IMAGE_FOLDER = "./images/211126"
BATCH_SIZE = 10

image_paths = [
    os.path.join(IMAGE_FOLDER, filename)
    for filename in sorted(os.listdir(IMAGE_FOLDER))
        if filename.lower().endswith((".jpeg", ".jpg", ".png"))
]

# Store the photo paths as arrays of up to 10 strings.
image_batches = [
    image_paths[start:start + BATCH_SIZE]
    for start in range(0, len(image_paths), BATCH_SIZE)
]

# Failed photo paths from the previous batch runs.
failed_image_paths = [
    "./images/211126/Lakshmi.jpeg"
]

# Batch the failed paths; each batch contains at most 10 photos.
failed_image_batches = [
    failed_image_paths[start:start + BATCH_SIZE]
    for start in range(0, len(failed_image_paths), BATCH_SIZE)
]

for failed_batch in failed_image_batches:
    process_photo_batch("21.11.26", failed_batch)