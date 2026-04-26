import pandas as pd
import json
import cv2
import numpy as np
from PIL import Image

def load_metadata(file_path):
    """Load metadata (variable dictionary) from Excel or JSON."""
    if file_path.endswith('.json'):
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            return pd.DataFrame(data)
    elif file_path.endswith(('.xls', '.xlsx')):
        return pd.read_excel(file_path)
    elif file_path.endswith('.csv'):
        return pd.read_csv(file_path)
    else:
        raise ValueError("Unsupported file format. Use JSON, CSV, or Excel.")

def save_spec(data, file_path):
    """Save the specification to an Excel or JSON file."""
    df = pd.DataFrame(data)
    if file_path.endswith('.json'):
        df.to_json(file_path, orient='records', force_ascii=False, indent=4)
    elif file_path.endswith(('.xls', '.xlsx')):
        df.to_excel(file_path, index=False)
    elif file_path.endswith('.csv'):
        df.to_csv(file_path, index=False, encoding='utf-8-sig')
    else:
        raise ValueError("Unsupported file format. Use JSON, CSV, or Excel.")

def draw_bboxes_on_image(image_path, ocr_results):
    """Draw bounding boxes on the original image based on OCR results."""
    # ocr_results is expected to be a list of dicts: {'bbox': [[x1,y1],[x2,y1],[x2,y2],[x1,y2]], 'text': 'abc', 'prob': 0.9}
    # easyocr format: (bbox, text, prob) where bbox is a list of 4 points [x, y]
    image = cv2.imread(image_path)
    if image is None:
        return None
    
    for item in ocr_results:
        bbox, text, prob = item
        # Get top-left and bottom-right points
        top_left = tuple([int(val) for val in bbox[0]])
        bottom_right = tuple([int(val) for val in bbox[2]])
        
        cv2.rectangle(image, top_left, bottom_right, (0, 255, 0), 2)
        
        # Optionally draw text
        # cv2.putText(image, text, top_left, cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 0, 255), 1)

    # Convert BGR to RGB for UI
    image = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
    return image
