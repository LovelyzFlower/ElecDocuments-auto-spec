import easyocr
import threading

class OCREngine:
    def __init__(self):
        # Initialize easyocr for Korean and English
        # Use GPU if available (easyocr handles it, but we can set gpu=True)
        self.reader = None
        self._initialize_reader()

    def _initialize_reader(self):
        # Loading the model might take some time, doing it in init for now
        # Ideally, this can be moved to a background thread if it blocks UI
        try:
            self.reader = easyocr.Reader(['ko', 'en'], gpu=True)
        except Exception as e:
            # Fallback if GPU fails (e.g. issues with MPS or missing CUDA)
            print(f"GPU initialization failed, falling back to CPU: {e}")
            self.reader = easyocr.Reader(['ko', 'en'], gpu=False)

    def extract_text(self, image_path):
        """
        Extract text and bounding boxes from an image.
        Returns a list of tuples: (bbox, text, prob)
        """
        if not self.reader:
            raise Exception("OCR Reader is not initialized.")
        
        # detail=1 returns bounding boxes, text, and confidence
        results = self.reader.readtext(image_path, detail=1)
        return results
