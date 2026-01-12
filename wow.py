"""
PDF End User Reader with OCR support for scanned images
Requires: pip install PyMuPDF pytesseract Pillow
Also requires Tesseract OCR installed on system
"""

import fitz  # PyMuPDF
import re
import os
from tkinter import Tk, filedialog, messagebox
import pytesseract
from PIL import Image
import io

# Configure Tesseract path (Windows)
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
if os.path.exists(TESSERACT_PATH):
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
else:
    print("⚠️ Tesseract not found at default path. OCR may not work.")
    print(f"   Expected: {TESSERACT_PATH}")
    print("   Download from: https://github.com/UB-Mannheim/tesseract/wiki\n")


def extract_end_user_from_pdf(pdf_path, max_pages=5, use_ocr=True):
    """
    Extract 'End User' value from PDF (supports scanned images via OCR)
    
    Args:
        pdf_path: Path to PDF file
        max_pages: Number of pages to search (default: 5)
        use_ocr: Enable OCR for scanned PDFs (default: True)
    
    Returns:
        str: End User value if found, None otherwise
    """
    try:
        print(f"\n{'='*60}")
        print(f"Reading PDF: {os.path.basename(pdf_path)}")
        print(f"{'='*60}\n")
        
        doc = fitz.open(pdf_path)
        print(f"Total pages in PDF: {len(doc)}")
        
        pages_to_check = min(max_pages, len(doc))
        print(f"Searching first {pages_to_check} page(s)...\n")
        
        for page_num in range(pages_to_check):
            print(f"--- Page {page_num + 1} ---")
            page = doc[page_num]
            
            # METHOD 1: Try direct text extraction first
            text = page.get_text()
            
            if text.strip():
                print("  ℹ️ Text found (searchable PDF)")
                preview = text[:200].replace('\n', ' ')
                print(f"  Preview: {preview}...")
                
                end_user = parse_end_user_from_text(text)
                
                if end_user:
                    print(f"\n✅ FOUND via text extraction on page {page_num + 1}")
                    print(f"{'='*60}")
                    print(f"End User : {end_user}")
                    print(f"{'='*60}\n")
                    doc.close()
                    return end_user
                else:
                    print("  ❌ 'End User' not found in text")
            else:
                print("  ⚠️ No searchable text (likely scanned image)")
            
            # METHOD 2: Try OCR if text extraction failed or found no match
            if use_ocr and (not text.strip() or not end_user):
                print("  🔍 Attempting OCR...")
                
                try:
                    ocr_text = extract_text_via_ocr(page)
                    
                    if ocr_text.strip():
                        print(f"  ✓ OCR extracted {len(ocr_text)} characters")
                        preview = ocr_text[:200].replace('\n', ' ')
                        print(f"  OCR Preview: {preview}...")
                        
                        end_user = parse_end_user_from_text(ocr_text)
                        
                        if end_user:
                            print(f"\n✅ FOUND via OCR on page {page_num + 1}")
                            print(f"{'='*60}")
                            print(f"End User : {end_user}")
                            print(f"{'='*60}\n")
                            doc.close()
                            return end_user
                        else:
                            print("  ❌ 'End User' not found in OCR text")
                    else:
                        print("  ⚠️ OCR extracted no text")
                        
                except Exception as e:
                    print(f"  ❌ OCR failed: {e}")
        
        doc.close()
        print(f"\n⚠️ 'End User' not found in first {pages_to_check} page(s)\n")
        return None
        
    except Exception as e:
        print(f"\n❌ Error reading PDF: {e}\n")
        import traceback
        traceback.print_exc()
        return None


def extract_text_via_ocr(page, zoom=2.0):
    """
    Extract text from PDF page using OCR
    
    Args:
        page: PyMuPDF page object
        zoom: Zoom level for better OCR accuracy (default: 2.0)
    
    Returns:
        str: Extracted text
    """
    # Convert page to high-resolution image
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat)
    
    # Convert to PIL Image
    img_data = pix.tobytes("png")
    img = Image.open(io.BytesIO(img_data))
    
    # Optional: Enhance image for better OCR
    # from PIL import ImageEnhance
    # img = ImageEnhance.Contrast(img).enhance(2.0)
    # img = ImageEnhance.Sharpness(img).enhance(1.5)
    
    # Perform OCR
    ocr_text = pytesseract.image_to_string(img, lang='eng')
    
    return ocr_text


def parse_end_user_from_text(text):
    """
    Parse 'End User' value from text using regex patterns
    
    Args:
        text: Text to search
    
    Returns:
        str: End User value if found, None otherwise
    """
    # Multiple pattern variations
    patterns = [
        r'End\s+User\s*:\s*(.+?)(?:\n|$)',      # End User : <value>
        r'End\s+User\s*:\s*(.+?)(?:\r\n|$)',    # Windows line ending
        r'EndUser\s*:\s*(.+?)(?:\n|$)',         # EndUser: <value>
        r'End\s+User\s*[-–]\s*(.+?)(?:\n|$)',   # End User - <value>
        r'End\s+User\s+:\s+(.+?)(?:\n|$)',      # Extra spaces
        r'End\s*User\s*:\s*(.+?)(?:\n|$)',      # Minimal spaces (OCR sometimes removes spaces)
    ]
    
    for i, pattern in enumerate(patterns, 1):
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if match:
            raw_value = match.group(1).strip()
            cleaned_value = clean_extracted_value(raw_value)
            
            # Validate the extracted value
            if cleaned_value and len(cleaned_value) > 2:
                print(f"  ✓ Matched with pattern #{i}")
                print(f"  Raw value: '{raw_value}'")
                print(f"  Cleaned value: '{cleaned_value}'")
                return cleaned_value
    
    return None


def clean_extracted_value(value):
    """
    Clean extracted value by removing unwanted characters
    
    Args:
        value: Raw extracted value
    
    Returns:
        str: Cleaned value
    """
    # Remove leading/trailing whitespace
    value = value.strip()
    
    # Remove common trailing punctuation
    value = re.sub(r'[,;\.]+$', '', value)
    
    # Remove extra whitespace
    value = re.sub(r'\s+', ' ', value)
    
    # Remove non-printable characters
    value = ''.join(char for char in value if char.isprintable())
    
    # Remove trailing metadata (sometimes PDFs have this)
    value = re.split(r'\s{3,}', value)[0]  # Split on 3+ spaces
    
    # OCR sometimes adds weird characters - clean them
    value = value.replace('|', 'I')  # Common OCR mistake
    value = value.replace('0', 'O') if not any(c.isdigit() for c in value) else value
    
    return value.strip()


def show_all_text_from_page(pdf_path, page_num=0, use_ocr=True):
    """
    Debug function: Show ALL text from a specific page (with OCR option)
    
    Args:
        pdf_path: Path to PDF
        page_num: Page number (0-indexed)
        use_ocr: Use OCR if no text found
    """
    try:
        doc = fitz.open(pdf_path)
        
        if page_num >= len(doc):
            print(f"❌ Page {page_num + 1} doesn't exist. PDF has {len(doc)} pages.")
            doc.close()
            return
        
        page = doc[page_num]
        
        # Try direct text extraction
        text = page.get_text()
        
        print(f"\n{'='*60}")
        print(f"TEXT FROM PAGE {page_num + 1} - Direct Extraction")
        print(f"{'='*60}\n")
        
        if text.strip():
            print(text)
        else:
            print("(No searchable text found)")
        
        # Try OCR
        if use_ocr:
            print(f"\n{'='*60}")
            print(f"TEXT FROM PAGE {page_num + 1} - OCR")
            print(f"{'='*60}\n")
            
            try:
                ocr_text = extract_text_via_ocr(page)
                print(ocr_text if ocr_text.strip() else "(OCR found no text)")
            except Exception as e:
                print(f"❌ OCR failed: {e}")
        
        print(f"\n{'='*60}\n")
        
        doc.close()
        
    except Exception as e:
        print(f"❌ Error: {e}")


def test_ocr_installation():
    """
    Test if Tesseract OCR is properly installed
    """
    print("\n" + "="*60)
    print("TESTING OCR INSTALLATION")
    print("="*60 + "\n")
    
    try:
        version = pytesseract.get_tesseract_version()
        print(f"✅ Tesseract OCR is installed!")
        print(f"   Version: {version}\n")
        return True
    except Exception as e:
        print(f"❌ Tesseract OCR is NOT installed or not found")
        print(f"   Error: {e}\n")
        print("Installation instructions:")
        print("  1. Download from: https://github.com/UB-Mannheim/tesseract/wiki")
        print("  2. Install to default location")
        print(f"  3. Or update TESSERACT_PATH variable in this script\n")
        return False


def main():
    """Main function with GUI file picker"""
    
    print("\n" + "="*60)
    print("PDF END USER READER (WITH OCR SUPPORT)")
    print("="*60)
    print("\nThis tool will:")
    print("  1. Let you select a PDF file")
    print("  2. Try direct text extraction first")
    print("  3. Use OCR for scanned images if needed")
    print("  4. Search for 'End User : <value>' pattern")
    print("  5. Print the extracted value")
    print("\n" + "="*60 + "\n")
    
    # Test OCR installation
    ocr_available = test_ocr_installation()
    
    if not ocr_available:
        proceed = input("OCR not available. Continue anyway? (y/n): ")
        if proceed.lower() != 'y':
            return
    
    # Create hidden tkinter window for file dialog
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    # Ask user to select PDF
    pdf_path = filedialog.askopenfilename(
        title="Select PDF to read End User from",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
    )
    
    if not pdf_path:
        print("❌ No file selected. Exiting.\n")
        return
    
    # Extract End User (with OCR support)
    end_user = extract_end_user_from_pdf(pdf_path, max_pages=5, use_ocr=ocr_available)
    
    # Show result in message box
    if end_user:
        result_msg = f"✅ Successfully extracted:\n\nEnd User : {end_user}"
        messagebox.showinfo("Success", result_msg)
        print(f"✅ Result: {end_user}\n")
    else:
        result_msg = ("⚠️ Could not find 'End User' in PDF\n\n"
                     "The pattern 'End User : <value>' was not found "
                     "in the first 5 pages.\n\n"
                     "Check console output for details.")
        messagebox.showwarning("Not Found", result_msg)
        print("⚠️ End User not found\n")
    
    # Ask if user wants to see full page text for debugging
    show_debug = messagebox.askyesno(
        "Debug",
        "Do you want to see ALL text from page 1?\n"
        "(Shows both direct extraction and OCR)"
    )
    
    if show_debug:
        show_all_text_from_page(pdf_path, page_num=0, use_ocr=ocr_available)


def quick_test_file(pdf_path):
    """
    Quick test function for a specific file
    """
    print("\n" + "="*60)
    print("QUICK TEST MODE")
    print("="*60 + "\n")
    
    if not os.path.exists(pdf_path):
        print(f"❌ File not found: {pdf_path}\n")
        return
    
    end_user = extract_end_user_from_pdf(pdf_path, max_pages=5, use_ocr=True)
    
    if end_user:
        print(f"\n✅ SUCCESS: {end_user}\n")
    else:
        print("\n❌ Not found\n")
        print("Running debug mode...\n")
        show_all_text_from_page(pdf_path, page_num=0, use_ocr=True)


if __name__ == "__main__":
    # Choose mode:
    
    # Mode 1: GUI file picker (DEFAULT)
    main()
    
    # Mode 2: Quick test specific file
    # quick_test_file(r"C:\path\to\your\scanned_pdf.pdf")
    
    # Mode 3: Just test OCR installation
    # test_ocr_installation()
