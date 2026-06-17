import fitz  # PyMuPDF

def decrypt_pdf_keep_structure(input_path, output_path, password):
    try:
        # Open the document
        doc = fitz.open(input_path)
        
        # Authenticate if it is password protected
        if doc.is_encrypted:
            is_authenticated = doc.authenticate(password)
            if not is_authenticated:
                raise ValueError("Incorrect password.")
        
        # Save a clean, unencrypted copy while preserving layout
        doc.save(
            output_path, 
            deflate=True,          # Compresses streams losslessly
            garbage=3,             # Removes unused and duplicate objects
            clean=True             # Cleans up content streams
        )
        return True
    except Exception as e:
        raise e

if __name__ == "__main__":
    decrypt_pdf_keep_structure(r"sbi.pdf",
     r"sbi_unlocked.pdf", "56914211198")