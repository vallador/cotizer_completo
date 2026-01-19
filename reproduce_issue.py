import os
import sys
from reportlab.pdfgen import canvas

# Add project root to path
sys.path.append(r"c:\Users\DELFOS\PycharmProjects\cotizer_completo")

from utils.pdf_merger import PDFMerger

def create_dummy_pdf(filename, text="Dummy PDF"):
    c = canvas.Canvas(filename)
    c.drawString(100, 750, text)
    c.save()
    return os.path.abspath(filename)

def test_merger():
    print("Iniciando prueba de PDFMerger con Propuesta Técnica...")
    
    # 1. Setup paths
    templates_dir = r"c:\Users\DELFOS\PycharmProjects\cotizer_completo\data\templates"
    output_pdf = "test_merged_output_v2.pdf"
    
    # Create dummy PDFs
    dummy_budget = create_dummy_pdf("dummy_budget.pdf", "PRESUPUESTO GENERADO")
    dummy_proposal = create_dummy_pdf("dummy_propuesta.pdf", "CONTENIDO PROPUESTA TÉCNICA (Word)")

    # 2. Check templates existence
    if not os.path.exists(templates_dir):
        print("ERROR: Templates dir does not exist!")
        return

    # 3. Instantiate Merger
    merger = PDFMerger(templates_dir)
    
    # 4. Define items to merge (simulate what MainWindow does now)
    # Including 'propuesta_tecnica' manual insertion logic
    ordered_items = [
        "portadas", 
        "contenido_separadores", 
        "propuesta_tecnica", # Added manually
        "paginas_estandar", 
        "presupuesto_programacion",
        "certificados_trabajos"
    ]
    
    external_map = {
        "propuesta_tecnica": dummy_proposal
    }
    
    print(f"Items to merge: {ordered_items}")
    
    # 5. Run Merge
    success = merger.merge_pdfs(
        output_path=output_pdf,
        ordered_items=ordered_items,
        generated_quotation_pdf=dummy_budget,
        external_files_map=external_map
    )
    
    if success:
        print(f"SUCCESS: PDF merged at {os.path.abspath(output_pdf)}")
        # Check size to ensure it's not just the dummy
        size = os.path.getsize(output_pdf)
        print(f"Output size: {size} bytes")
    else:
        print("FAILURE: merge_pdfs returned False")

    # Cleanup
    if os.path.exists("dummy_budget.pdf"): os.remove("dummy_budget.pdf")
    if os.path.exists("dummy_propuesta.pdf"): os.remove("dummy_propuesta.pdf")

    # Cleanup
    if os.path.exists("dummy_budget.pdf"): os.remove("dummy_budget.pdf")
    # if os.path.exists(output_pdf): os.remove(output_pdf) # Keep it for inspection

if __name__ == "__main__":
    test_merger()
