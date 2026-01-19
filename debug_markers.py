import sys
import os

# Add the project root to sys.path so we can import from controllers
sys.path.append(os.getcwd())

from controllers.word_controller import WordController

class MockCotizacionController:
    pass

def main():
    try:
        controller = WordController(MockCotizacionController())
        print("Checking Juridica Larga template markers...")
        
        template_path = controller.template_juridica_larga
        if not os.path.exists(template_path):
            print(f"Error: Template not found at {template_path}")
            return

        markers = controller._extract_markers_from_template(template_path)
        print(f"Markers found in {os.path.basename(template_path)}:")
        for marker in markers:
            print(f"  - {marker}")
            
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
