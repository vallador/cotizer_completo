import win32com.client as win32
import os
import pythoncom
import time

class ExcelToWordAutomation:
    def __init__(self):
        self.excel = None
        self.word = None

    def ejecutar_flujo_completo(self, excel_path, word_template_path, pdf_output_path):
        """
        Coordina la apertura de Excel, copia de tabla, pegado en Word y exportación a PDF.
        """
        try:
            # Inicializar aplicaciones
            # win32.gencache.EnsureDispatch es más lento pero genera los métodos estáticos de Office
            self.excel = win32.Dispatch("Excel.Application")
            self.word = win32.Dispatch("Word.Application")

            self.excel.Visible = False
            self.word.Visible = False

            # --- PARTE 1: EXCEL ---
            wb = self.excel.Workbooks.Open(os.path.abspath(excel_path))
            ws = wb.ActiveSheet

            # Seleccionar y copiar el rango con datos
            ws.UsedRange.Copy()

            # --- PARTE 2: WORD ---
            doc = self.word.Documents.Open(os.path.abspath(word_template_path))

            # Moverse al final del documento para pegar
            rango = doc.Content
            rango.Collapse(0)  # 0 = wdCollapseEnd

            # Pegar manteniendo el formato original de Excel
            # Parámetros: LinkedToExcel, WordFormatting, RTF
            rango.PasteExcelTable(False, False, False)

            # --- PARTE 3: AUTOAJUSTE ---
            if doc.Tables.Count > 0:
                ultima_tabla = doc.Tables(doc.Tables.Count)
                # 1 = wdAutoFitWindow
                ultima_tabla.AutoFitBehavior(1)

            # --- PARTE 4: GENERAR PDF ---
            # 17 = wdFormatPDF
            doc.ExportAsFixedFormat(os.path.abspath(pdf_output_path), 17)

            self.excel.CutCopyMode = False

            # Ahora puedes cerrar sin que salte el aviso
            doc.Close(False)
            wb.Close(False)

            return True, "Proceso completado correctamente."

        except Exception as e:
            return False, str(e)

        finally:
            # Asegurar que los procesos de Office se cierren incluso si hay error
            if self.excel: self.excel.Quit()
            if self.word: self.word.Quit()

    def insertar_tabla_y_convertir_pdf(self, excel_path, word_path, pdf_output_path):
        try:
            import pythoncom
            pythoncom.CoInitialize()

            self.excel = win32.Dispatch("Excel.Application")
            self.word = win32.Dispatch("Word.Application")

            # --- CLAVE PARA OMITIR AVISOS ---
            self.excel.DisplayAlerts = False
            self.word.DisplayAlerts = False
            # -------------------------------

            self.excel.Visible = False
            self.word.Visible = False

            # 1. Copiar de Excel
            wb = self.excel.Workbooks.Open(os.path.abspath(excel_path))
            ws = wb.ActiveSheet
            ws.UsedRange.Copy()

            # 2. Abrir Word
            doc = self.word.Documents.Open(os.path.abspath(word_path))

            # 3. Buscar título y pegar
            rango = doc.Content
            if rango.Find.Execute("ALCANCE DE LOS TRABAJOS"):
                rango.Select()
                self.word.Selection.MoveDown(Unit=5, Count=1)
                self.word.Selection.TypeParagraph()
                self.word.Selection.PasteExcelTable(False, False, False)

                if doc.Tables.Count > 0:
                    doc.Tables(doc.Tables.Count).AutoFitBehavior(1)

            # --- SEGUNDA CLAVE: LIMPIAR PORTAPAPELES ---
            # Esto le dice a Excel que ya no necesitamos lo que se copió
            self.excel.CutCopyMode = False
            # -------------------------------------------

            # 4. Exportar a PDF
            doc.ExportAsFixedFormat(os.path.abspath(pdf_output_path), 17)

            doc.Save()
            doc.Close(False)
            wb.Close(False)
            return True, "OK"

        except Exception as e:
            return False, str(e)
        finally:
            if self.excel:
                self.excel.Quit()
            if self.word:
                self.word.Quit()
            pythoncom.CoUninitialize()