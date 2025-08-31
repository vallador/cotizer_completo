class AIUManager:
    def __init__(self, database_manager):
        self.database_manager = database_manager

    def get_aiu_values(self):
        """Obtiene los valores de AIU."""
        try:
            cursor = self.database_manager.connection.cursor()
            cursor.execute("SELECT administracion, imprevistos, utilidad, iva_sobre_utilidad FROM aiu_values LIMIT 1")
            row = cursor.fetchone()
            if row:
                return {
                    'administracion': row[0],
                    'imprevistos': row[1],
                    'utilidad': row[2],
                    'iva_sobre_utilidad': row[3]
                }
            return None
        except Exception as e:
            print(f"Error al obtener valores AIU: {str(e)}")
            return None

    def update_aiu_values(self, administracion, imprevistos, utilidad, iva_sobre_utilidad):
        """Actualiza los valores de AIU."""
        try:
            cursor = self.database_manager.connection.cursor()
            cursor.execute("UPDATE aiu_values SET administracion = ?, imprevistos = ?, utilidad = ?, iva_sobre_utilidad = ? WHERE id = 1",
                           (administracion, imprevistos, utilidad, iva_sobre_utilidad))
            self.database_manager.connection.commit()
            return True
        except Exception as e:
            print(f"Error al actualizar valores AIU: {str(e)}")
            return False

    # Agregar este método a tu clase AIUManager

    def set_aiu_values(self, aiu_data):
        """
        Establece los valores AIU desde datos importados

        Args:
            aiu_data (dict): Diccionario con valores AIU
        """
        try:
            if not isinstance(aiu_data, dict):
                print(f"Warning: aiu_data no es un diccionario: {type(aiu_data)}")
                return

            # Actualizar los campos en la interfaz si existen
            if hasattr(self, 'admin_spinbox') and 'administracion' in aiu_data:
                self.admin_spinbox.setValue(float(aiu_data['administracion']))

            if hasattr(self, 'imprev_spinbox') and 'imprevistos' in aiu_data:
                self.imprev_spinbox.setValue(float(aiu_data['imprevistos']))

            if hasattr(self, 'util_spinbox') and 'utilidad' in aiu_data:
                self.util_spinbox.setValue(float(aiu_data['utilidad']))

            if hasattr(self, 'iva_util_spinbox') and 'iva_utilidad' in aiu_data:
                self.iva_util_spinbox.setValue(float(aiu_data['iva_utilidad']))

            print(f"Valores AIU actualizados: {aiu_data}")

        except Exception as e:
            print(f"Error estableciendo valores AIU: {e}")

    def update_from_imported_data(self, aiu_data):
        """
        Método alternativo por si el anterior no funciona
        Actualiza los valores AIU desde datos importados
        """
        try:
            if not aiu_data:
                return

            # Si tienes métodos específicos para actualizar cada campo
            if 'administracion' in aiu_data:
                self.set_administracion(float(aiu_data['administracion']))
            if 'imprevistos' in aiu_data:
                self.set_imprevistos(float(aiu_data['imprevistos']))
            if 'utilidad' in aiu_data:
                self.set_utilidad(float(aiu_data['utilidad']))
            if 'iva_utilidad' in aiu_data:
                self.set_iva_utilidad(float(aiu_data['iva_utilidad']))

        except Exception as e:
            print(f"Error en update_from_imported_data: {e}")

    # Métodos auxiliares si no existen
    def set_administracion(self, value):
        if hasattr(self, 'admin_spinbox'):
            self.admin_spinbox.setValue(value)

    def set_imprevistos(self, value):
        if hasattr(self, 'imprev_spinbox'):
            self.imprev_spinbox.setValue(value)

    def set_utilidad(self, value):
        if hasattr(self, 'util_spinbox'):
            self.util_spinbox.setValue(value)

    def set_iva_utilidad(self, value):
        if hasattr(self, 'iva_util_spinbox'):
            self.iva_util_spinbox.setValue(value)