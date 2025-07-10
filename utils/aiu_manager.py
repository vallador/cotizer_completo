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

