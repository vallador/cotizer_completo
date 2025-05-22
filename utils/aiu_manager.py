class AIUManager:
    def __init__(self, administracion=7, imprevistos=3, utilidad=5, iva_sobre_utilidad=19):
        # Inicializar valores con los predeterminados o los valores pasados como parámetros
        self.administracion = administracion
        self.imprevistos = imprevistos
        self.utilidad = utilidad
        self.iva_sobre_utilidad = iva_sobre_utilidad

    def calcular_aiu(self, subtotal):
        """Aplica los porcentajes de AIU al subtotal."""
        admin_value = subtotal * (self.administracion / 100)
        imprevistos_value = subtotal * (self.imprevistos / 100)
        utilidad_value = subtotal * (self.utilidad / 100)
        iva_value = utilidad_value * (self.iva_sobre_utilidad / 100)

        return admin_value, imprevistos_value, utilidad_value, iva_value

    def modificar_aiu(self, administracion, imprevistos, utilidad, iva_sobre_utilidad):
        """Permite modificar los valores de AIU"""
        self.administracion = administracion
        self.imprevistos = imprevistos
        self.utilidad = utilidad
        self.iva_sobre_utilidad = iva_sobre_utilidad

    def obtener_aiu(self):
        """Retorna los valores actuales de AIU."""
        return {
            'administracion': self.administracion,
            'imprevistos': self.imprevistos,
            'utilidad': self.utilidad,
            'iva_sobre_utilidad': self.iva_sobre_utilidad
        }
