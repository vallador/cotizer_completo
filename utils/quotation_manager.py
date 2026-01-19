# utils/quotation_manager.py
"""
Extension methods for DatabaseManager to handle quotations, history, and snapshots.
This module contains all the methods for the dashboard and quotation management system.
"""
import json
from datetime import datetime, timedelta
import sqlite3


class QuotationManager:
    """Mixin class for quotation management - extends DatabaseManager"""
    
    # ===== COTIZACIONES GENERADAS =====
    
    def save_quotation(self, **kwargs):
        """
        Saves a generated quotation to the database.
        
        Args:
            cliente_id (int): Client ID
            nombre_proyecto (str): Project name
            monto_total (float): Total amount
            estado (str): State - 'pendiente', 'ganada', 'perdida', 'cancelada'
            es_prueba (bool): Whether this is a test quotation
            ruta_pdf (str): Path to generated PDF
            ruta_excel (str): Path to generated Excel
            ruta_word (str): Path to generated Word
            notas (str): Optional notes
            validez_dias (int): Validity in days (default 30)
            tipo_cliente (str): Client type - 'natural' or 'juridica'
        
        Returns:
            int: The ID of the saved quotation, or None if error
        """
        try:
            cursor = self.connection.cursor()
            
            # Calculate expiration date
            fecha_vencimiento = (datetime.now() + timedelta(days=kwargs.get('validez_dias', 30))).date()
            
            cursor.execute("""
                INSERT INTO cotizaciones_generadas (
                    cliente_id, nombre_proyecto, monto_total, estado, es_prueba,
                    ruta_pdf, ruta_excel, ruta_word, notas, validez_dias, 
                    fecha_vencimiento, tipo_cliente
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                kwargs.get('cliente_id'),
                kwargs.get('nombre_proyecto'),
                kwargs.get('monto_total'),
                kwargs.get('estado', 'pendiente'),
                kwargs.get('es_prueba', False),
                kwargs.get('ruta_pdf'),
                kwargs.get('ruta_excel'),
                kwargs.get('ruta_word'),
                kwargs.get('notas', ''),
                kwargs.get('validez_dias', 30),
                fecha_vencimiento,
                kwargs.get('tipo_cliente', 'natural')
            ))
            
            self.connection.commit()
            quotation_id = cursor.lastrowid
            print(f"Cotización guardada con ID: {quotation_id}")
            return quotation_id
            
        except sqlite3.Error as e:
            print(f"Error al guardar cotización: {e}")
            return None
    
    def get_all_quotations(self, include_test=True, filters=None):
        """
        Gets all quotations with optional filters.
        
        Args:
            include_test (bool): Whether to include test quotations
            filters (dict): Optional filters:
                - estado: Filter by state
                - cliente_id: Filter by client
                - fecha_inicio: Start date
                - fecha_fin: End date
                - monto_min: Minimum amount
                - monto_max: Maximum amount
        
        Returns:
            list: List of quotation dictionaries
        """
        try:
            cursor = self.connection.cursor()
            
            query = """
                SELECT 
                    cg.id, cg.cliente_id, c.nombre as cliente_nombre,
                    cg.fecha_creacion, cg.fecha_modificacion, cg.nombre_proyecto,
                    cg.monto_total, cg.estado, cg.es_prueba, cg.ruta_pdf,
                    cg.ruta_excel, cg.ruta_word, cg.notas, cg.validez_dias,
                    cg.fecha_vencimiento, cg.tipo_cliente
                FROM cotizaciones_generadas cg
                LEFT JOIN clientes c ON cg.cliente_id = c.id
                WHERE 1=1
            """
            
            params = []
            
            if not include_test:
                query += " AND cg.es_prueba = 0"
            
            if filters:
                if 'estado' in filters and filters['estado']:
                    query += " AND cg.estado = ?"
                    params.append(filters['estado'])
                
                if 'cliente_id' in filters and filters['cliente_id']:
                    query += " AND cg.cliente_id = ?"
                    params.append(filters['cliente_id'])
                
                if 'fecha_inicio' in filters and filters['fecha_inicio']:
                    query += " AND date(cg.fecha_creacion) >= ?"
                    params.append(filters['fecha_inicio'])
                
                if 'fecha_fin' in filters and filters['fecha_fin']:
                    query += " AND date(cg.fecha_creacion) <= ?"
                    params.append(filters['fecha_fin'])
                
                if 'monto_min' in filters and filters['monto_min']:
                    query += " AND cg.monto_total >= ?"
                    params.append(filters['monto_min'])
                
                if 'monto_max' in filters and filters['monto_max']:
                    query += " AND cg.monto_total <= ?"
                    params.append(filters['monto_max'])
            
            query += " ORDER BY cg.fecha_creacion DESC"
            
            cursor.execute(query, params)
            
            quotations = []
            for row in cursor.fetchall():
                quotations.append({
                    'id': row[0],
                    'cliente_id': row[1],
                    'cliente_nombre': row[2],
                    'fecha_creacion': row[3],
                    'fecha_modificacion': row[4],
                    'nombre_proyecto': row[5],
                    'monto_total': row[6],
                    'estado': row[7],
                    'es_prueba': bool(row[8]),
                    'ruta_pdf': row[9],
                    'ruta_excel': row[10],
                    'ruta_word': row[11],
                    'notas': row[12],
                    'validez_dias': row[13],
                    'fecha_vencimiento': row[14],
                    'tipo_cliente': row[15]
                })
            
            return quotations
            
        except sqlite3.Error as e:
            print(f"Error al obtener cotizaciones: {e}")
            return []
    
    def get_quotation_by_id(self, quotation_id):
        """Gets a single quotation by ID"""
        try:
            cursor = self.connection.cursor()
            cursor.execute("""
                SELECT 
                    cg.id, cg.cliente_id, c.nombre as cliente_nombre,
                    cg.fecha_creacion, cg.fecha_modificacion, cg.nombre_proyecto,
                    cg.monto_total, cg.estado, cg.es_prueba, cg.ruta_pdf,
                    cg.ruta_excel, cg.ruta_word, cg.notas, cg.validez_dias,
                    cg.fecha_vencimiento, cg.tipo_cliente
                FROM cotizaciones_generadas cg
                LEFT JOIN clientes c ON cg.cliente_id = c.id
                WHERE cg.id = ?
            """, (quotation_id,))
            
            row = cursor.fetchone()
            if row:
                return {
                    'id': row[0],
                    'cliente_id': row[1],
                    'cliente_nombre': row[2],
                    'fecha_creacion': row[3],
                    'fecha_modificacion': row[4],
                    'nombre_proyecto': row[5],
                    'monto_total': row[6],
                    'estado': row[7],
                    'es_prueba': bool(row[8]),
                    'ruta_pdf': row[9],
                    'ruta_excel': row[10],
                    'ruta_word': row[11],
                    'notas': row[12],
                    'validez_dias': row[13],
                    'fecha_vencimiento': row[14],
                    'tipo_cliente': row[15]
                }
            return None
            
        except sqlite3.Error as e:
            print(f"Error al obtener cotización: {e}")
            return None
    
    def update_quotation_state(self, quotation_id, new_state, notas=None):
        """Updates the state of a quotation and logs it"""
        try:
            cursor = self.connection.cursor()
            
            # Update state
            cursor.execute("""
                UPDATE cotizaciones_generadas 
                SET estado = ?, fecha_modificacion = CURRENT_TIMESTAMP
                WHERE id = ?
            """, (new_state, quotation_id))
            
            # Update notes if provided
            if notas:
                cursor.execute("""
                    UPDATE cotizaciones_generadas 
                    SET notas = ?
                    WHERE id = ?
                """, (notas, quotation_id))
            
            self.connection.commit()
            
            # Log the change in history
            self.add_quotation_history(
                quotation_id,
                accion=f"estado_cambiado_a_{new_state}",
                notas=notas
            )
            
            return True
            
        except sqlite3.Error as e:
            print(f"Error al actualizar estado: {e}")
            return False
    
    def delete_quotation(self, quotation_id):
        """Deletes a quotation (cascades to history and snapshots)"""
        try:
            cursor = self.connection.cursor()
            cursor.execute("DELETE FROM cotizaciones_generadas WHERE id = ?", (quotation_id,))
            self.connection.commit()
            return  True
        except sqlite3.Error as e:
            print(f"Error al eliminar cotización: {e}")
            return False
    
    # ===== HISTORIAL =====
    
    def add_quotation_history(self, quotation_id, accion, usuario=None, notas=None):
        """Adds an entry to the quotation history"""
        try:
            cursor = self.connection.cursor()
            cursor.execute("""
                INSERT INTO historial_cotizacion (cotizacion_id, accion, usuario, notas)
                VALUES (?, ?, ?, ?)
            """, (quotation_id, accion, usuario, notas))
            self.connection.commit()
            return cursor.lastrowid
        except sqlite3.Error as e:
            print(f"Error al agregar historial: {e}")
            return None
    
    def get_quotation_history(self, quotation_id):
        """Gets the complete history of a quotation"""
        try:
            cursor = self.connection.cursor()
            cursor.execute("""
                SELECT id, fecha, accion, usuario, notas
                FROM historial_cotizacion
                WHERE cotizacion_id = ?
                ORDER BY fecha DESC
            """, (quotation_id,))
            
            history = []
            for row in cursor.fetchall():
                history.append({
                    'id': row[0],
                    'fecha': row[1],
                    'accion': row[2],
                    'usuario': row[3],
                    'notas': row[4]
                })
            return history
            
        except sqlite3.Error as e:
            print(f"Error al obtener historial: {e}")
            return []
    
    # ===== SNAPSHOTS =====
    
    def save_snapshot(self, quotation_id, datos_dict, table_rows_list, config_dict=None):
        """
        Saves a complete snapshot of a quotation for recovery.
        
        Args:
            quotation_id (int): Quotation ID
            datos_dict (dict): Client data, AIU values, etc.
            table_rows_list (list): List of table rows
            config_dict (dict): Optional Word/PDF configuration
        
        Returns:
            int: Snapshot ID or None
        """
        try:
            cursor = self.connection.cursor()
            
            # Serialize to JSON
            datos_json = json.dumps(datos_dict, ensure_ascii=False)
            table_rows_json = json.dumps(table_rows_list, ensure_ascii=False)
            config_json = json.dumps(config_dict, ensure_ascii=False) if config_dict else None
            
            cursor.execute("""
                INSERT INTO cotizaciones_snapshot (cotizacion_id, datos_json, table_rows_json, config_json)
                VALUES (?, ?, ?, ?)
            """, (quotation_id, datos_json, table_rows_json, config_json))
            
            self.connection.commit()
            return cursor.lastrowid
            
        except (sqlite3.Error, json.JSONDecodeError) as e:
            print(f"Error al guardar snapshot: {e}")
            return None
    
    def get_latest_snapshot(self, quotation_id):
        """Gets the most recent snapshot of a quotation"""
        try:
            cursor = self.connection.cursor()
            cursor.execute("""
                SELECT id, fecha_snapshot, datos_json, table_rows_json, config_json
                FROM cotizaciones_snapshot
                WHERE cotizacion_id = ?
                ORDER BY fecha_snapshot DESC
                LIMIT 1
            """, (quotation_id,))
            
            row = cursor.fetchone()
            if row:
                return {
                    'id': row[0],
                    'fecha_snapshot': row[1],
                    'datos': json.loads(row[2]),
                    'table_rows': json.loads(row[3]),
                    'config': json.loads(row[4]) if row[4] else None
                }
            return None
            
        except (sqlite3.Error, json.JSONDecodeError) as e:
            print(f"Error al obtener snapshot: {e}")
            return None
    
    # ===== ESTADÍSTICAS =====
    
    def get_quotation_stats(self, mes=None, anio=None, include_test=False):
        """
        Gets statistics for quotations.
        
        Args:
            mes (int): Month (1-12), None for all
            anio (int): Year, None for current year
            include_test (bool): Include test quotations
        
        Returns:
            dict: Statistics including counts and totals by state
        """
        try:
            cursor = self.connection.cursor()
            
            # Base query
            where_clauses = []
            params = []
            
            if not include_test:
                where_clauses.append("es_prueba = 0")
            
            if mes and anio:
                where_clauses.append("strftime('%m', fecha_creacion) = ?")
                where_clauses.append("strftime('%Y', fecha_creacion) = ?")
                params.extend([f"{mes:02d}", str(anio)])
            elif anio:
                where_clauses.append("strftime('%Y', fecha_creacion) = ?")
                params.append(str(anio))
            
            where_sql = " AND ".join(where_clauses) if where_clauses else "1=1"
            
            # Get counts and totals by state
            cursor.execute(f"""
                SELECT 
                    estado,
                    COUNT(*) as count,
                    SUM(monto_total) as total
                FROM cotizaciones_generadas
                WHERE {where_sql}
                GROUP BY estado
            """, params)
            
            stats = {
                'pendiente': {'count': 0, 'total': 0},
                'ganada': {'count': 0, 'total': 0},
                'perdida': {'count': 0, 'total': 0},
                'cancelada': {'count': 0, 'total': 0}
            }
            
            for row in cursor.fetchall():
                estado = row[0]
                if estado in stats:
                    stats[estado] = {
                        'count': row[1],
                        'total': row[2] if row[2] else 0
                    }
            
            # Calculate totals
            stats['total_cotizaciones'] = sum(s['count'] for s in stats.values() if isinstance(s, dict))
            stats['total_monto'] = sum(s['total'] for s in stats.values() if isinstance(s, dict))
            
            # Conversion rate
            if stats['total_cotizaciones'] > 0:
                stats['tasa_conversion'] = (stats['ganada']['count'] / stats['total_cotizaciones']) * 100
            else:
                stats['tasa_conversion'] = 0
            
            return stats
            
        except sqlite3.Error as e:
            print(f"Error al obtener estadísticas: {e}")
            return None
