/**
 * Repositorio para la ficha clínica (Bloque 19).
 */
class ClinicalDataRepository extends BaseRepository {
  constructor() {
    // Nota: Esta hoja se crea dinámicamente en el bloque 19, 
    // pero definimos el repo para estandarizar el acceso.
    super('DATOS_CLINICOS_PACIENTES', []); 
  }
}