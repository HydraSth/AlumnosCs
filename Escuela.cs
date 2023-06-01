using SpreadsheetLight;
using CrearAlumno;

namespace CrearEscuela
{
    public class Escuela
    {
        public List<int> grado = new List<int> { 0 };
        public List<int> vacantes = new List<int> { 0 };
        public int indice = 2;
        public int indiceError = 2;
        public static SLDocument CrearVacantes(){
            SLDocument XLSX= new SLDocument();
            XLSX.SetCellValue(1, 1, "nro_inscripcion");
            XLSX.SetCellValue(1, 2, "nombre");
            XLSX.SetCellValue(1, 3, "edad");
            XLSX.SetCellValue(1, 4, "grado");
            return XLSX;
        }
        public static Escuela RecopilarDatos(string Nombre)
        {
            Escuela Escuela_X = new Escuela();
            string doc_toay = @$"C:\Users\Usuario\Desktop\Evaluativo_Practica\DATA\Colegio_{Nombre}.xlsx";
            SLDocument xls_toay = new SLDocument(doc_toay);
            int fila = 2;
            while (!string.IsNullOrEmpty(xls_toay.GetCellValueAsString(fila, 1)))
            {
                int grado = xls_toay.GetCellValueAsInt32(fila, 1);
                int vacante = xls_toay.GetCellValueAsInt32(fila, 2);
                Escuela_X.agregarAlumnos(grado, vacante);
                fila++;
            }
            return Escuela_X;
        }
        public void agregarAlumnos(int GRADO, int VACANTE)
        {
            grado.Add(GRADO);
            vacantes.Add(VACANTE);
        }
        public bool actualizarAlumnos(int GRADO)
        {
            if (vacantes != null && vacantes[GRADO] > 0)
            {
                vacantes[GRADO]--;
                return true;
            }
            return false;
        }
        public Tuple<SLDocument,bool> datosActualizado(Alumno obj_alumno, SLDocument archivo, Escuela escuela , SLDocument error)
        {
            if(archivo?.GetCellValueAsInt32(indice, 1) != null)
            {
                if(escuela.vacantes[obj_alumno.grado] > 0){
                    archivo.SetCellValue(indice, 1, obj_alumno.id);
                    archivo.SetCellValue(indice, 2, obj_alumno.nombre);
                    archivo.SetCellValue(indice, 3, obj_alumno.edad);
                    archivo.SetCellValue(indice, 4, obj_alumno.grado);
                    escuela.actualizarAlumnos(obj_alumno.grado);
                    this.indice++;
                    return Tuple.Create(archivo,true);
                }else{
                    error.SetCellValue(indiceError, 1, obj_alumno.id);
                    error.SetCellValue(indiceError, 2, obj_alumno.nombre);
                    error.SetCellValue(indiceError, 3, obj_alumno.edad);
                    error.SetCellValue(indiceError, 4, obj_alumno.grado);
                    this.indiceError++;
                }
            }
            return Tuple.Create(error,false);
        }

    }
}