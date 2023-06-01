namespace CrearAlumno
{
    public class Alumno{
        public int id  { get; set; }
        public string nombre  { get; set; }
        public int edad  { get; set; }
        public int grado  { get; set; }
        public string preferencia  { get; set; }
        public Alumno(int id, string nombre, int edad, int grado, string preferencia){
            this.id = id;
            this.nombre = nombre;
            this.edad = edad;
            this.grado = grado;
            this.preferencia = preferencia;
        }
    }
}