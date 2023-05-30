namespace CrearAlumno
{
    public class Alumno{
        public int ?id  { get; set; }
        public string ?nombre  { get; set; }
        public int ?edad  { get; set; }
        public int ?grado  { get; set; }
        public string ?escuela  { get; set; }
        public Alumno(int id, string nombre, int edad, int grado, string escuela){
            this.id = id;
            this.nombre = nombre;
            this.edad = edad;
            this.grado = grado;
            this.escuela = escuela;
        }
    }
}