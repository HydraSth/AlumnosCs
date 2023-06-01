using SpreadsheetLight;
using CrearAlumno;
using CrearEscuela;

string doc_incripcione = @"C:\Users\Usuario\Desktop\Evaluativo_Practica\DATA\Inscripciones.xlsx";
SLDocument incripciones = new SLDocument(doc_incripcione);
List<Alumno> Alumnos = new List<Alumno>();

int fila = 2;
while (!string.IsNullOrEmpty(incripciones.GetCellValueAsString(fila, 1)))
{
    int id = incripciones.GetCellValueAsInt32(fila, 1);
    string nombre = incripciones.GetCellValueAsString(fila, 2);
    int edad = incripciones.GetCellValueAsInt32(fila, 3);
    int grado = incripciones.GetCellValueAsInt32(fila, 4);
    string preferencia = incripciones.GetCellValueAsString(fila, 5);
    Alumnos.Add(new Alumno(id, nombre, edad, grado, preferencia));
    fila++;
}

Escuela Escuela_SRosa = Escuela.RecopilarDatos("SantaRosa");
Escuela Escuela_Anguil = Escuela.RecopilarDatos("Anguil");
Escuela Escuela_Toay = Escuela.RecopilarDatos("Toay");

SLDocument VacantesToay = Escuela.CrearVacantes();
SLDocument VacantesSRosa = Escuela.CrearVacantes();
SLDocument VacantesAnguil = Escuela.CrearVacantes();

SLDocument VacantesListaDeEspera = Escuela.CrearVacantes();
int indiceError = 2;
foreach (var alumno in Alumnos)
{
    switch (alumno.preferencia)
    {
        case "Nro. 1 - Toay":
            var respuesta_vacante = Escuela_Toay.datosActualizado(alumno, VacantesToay, Escuela_Toay, VacantesListaDeEspera, indiceError);
            if (respuesta_vacante.Item2)
            {
                VacantesToay = respuesta_vacante.Item1;
            }
            else
            {
                VacantesListaDeEspera = respuesta_vacante.Item1;
                indiceError++;
            }
            break;
        case "Nro. 2 - Santa Rosa":
            respuesta_vacante = Escuela_SRosa.datosActualizado(alumno, VacantesSRosa, Escuela_SRosa, VacantesListaDeEspera, indiceError);
            if (respuesta_vacante.Item2)
            {
                VacantesSRosa = respuesta_vacante.Item1;
            }
            else
            {
                VacantesListaDeEspera = respuesta_vacante.Item1;
                indiceError++;
            }
            break;
        case "Nro. 3 - Anguil":
            respuesta_vacante = Escuela_Anguil.datosActualizado(alumno, VacantesAnguil, Escuela_Anguil, VacantesListaDeEspera, indiceError);
            if (respuesta_vacante.Item2)
            {
                VacantesAnguil = respuesta_vacante.Item1;
            }
            else
            {
                VacantesListaDeEspera = respuesta_vacante.Item1;
                indiceError++;
            }
            break;
    }
}

List<Alumno> AlumnosEnListaDeEspera = new List<Alumno>();
fila = 2;
while (!string.IsNullOrEmpty(VacantesListaDeEspera.GetCellValueAsString(fila, 1)))
{
    int id = VacantesListaDeEspera.GetCellValueAsInt32(fila, 1);
    string nombre = VacantesListaDeEspera.GetCellValueAsString(fila, 2);
    int edad = VacantesListaDeEspera.GetCellValueAsInt32(fila, 3);
    int grado = VacantesListaDeEspera.GetCellValueAsInt32(fila, 4);
    string preferencia = "x";
    AlumnosEnListaDeEspera.Add(new Alumno(id, nombre, edad, grado, preferencia));
    fila++;
}

VacantesListaDeEspera = Escuela.CrearVacantes();
indiceError = 2;
foreach (var alumno in AlumnosEnListaDeEspera)
{
    //Prueba con escuela Toay
    var respuesta_vacante = Escuela_Toay.datosActualizado(alumno, VacantesToay, Escuela_Toay, VacantesListaDeEspera, indiceError);
    if (respuesta_vacante.Item2)
    {
        VacantesToay = respuesta_vacante.Item1;
    }
    else
    {
        //Prueba con escuela Anguil
        respuesta_vacante = Escuela_Anguil.datosActualizado(alumno, VacantesAnguil, Escuela_Anguil, VacantesListaDeEspera, indiceError);
        if (respuesta_vacante.Item2)
        {
            VacantesAnguil = respuesta_vacante.Item1;
        }
        else
        {
            //Prueba con escuela Santa Rosa
            respuesta_vacante = Escuela_Toay.datosActualizado(alumno, VacantesToay, Escuela_Toay, VacantesListaDeEspera, indiceError);
            if (respuesta_vacante.Item2)
            {
                VacantesToay = respuesta_vacante.Item1;
            }
            else
            {
                //No pudo ser ingresado en ningun colegio queda en lista de espera
                VacantesListaDeEspera = respuesta_vacante.Item1;
                indiceError++;
            }
        }
    }
}

//Chequeo para mostrar las vacantes
// for (int i = 0; i < 8; i++){
//     System.Console.WriteLine($"A la escuela santa rosa grado {Escuela_SRosa.grado[i]} le quedan {Escuela_SRosa.vacantes[i]}");
//     System.Console.WriteLine($"A la escuela toay grado {Escuela_Toay.grado[i]} le quedan {Escuela_Toay.vacantes[i]}");
//     System.Console.WriteLine($"A la escuela anguil grado {Escuela_Anguil.grado[i]} le quedan {Escuela_Anguil.vacantes[i]}");
// }

VacantesToay.SaveAs(@"C:\Users\Usuario\Desktop\Evaluativo_Practica\Vacantes\Vacantes_ColegioToay.xlsx");
VacantesSRosa.SaveAs(@"C:\Users\Usuario\Desktop\Evaluativo_Practica\Vacantes\Vacantes_ColegioSRosa.xlsx");
VacantesAnguil.SaveAs(@"C:\Users\Usuario\Desktop\Evaluativo_Practica\Vacantes\Vacantes_ColegioAnguil.xlsx");

VacantesListaDeEspera.SaveAs(@"C:\Users\Usuario\Desktop\Evaluativo_Practica\Vacantes\Error\VacantesListaDeEspera.xlsx");
