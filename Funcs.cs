using SpreadsheetLight;
using CrearAlumno;
using CrearEscuela;
using Spectre.Console;

namespace Funcs
{
    public class Functions
    {
        SLDocument incripciones = new SLDocument(@"DATA\Inscripciones.xlsx");
        List<Alumno> Alumnos = new List<Alumno>();
        Escuela Escuela_SRosa = Escuela.RecopilarDatos("SantaRosa");
        Escuela Escuela_Anguil = Escuela.RecopilarDatos("Anguil");
        Escuela Escuela_Toay = Escuela.RecopilarDatos("Toay");
        SLDocument VacantesToay = Escuela.CrearVacantes();
        SLDocument VacantesSRosa = Escuela.CrearVacantes();
        SLDocument VacantesAnguil = Escuela.CrearVacantes();
        SLDocument VacantesListaDeEspera = Escuela.CrearVacantes();
        List<Alumno> AlumnosEnListaDeEspera = new List<Alumno>();
        public void GestionarInscripciones()
        {
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
        }
        public void VerVacantes()
        {
            TablaVacante(Escuela_SRosa);
            TablaVacante(Escuela_Anguil);
            TablaVacante(Escuela_Toay);
            Mensaje("Exito");
        }
        public void Mensaje(string opcion)
        {
            switch (opcion)
            {
                case "Exito":
                    AnsiConsole.WriteLine("Operacion exitosa!");
                    Mensaje("Continuar");
                    break;
                case "Continuar":
                    AnsiConsole.WriteLine("Para continuar oprima un tecla");
                    Console.ReadKey();
                    AnsiConsole.Clear();
                    AnsiConsole.ResetColors();
                    break;
                case "Error":
                    AnsiConsole.WriteLine("Ocurrio un error!");
                    Mensaje("Continuar");
                    break;
            }
        }
        public void TablaVacante(Escuela escuela)
        {
            var table = new Table();
            table.Title($"Escuela {escuela.nombre}").Centered();
            table.AddColumn("Grado 1").Centered();
            table.AddColumn("Grado 2").Centered();
            table.AddColumn("Grado 3").Centered();
            table.AddColumn("Grado 4").Centered();
            table.AddColumn("Grado 5").Centered();
            table.AddColumn("Grado 6").Centered();
            table.AddColumn("Grado 7").Centered();
            table.AddRow($"{escuela.vacantes[1]}", $"{escuela.vacantes[2]}", $"{escuela.vacantes[3]}", $"{escuela.vacantes[4]}", $"{escuela.vacantes[5]}", $"{escuela.vacantes[6]}", $"{escuela.vacantes[7]}");
            AnsiConsole.Write(table);
            AnsiConsole.Foreground = ConsoleColor.White;
        }
        public void GuardarDocumentos()
        {
            try{
                VacantesToay.SaveAs(@"Vacantes\Vacantes_ColegioToay.xlsx");
                VacantesSRosa.SaveAs(@"Vacantes\Vacantes_ColegioSRosa.xlsx");
                VacantesAnguil.SaveAs(@"Vacantes\Vacantes_ColegioAnguil.xlsx");
                VacantesListaDeEspera.SaveAs(@"Vacantes\Error\VacantesListaDeEspera.xlsx");                
                Mensaje("Exito");
            }catch{
                Mensaje("Error");
            }
        }
    }
}