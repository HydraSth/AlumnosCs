using SpreadsheetLight;
using CrearAlumno;
using CrearEscuela;

//Leer incripciones
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

SLDocument VacantesError_Toay = Escuela.CrearVacantes();
SLDocument VacantesError_SRosa = Escuela.CrearVacantes();
SLDocument VacantesError_Anguil = Escuela.CrearVacantes();

foreach (var alumno in Alumnos)
{
    switch (alumno.preferencia)
    {
        case "Nro. 1 - Toay":
            var resToay=Escuela_Toay.datosActualizado(alumno,VacantesToay,Escuela_Toay, VacantesError_Toay);
            if(resToay.Item2){
                VacantesToay=resToay.Item1;
            }else{
                resToay=Escuela_SRosa.datosActualizado(alumno,VacantesSRosa,Escuela_SRosa, VacantesError_Toay);
                if(resToay.Item2){
                    VacantesSRosa=resToay.Item1;
                }else{
                    resToay=Escuela_Anguil.datosActualizado(alumno,VacantesAnguil,Escuela_Anguil, VacantesError_Toay);
                    if(resToay.Item2){
                        VacantesAnguil=resToay.Item1;
                    }else{
                        VacantesError_Toay=resToay.Item1;
                    }
                }
            }
            break; 
        case "Nro. 2 - Santa Rosa":
            var resRsa=Escuela_SRosa.datosActualizado(alumno,VacantesSRosa,Escuela_SRosa, VacantesError_SRosa);
            if(resRsa.Item2){
                VacantesSRosa=resRsa.Item1;
            }else{
                resRsa=Escuela_Toay.datosActualizado(alumno,VacantesToay,Escuela_Toay, VacantesError_SRosa);
                if(resRsa.Item2){
                    VacantesToay=resRsa.Item1;
                }else{
                    resRsa=Escuela_Anguil.datosActualizado(alumno,VacantesAnguil,Escuela_Anguil, VacantesError_SRosa);
                    if(resRsa.Item2){
                        VacantesAnguil=resRsa.Item1;
                    }else{
                        VacantesError_SRosa=resRsa.Item1;
                    }
                }
            }
            break;
        case "Nro. 3 - Anguil":
            var resAnguil=Escuela_Anguil.datosActualizado(alumno,VacantesAnguil,Escuela_Anguil, VacantesError_Anguil);
            if(resAnguil.Item2){
                VacantesAnguil=resAnguil.Item1;
            }else{
                resAnguil=Escuela_SRosa.datosActualizado(alumno,VacantesSRosa,Escuela_SRosa, VacantesError_Anguil);
                if(resAnguil.Item2){
                    VacantesSRosa=resAnguil.Item1;
                }else{
                    resAnguil=Escuela_Toay.datosActualizado(alumno,VacantesToay,Escuela_Toay, VacantesError_Anguil);
                    if(resAnguil.Item2){
                        VacantesToay=resAnguil.Item1;
                    }else{
                        VacantesError_Anguil=resAnguil.Item1;
                    }
                }
            }
            break;
    }
}

VacantesToay.SaveAs(@"C:\Users\Usuario\Desktop\Evaluativo_Practica\Vacantes\Vacantes_ColegioToay.xlsx");
VacantesSRosa.SaveAs(@"C:\Users\Usuario\Desktop\Evaluativo_Practica\Vacantes\Vacantes_ColegioSRosa.xlsx");
VacantesAnguil.SaveAs(@"C:\Users\Usuario\Desktop\Evaluativo_Practica\Vacantes\Vacantes_ColegioAnguil.xlsx");

VacantesError_Toay.SaveAs(@"C:\Users\Usuario\Desktop\Evaluativo_Practica\Vacantes\Error\VacantesError_ColegioToay.xlsx");
VacantesError_SRosa.SaveAs(@"C:\Users\Usuario\Desktop\Evaluativo_Practica\Vacantes\Error\VacantesError_ColegioSRosa.xlsx");
VacantesError_Anguil.SaveAs(@"C:\Users\Usuario\Desktop\Evaluativo_Practica\Vacantes\Error\VacantesError_ColegioAnguil.xlsx");

