using SpreadsheetLight;
using System.Drawing;
using CrearAlumno;


        //Leer incripciones
        string doc_incripcione = @"C:\Users\Usuario\Desktop\Evaluativo_Practica\DATA\Inscripciones.xlsx";
        SLDocument incripciones = new SLDocument(doc_incripcione);
        List<Alumno> Alumnos= new List<Alumno>();
        int fila = 1;
        while (!string.IsNullOrEmpty(incripciones.GetCellValueAsString(fila, 1)))
        {
            if (fila == 1)
            {
                string enc1 = incripciones.GetCellValueAsString(fila, 1);
                string enc2 = incripciones.GetCellValueAsString(fila, 2);
                string enc3 = incripciones.GetCellValueAsString(fila, 3);
                string enc4 = incripciones.GetCellValueAsString(fila, 4);
                string enc5 = incripciones.GetCellValueAsString(fila, 5);
            }
            // Y aca para los datos 
            else
            {
                int ni = incripciones.GetCellValueAsInt32(fila, 1);
                string nom = incripciones.GetCellValueAsString(fila, 2);
                int edad = incripciones.GetCellValueAsInt32(fila, 3);
                int grado = incripciones.GetCellValueAsInt32(fila, 4);
                string preferencia = incripciones.GetCellValueAsString(fila, 5);
                Alumnos.Add(new Alumno(ni, nom, edad, grado, preferencia));
            }
            fila++;
        }
        foreach (var alumno in Alumnos)
        {
            System.Console.WriteLine(alumno.id);
        }

        // //Leer el excel colegio toay
        // string doc_toay = @"C:\Users\Usuario\Desktop\Evaluativo_Practica\DATA\Colegio_Toay.xlsx";
        // SLDocument colegio_toay = new SLDocument(doc_toay);

        // fila = 1;
        // while (!string.IsNullOrEmpty(colegio_toay.GetCellValueAsString(fila, 1)))
        // {
        //     if (fila == 1)
        //     {
        //         string enc1 = colegio_toay.GetCellValueAsString(fila, 1);
        //         string enc2 = colegio_toay.GetCellValueAsString(fila, 2);
        //         //Console.WriteLine(enc1 + "    " + enc2);
        //     }
        //     else
        //     {
        //         int grado = colegio_toay.GetCellValueAsInt32(fila, 1);
        //         int vacantes = colegio_toay.GetCellValueAsInt32(fila, 2);
        //         //Console.WriteLine(grado + "        " + vacantes);
        //     }
        //     fila++;
        // }

        // //Leer el excel colegio santarosa
        // string doc_santarosa = @"C:\Users\Usuario\Desktop\Evaluativo_Practica\DATA\Colegio_SantaRosa.xlsx";
        // SLDocument colegio_santarosa = new SLDocument(doc_santarosa);
        // fila = 1;
        // while (!string.IsNullOrEmpty(colegio_toay.GetCellValueAsString(fila, 1)))
        // {
        //     if (fila == 1)
        //     {
        //         string enc1 = colegio_santarosa.GetCellValueAsString(fila, 1);
        //         string enc2 = colegio_santarosa.GetCellValueAsString(fila, 2);
        //         //Console.WriteLine(enc1 + "    " + enc2);
        //     }
        //     else
        //     {
        //         int grado = colegio_santarosa.GetCellValueAsInt32(fila, 1);
        //         int vacantes = colegio_santarosa.GetCellValueAsInt32(fila, 2);
        //         // Console.WriteLine(grado + "        " + vacantes);
        //     }
        //     fila++;
        // }

        // //Leer el xcel de colegio anguil
        // string doc_anguil = @"C:\Users\Usuario\Desktop\Evaluativo_Practica\DATA\Colegio_Anguil.xlsx";
        // SLDocument colegio_anguil = new SLDocument(doc_anguil);
        // fila = 1;
        // while (!string.IsNullOrEmpty(colegio_anguil.GetCellValueAsString(fila, 1)))
        // {
        //     if (fila == 1)
        //     {
        //         string enc1 = colegio_anguil.GetCellValueAsString(fila, 1);
        //         string enc2 = colegio_anguil.GetCellValueAsString(fila, 2);
        //         //Console.WriteLine(enc1 + "    " + enc2);
        //     }
        //     else
        //     {
        //         int grado = colegio_anguil.GetCellValueAsInt32(fila, 1);
        //         int vacantes = colegio_anguil.GetCellValueAsInt32(fila, 2);
        //         //Console.WriteLine(grado + "        " + vacantes);
        //     }
        //     fila++;
        // }


        // //columna de vacantes toay
        // SLDocument vac_toay = new SLDocument();
        // vac_toay.SetCellValue(1, 1, "nro_inscripcion");
        // vac_toay.SetCellValue(1, 2, "Nombre");
        // vac_toay.SetCellValue(1, 3, "Edad");
        // vac_toay.SetCellValue(1, 4, "Grado");
        // //agregar las filas

        // //crear excel vacante toay
        // vac_toay.SaveAs(@"C:\Users\Usuario\Desktop\Evaluativo_Practica\Vacantes_ColegioToay.xlsx");


        // //columna de vacantes Santa rosa
        // SLDocument vac_santarosa = new SLDocument();
        // vac_santarosa.SetCellValue(1, 1, "nro_inscripcion");
        // vac_santarosa.SetCellValue(1, 2, "Nombre");
        // vac_santarosa.SetCellValue(1, 3, "Edad");
        // vac_santarosa.SetCellValue(1, 4, "Grado");
        // //agregar las filas

        // //crear excel vacante santa rosa 
        // vac_santarosa.SaveAs(@"C:\Users\Usuario\Desktop\Evaluativo_Practica\Vacantes_ColegioSantaRosa.xlsx");

        // //columna de vacantes anguil
        // SLDocument vac_anguil = new SLDocument();
        // vac_anguil.SetCellValue(1, 1, "nro_inscripcion");
        // vac_anguil.SetCellValue(1, 2, "Nombre");
        // vac_anguil.SetCellValue(1, 3, "Edad");
        // vac_anguil.SetCellValue(1, 4, "Grado");
        // //agregar las filas

        // //crear excel vacante Santa rosa
        // vac_anguil.SaveAs(@"C:\Users\Usuario\Desktop\Evaluativo_Practica\Vacantes_ColegioAnguil.xlsx");


        // Console.ReadLine();
