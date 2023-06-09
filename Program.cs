using SpreadsheetLight;
using CrearAlumno;
using CrearEscuela;
using Funcs;
using Spectre.Console;



Functions funciones = new Functions();

bool bandera = true;
while (bandera)
{
    AnsiConsole.Background = ConsoleColor.Yellow;
    AnsiConsole.Foreground = ConsoleColor.Red;
    AnsiConsole.WriteLine("PROGRAMA DE ADMINISTRADOR DE EMPLEADOS");
    AnsiConsole.WriteLine("=======================================");
    AnsiConsole.WriteLine();
    AnsiConsole.Reset();
    var menu = AnsiConsole.Prompt(new SelectionPrompt<String>().Title("[green]ELIJAN UNA OPCION[/]")
            .AddChoices(new string[] { "HACER GESTION","VER VACANTES","GUARDAR VACANTES", "SALIR" }));
    switch (menu)
    {
        case "HACER GESTION":
            funciones.GestionarInscripciones();
            break;
        case "VER VACANTES":
            funciones.VerVacantes();
            break;
        case "GUARDAR VACANTES":
            funciones.GuardarDocumentos();
            break;
        case "SALIR":
            bandera = false;
            break;
    }
}