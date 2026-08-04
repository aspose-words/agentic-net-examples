using System;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject project = new VbaProject();
        project.Name = "ExportMacrosProject";
        doc.VbaProject = project;

        // Add a sample procedural module.
        VbaModule module1 = new VbaModule();
        module1.Name = "ModuleOne";
        module1.Type = VbaModuleType.ProceduralModule;
        module1.SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello, World!""
End Sub
";
        doc.VbaProject.Modules.Add(module1);

        // Add another sample module.
        VbaModule module2 = new VbaModule();
        module2.Name = "ModuleTwo";
        module2.Type = VbaModuleType.ProceduralModule;
        module2.SourceCode = @"
Function AddNumbers(a As Integer, b As Integer) As Integer
    AddNumbers = a + b
End Function
";
        doc.VbaProject.Modules.Add(module2);

        // Save the document as a macro-enabled file.
        string docPath = "ExportMacros.docm";
        doc.Save(docPath);

        // Extract macro information.
        var macros = new List<object>();
        foreach (VbaModule mod in doc.VbaProject.Modules)
        {
            string source = mod.SourceCode ?? string.Empty;
            macros.Add(new { Name = mod.Name, SourceCode = source });
        }

        // Serialize to JSON.
        var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
        string json = JsonSerializer.Serialize(macros, jsonOptions);

        // Output JSON.
        Console.WriteLine(json);
    }
}
