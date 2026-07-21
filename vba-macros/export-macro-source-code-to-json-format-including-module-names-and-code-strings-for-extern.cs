using System;
using System.Collections.Generic;
using System.IO;
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
        project.Name = "SampleProject";
        doc.VbaProject = project;

        // Create the first VBA module.
        VbaModule module1 = new VbaModule();
        module1.Name = "Module1";
        module1.Type = VbaModuleType.ProceduralModule;
        module1.SourceCode = "Sub Hello()\n    MsgBox \"Hello\"\nEnd Sub";

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(module1);

        // Create a second VBA module.
        VbaModule module2 = new VbaModule();
        module2.Name = "Module2";
        module2.Type = VbaModuleType.ProceduralModule;
        module2.SourceCode = "Function Add(a As Integer, b As Integer) As Integer\n    Add = a + b\nEnd Function";

        // Add the second module.
        doc.VbaProject.Modules.Add(module2);

        // Save the document as a macro‑enabled file.
        string filePath = "MacroDocument.docm";
        doc.Save(filePath);

        // Export macro source code to JSON.
        var macroList = new List<object>();

        if (doc.VbaProject != null)
        {
            foreach (VbaModule module in doc.VbaProject.Modules)
            {
                string source = module.SourceCode ?? string.Empty;
                macroList.Add(new { Name = module.Name, SourceCode = source });
            }
        }

        string json = JsonSerializer.Serialize(macroList, new JsonSerializerOptions { WriteIndented = true });
        Console.WriteLine(json);
    }
}
