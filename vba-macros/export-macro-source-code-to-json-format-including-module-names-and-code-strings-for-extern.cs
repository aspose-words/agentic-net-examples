using System;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();

        // Ensure the document has a VBA project.
        if (doc.VbaProject == null)
        {
            doc.VbaProject = new VbaProject();
            doc.VbaProject.Name = "SampleProject";
        }

        // Add a procedural module with sample code.
        VbaModule module1 = new VbaModule();
        module1.Name = "Module1";
        module1.Type = VbaModuleType.ProceduralModule;
        module1.SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello from VBA!\"\nEnd Sub";
        doc.VbaProject.Modules.Add(module1);

        // Add a class module with sample code.
        VbaModule module2 = new VbaModule();
        module2.Name = "Class1";
        module2.Type = VbaModuleType.ClassModule;
        module2.SourceCode = "Public Sub Greet()\n    MsgBox \"Greetings from class!\"\nEnd Sub";
        doc.VbaProject.Modules.Add(module2);

        // Save the document as a macro-enabled file.
        string docPath = "MacroDocument.docm";
        doc.Save(docPath);

        // Export macro source code to JSON.
        var macros = new List<MacroInfo>();
        foreach (VbaModule mod in doc.VbaProject.Modules)
        {
            string code = mod.SourceCode ?? string.Empty; // Guard against null.
            macros.Add(new MacroInfo { Name = mod.Name, SourceCode = code });
        }

        string json = JsonSerializer.Serialize(macros, new JsonSerializerOptions { WriteIndented = true });
        Console.WriteLine(json);
    }

    private class MacroInfo
    {
        public string Name { get; set; }
        public string SourceCode { get; set; }
    }
}
