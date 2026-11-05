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

        // Ensure the document has a VBA project.
        if (doc.VbaProject == null)
        {
            doc.VbaProject = new VbaProject();
        }

        // Add a VBA module with sample code.
        VbaModule module1 = new VbaModule();
        module1.Name = "Module1";
        module1.SourceCode = @"Sub Hello()
    MsgBox ""Hello from VBA!""
End Sub";
        doc.VbaProject.Modules.Add(module1);

        // Add another VBA module.
        VbaModule module2 = new VbaModule();
        module2.Name = "Module2";
        module2.SourceCode = @"Function AddNumbers(a As Integer, b As Integer) As Integer
    AddNumbers = a + b
End Function";
        doc.VbaProject.Modules.Add(module2);

        // Save the document as a macro-enabled file.
        string filePath = Path.Combine(Path.GetTempPath(), "SampleMacroDocument.docm");
        doc.Save(filePath, SaveFormat.Docm);

        // Reload the document to demonstrate reading.
        Document loadedDoc = new Document(filePath);

        // Prepare a list to hold module information.
        var modulesInfo = new List<Dictionary<string, string>>();

        // Export each module's name and source code to the list.
        if (loadedDoc.VbaProject != null)
        {
            foreach (VbaModule mod in loadedDoc.VbaProject.Modules)
            {
                string source = mod.SourceCode ?? string.Empty;
                var info = new Dictionary<string, string>
                {
                    { "Name", mod.Name },
                    { "SourceCode", source }
                };
                modulesInfo.Add(info);
            }
        }

        // Serialize the list to JSON.
        string json = JsonSerializer.Serialize(modulesInfo, new JsonSerializerOptions { WriteIndented = true });

        // Output the JSON.
        Console.WriteLine(json);
    }
}
