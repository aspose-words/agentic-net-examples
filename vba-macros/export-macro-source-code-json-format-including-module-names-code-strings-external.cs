using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Vba;

class ExportVbaMacrosToJson
{
    static void Main()
    {
        // Path to the input macro‑enabled document.
        string inputPath = @"C:\Docs\MacroDocument.docm";

        // Path where the JSON file will be written.
        string outputPath = @"C:\Docs\VbaModules.json";

        // Load the document (uses Aspose.Words loading rule).
        Document doc = new Document(inputPath);

        // Prepare a collection to hold module name and source code.
        var modulesInfo = new List<object>();

        // Ensure the document actually contains a VBA project.
        if (doc.VbaProject != null)
        {
            // Iterate through all VBA modules.
            foreach (VbaModule module in doc.VbaProject.Modules)
            {
                // Add an anonymous object with the required fields.
                modulesInfo.Add(new
                {
                    Name = module.Name,
                    SourceCode = module.SourceCode
                });
            }
        }

        // Serialize the collection to a formatted JSON string.
        string json = JsonSerializer.Serialize(modulesInfo, new JsonSerializerOptions
        {
            WriteIndented = true
        });

        // Write the JSON to the specified file (uses standard .NET I/O).
        File.WriteAllText(outputPath, json);

        Console.WriteLine($"Exported {modulesInfo.Count} VBA module(s) to \"{outputPath}\".");
    }
}
