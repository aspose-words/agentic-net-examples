using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Vba;

class ExportVbaMacrosToJson
{
    static void Main(string[] args)
    {
        // Determine input and output paths.
        string inputPath = args.Length > 0 ? args[0] : Path.Combine(Directory.GetCurrentDirectory(), "MacroDocument.docm");
        string outputPath = args.Length > 1 ? args[1] : Path.Combine(Directory.GetCurrentDirectory(), "VbaModules.json");

        // Verify that the input file exists.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: \"{inputPath}\"");
            Console.WriteLine("Place a macro‑enabled document (.docm) at the specified location or pass the path as an argument.");
            return;
        }

        // Load the document.
        Document doc = new Document(inputPath);

        // Collect module information.
        var modulesInfo = new List<object>();

        if (doc.VbaProject != null)
        {
            foreach (VbaModule module in doc.VbaProject.Modules)
            {
                modulesInfo.Add(new
                {
                    Name = module.Name,
                    SourceCode = module.SourceCode
                });
            }
        }

        // Serialize to JSON.
        string json = JsonSerializer.Serialize(modulesInfo, new JsonSerializerOptions { WriteIndented = true });

        // Write JSON to the output file.
        File.WriteAllText(outputPath, json);

        Console.WriteLine($"Exported {modulesInfo.Count} VBA module(s) to \"{outputPath}\".");
    }
}
