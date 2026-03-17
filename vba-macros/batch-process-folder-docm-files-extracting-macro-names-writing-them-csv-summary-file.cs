using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Vba;

class MacroExtractor
{
    static void Main()
    {
        // Folder containing the DOCM files.
        string inputFolder = @"C:\Docs\MacroEnabled";
        // Path for the CSV summary file.
        string csvPath = @"C:\Docs\MacroSummary.csv";

        // Prepare the CSV file with a header.
        using (var writer = new StreamWriter(csvPath, false))
        {
            writer.WriteLine("Document,ModuleName,MacroName");

            // Process each .docm file in the folder.
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.docm"))
            {
                // Load the document (lifecycle rule: use Document constructor).
                Document doc = new Document(filePath);

                // If the document does not contain macros, skip it.
                if (!doc.HasMacros || doc.VbaProject == null)
                    continue;

                // Iterate through all VBA modules.
                foreach (VbaModule module in doc.VbaProject.Modules)
                {
                    // The module name is a reasonable identifier for the macro container.
                    // For a more detailed list, you could parse module.SourceCode for Sub/Function names.
                    string moduleName = module.Name ?? string.Empty;

                    // Write a CSV line: Document file name, module name, (placeholder for macro name).
                    // Since extracting individual macro names requires parsing the source code,
                    // we record the module name as the macro identifier.
                    writer.WriteLine($"{Path.GetFileName(filePath)},{EscapeCsv(moduleName)},{EscapeCsv(moduleName)}");
                }
            }
        }

        // Optionally, open the CSV after creation (not required for batch processing).
        Console.WriteLine($"Macro summary written to: {csvPath}");
    }

    // Helper to escape CSV fields that may contain commas or quotes.
    private static string EscapeCsv(string field)
    {
        if (field.Contains(',') || field.Contains('\"') || field.Contains('\n'))
        {
            field = field.Replace("\"", "\"\"");
            return $"\"{field}\"";
        }
        return field;
    }
}
