using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

class MacroExtractor
{
    static void Main(string[] args)
    {
        // Determine input folder and CSV output path.
        string inputFolder = args.Length > 0 ? args[0] : Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        string csvPath = args.Length > 1 ? args[1] : Path.Combine(inputFolder, "MacroSummary.csv");

        // Ensure the input folder exists.
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder \"{inputFolder}\" does not exist. Creating it.");
            Directory.CreateDirectory(inputFolder);
            Console.WriteLine("Place .docm files in this folder and rerun the program.");
            return;
        }

        // Prepare the CSV file with a header.
        using (var writer = new StreamWriter(csvPath, false))
        {
            writer.WriteLine("Document,ModuleName,MacroName");

            // Process each .docm file in the folder.
            foreach (string filePath in Directory.EnumerateFiles(inputFolder, "*.docm"))
            {
                try
                {
                    // Load the document.
                    Document doc = new Document(filePath);

                    // Skip if the document does not contain macros.
                    if (!doc.HasMacros || doc.VbaProject == null)
                        continue;

                    // Iterate through all VBA modules.
                    foreach (VbaModule module in doc.VbaProject.Modules)
                    {
                        string moduleName = module.Name ?? string.Empty;
                        // Record the module name as the macro identifier (parsing individual macros is out of scope).
                        writer.WriteLine($"{EscapeCsv(Path.GetFileName(filePath))},{EscapeCsv(moduleName)},{EscapeCsv(moduleName)}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing \"{filePath}\": {ex.Message}");
                }
            }
        }

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
