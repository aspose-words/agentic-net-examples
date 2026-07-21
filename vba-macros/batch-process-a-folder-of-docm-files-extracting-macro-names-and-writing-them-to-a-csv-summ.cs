using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Folder that will contain the macro-enabled documents.
        string docsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        Directory.CreateDirectory(docsFolder);

        // Create sample DOCM files with VBA macros.
        for (int i = 1; i <= 2; i++)
        {
            Document doc = new Document();

            // Ensure the document has a VBA project.
            VbaProject project = new VbaProject();
            project.Name = $"Project{i}";
            doc.VbaProject = project;

            // Create a procedural module with a couple of macros.
            VbaModule module = new VbaModule();
            module.Name = $"Module{i}";
            module.Type = VbaModuleType.ProceduralModule;
            module.SourceCode = $@"
Sub Macro{i}_One()
    MsgBox ""Hello from macro {i}_One""
End Sub

Sub Macro{i}_Two()
    MsgBox ""Hello from macro {i}_Two""
End Sub
";
            doc.VbaProject.Modules.Add(module);

            // Save as a macro‑enabled document.
            string docPath = Path.Combine(docsFolder, $"Sample{i}.docm");
            doc.Save(docPath);
        }

        // Prepare CSV output.
        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "MacroSummary.csv");
        using (var writer = new StreamWriter(csvPath, false, Encoding.UTF8))
        {
            writer.WriteLine("FileName,MacroName");

            // Process each DOCM file in the folder.
            foreach (string filePath in Directory.GetFiles(docsFolder, "*.docm"))
            {
                Document doc = new Document(filePath);

                // Skip files without macros.
                if (!doc.HasMacros || doc.VbaProject == null)
                    continue;

                foreach (VbaModule module in doc.VbaProject.Modules)
                {
                    // Guard against null source code.
                    string source = module.SourceCode ?? string.Empty;

                    // Simple regex to capture Sub or Function names.
                    foreach (Match match in Regex.Matches(source, @"\b(Sub|Function)\s+(\w+)", RegexOptions.IgnoreCase))
                    {
                        string macroName = match.Groups[2].Value;
                        string fileName = Path.GetFileName(filePath);
                        writer.WriteLine($"{EscapeCsv(fileName)},{EscapeCsv(macroName)}");
                    }
                }
            }
        }

        // Optional: indicate completion.
        Console.WriteLine($"Macro summary written to: {csvPath}");
    }

    // Helper to escape CSV fields that may contain commas or quotes.
    private static string EscapeCsv(string field)
    {
        if (field.Contains("\""))
            field = field.Replace("\"", "\"\"");
        if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
            field = $"\"{field}\"";
        return field;
    }
}
