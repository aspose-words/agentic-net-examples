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
        // Folder to hold sample DOCM files and the output CSV.
        string folderPath = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        Directory.CreateDirectory(folderPath);

        // Create sample macro-enabled documents.
        CreateSampleDocM(Path.Combine(folderPath, "Sample1.docm"), new[]
        {
            new Tuple<string, string>("ModuleA", @"
Sub HelloWorld()
    MsgBox ""Hello World!""
End Sub

Function AddNumbers(a As Integer, b As Integer) As Integer
    AddNumbers = a + b
End Function")
        });

        CreateSampleDocM(Path.Combine(folderPath, "Sample2.docm"), new[]
        {
            new Tuple<string, string>("ModuleB", @"
Sub ShowDate()
    MsgBox Date
End Sub")
        });

        // Prepare CSV output.
        string csvPath = Path.Combine(folderPath, "MacroSummary.csv");
        var csvLines = new List<string>();
        csvLines.Add("FileName,ModuleName,MacroName");

        // Process each DOCM file in the folder.
        foreach (string file in Directory.GetFiles(folderPath, "*.docm"))
        {
            Document doc = new Document(file);

            if (!doc.HasMacros || doc.VbaProject == null)
                continue; // No macros to extract.

            foreach (VbaModule module in doc.VbaProject.Modules)
            {
                string source = module.SourceCode ?? string.Empty;

                // Find macro (Sub/Function) names using a simple regex.
                foreach (Match match in Regex.Matches(source, @"\b(Sub|Function)\s+(\w+)", RegexOptions.IgnoreCase))
                {
                    string macroName = match.Groups[2].Value;
                    // Escape commas in fields if needed.
                    string line = $"{Path.GetFileName(file)},{EscapeCsv(module.Name)},{EscapeCsv(macroName)}";
                    csvLines.Add(line);
                }
            }
        }

        // Write the CSV file.
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);
    }

    // Helper to create a macro-enabled document with specified modules.
    private static void CreateSampleDocM(string filePath, Tuple<string, string>[] modulesInfo)
    {
        Document doc = new Document();

        // Ensure a VBA project exists.
        if (doc.VbaProject == null)
        {
            VbaProject project = new VbaProject { Name = "SampleProject" };
            doc.VbaProject = project;
        }

        foreach (var info in modulesInfo)
        {
            VbaModule module = new VbaModule
            {
                Name = info.Item1,
                Type = VbaModuleType.ProceduralModule,
                SourceCode = info.Item2
            };
            doc.VbaProject.Modules.Add(module);
        }

        // Save as a macro-enabled document.
        doc.Save(filePath);
    }

    // Simple CSV field escaper.
    private static string EscapeCsv(string field)
    {
        if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
        {
            string escaped = field.Replace("\"", "\"\"");
            return $"\"{escaped}\"";
        }
        return field;
    }
}
