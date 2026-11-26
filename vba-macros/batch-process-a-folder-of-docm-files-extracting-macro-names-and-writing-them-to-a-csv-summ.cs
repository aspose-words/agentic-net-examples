using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        Directory.CreateDirectory(baseDir);

        // Create sample DOCM files
        for (int i = 1; i <= 2; i++)
        {
            string filePath = Path.Combine(baseDir, $"Sample{i}.docm");
            CreateSampleDocm(filePath, i);
        }

        // Process files and collect macro names
        var csvBuilder = new StringBuilder();
        csvBuilder.AppendLine("FileName,MacroName");

        foreach (string docPath in Directory.GetFiles(baseDir, "*.docm"))
        {
            var doc = new Document(docPath);
            var vbaProject = doc.VbaProject;
            if (vbaProject == null) continue;

            foreach (VbaModule module in vbaProject.Modules)
            {
                string source = module.SourceCode ?? string.Empty;
                foreach (string macro in ExtractMacroNames(source))
                {
                    csvBuilder.AppendLine($"{Path.GetFileName(docPath)},{macro}");
                }
            }
        }

        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "macro_summary.csv");
        File.WriteAllText(csvPath, csvBuilder.ToString());
    }

    private static void CreateSampleDocm(string path, int index)
    {
        var doc = new Document();

        // Ensure the document has a VBA project
        var vbaProject = new VbaProject();

        string moduleSource = $@"
Sub Macro{index}_One()
    MsgBox ""Hello from Macro {index}_One""
End Sub

Function Macro{index}_Two() As String
    Macro{index}_Two = ""Result""
End Function
";

        // Create a module using the parameterless constructor and set properties
        var module = new VbaModule
        {
            Name = $"Module{index}",
            SourceCode = moduleSource
        };

        vbaProject.Modules.Add(module);
        doc.VbaProject = vbaProject;

        // Save as a macro‑enabled document
        doc.Save(path, SaveFormat.Docm);
    }

    private static IEnumerable<string> ExtractMacroNames(string source)
    {
        var names = new List<string>();
        if (string.IsNullOrEmpty(source)) return names;

        var regex = new Regex(@"^\s*(Sub|Function)\s+([A-Za-z0-9_]+)", RegexOptions.Multiline | RegexOptions.IgnoreCase);
        foreach (Match match in regex.Matches(source))
        {
            if (match.Groups.Count > 2)
            {
                names.Add(match.Groups[2].Value);
            }
        }
        return names;
    }
}
