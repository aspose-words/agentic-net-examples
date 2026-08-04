using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define the folder that will contain the DOCM files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        Directory.CreateDirectory(inputFolder);

        // If the folder is empty, create sample macro‑enabled documents.
        if (Directory.GetFiles(inputFolder, "*.docm").Length == 0)
        {
            CreateSampleDocument(Path.Combine(inputFolder, "Sample1.docm"));
            CreateSampleDocument(Path.Combine(inputFolder, "Sample2.docm"));
        }

        // Prepare a list to hold CSV rows.
        var csvRows = new List<string> { "Document,MacroName" };

        // Process each DOCM file in the folder.
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.docm"))
        {
            Document doc = new Document(filePath);

            // Ensure the document actually contains a VBA project.
            if (doc.HasMacros && doc.VbaProject != null)
            {
                foreach (VbaModule module in doc.VbaProject.Modules)
                {
                    // Guard against null source code.
                    string source = module.SourceCode ?? string.Empty;

                    // Extract macro names from the source code.
                    foreach (string macroName in ExtractMacroNames(source))
                    {
                        // Add a CSV row: document file name and macro name.
                        csvRows.Add($"{Path.GetFileName(filePath)},{macroName}");
                    }
                }
            }
        }

        // Write the CSV summary file.
        string csvPath = Path.Combine(inputFolder, "MacroSummary.csv");
        File.WriteAllLines(csvPath, csvRows);
    }

    // Creates a simple macro‑enabled document with one procedural module containing two macros.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject project = new VbaProject();
        project.Name = "SampleProject";
        doc.VbaProject = project;

        // Create a module with sample VBA code.
        VbaModule module = new VbaModule();
        module.Name = "SampleModule";
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = @"
Sub MacroOne()
    MsgBox ""Hello from MacroOne""
End Sub

Sub MacroTwo()
    MsgBox ""Hello from MacroTwo""
End Sub
";

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(module);

        // Save as a macro‑enabled document.
        doc.Save(filePath);
    }

    // Parses VBA source code and returns the names of Sub and Function macros.
    private static IEnumerable<string> ExtractMacroNames(string source)
    {
        var macroNames = new List<string>();
        using (StringReader reader = new StringReader(source))
        {
            string line;
            while ((line = reader.ReadLine()) != null)
            {
                line = line.TrimStart();

                // Look for Sub or Function declarations.
                if (line.StartsWith("Sub ", StringComparison.OrdinalIgnoreCase) ||
                    line.StartsWith("Function ", StringComparison.OrdinalIgnoreCase))
                {
                    // Remove the keyword.
                    int startIdx = line.IndexOf(' ') + 1;
                    if (startIdx > 0 && startIdx < line.Length)
                    {
                        // Extract the name up to the first '(' or whitespace.
                        int endIdx = line.IndexOf('(', startIdx);
                        if (endIdx == -1)
                            endIdx = line.IndexOf(' ', startIdx);
                        if (endIdx == -1)
                            endIdx = line.Length;

                        string name = line.Substring(startIdx, endIdx - startIdx).Trim();
                        if (!string.IsNullOrEmpty(name))
                            macroNames.Add(name);
                    }
                }
            }
        }
        return macroNames;
    }
}
