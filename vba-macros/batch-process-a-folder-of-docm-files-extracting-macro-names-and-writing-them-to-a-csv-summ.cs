using System;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define the folder that will contain the sample DOCM files and the output CSV.
        string folderPath = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        Directory.CreateDirectory(folderPath);

        // Create sample macro‑enabled documents.
        CreateSampleDocument(Path.Combine(folderPath, "Sample1.docm"),
            new[]
            {
                new VbaModuleInfo
                {
                    Name = "ModuleA",
                    SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello, World!""
End Sub

Function AddNumbers(a As Integer, b As Integer) As Integer
    AddNumbers = a + b
End Function"
                }
            });

        CreateSampleDocument(Path.Combine(folderPath, "Sample2.docm"),
            new[]
            {
                new VbaModuleInfo
                {
                    Name = "ModuleB",
                    SourceCode = @"
Sub ShowDate()
    MsgBox Date
End Sub

Sub EmptySub()
End Sub"
                },
                new VbaModuleInfo
                {
                    Name = "ModuleC",
                    SourceCode = null // Intentional empty source to test null handling.
                }
            });

        // Prepare the CSV file.
        string csvPath = Path.Combine(folderPath, "MacroSummary.csv");
        using (var writer = new StreamWriter(csvPath))
        {
            writer.WriteLine("Document,MacroName");

            // Process each DOCM file in the folder.
            foreach (string filePath in Directory.GetFiles(folderPath, "*.docm"))
            {
                // Load the document.
                Document doc = new Document(filePath);

                // Ensure the document actually contains macros.
                if (doc.HasMacros && doc.VbaProject != null)
                {
                    // Iterate through all VBA modules.
                    foreach (VbaModule module in doc.VbaProject.Modules)
                    {
                        // Guard against null source code.
                        string source = module.SourceCode ?? string.Empty;

                        // Use a simple regex to find Sub and Function declarations.
                        // Pattern matches lines starting with optional whitespace, then Sub or Function,
                        // then the macro name (word characters), ignoring case.
                        foreach (Match match in Regex.Matches(source,
                            @"^\s*(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)", RegexOptions.Multiline | RegexOptions.IgnoreCase))
                        {
                            string macroName = match.Groups[2].Value;
                            // Write the document name and macro name to the CSV.
                            writer.WriteLine($"{Path.GetFileName(filePath)},{macroName}");
                        }
                    }
                }
            }
        }

        // The program finishes automatically; no user interaction required.
    }

    // Helper method to create a macro‑enabled document with specified modules.
    private static void CreateSampleDocument(string filePath, VbaModuleInfo[] modulesInfo)
    {
        // Create a blank document.
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject project = new VbaProject();
        doc.VbaProject = project;

        // Add each module defined in modulesInfo.
        foreach (var info in modulesInfo)
        {
            VbaModule module = new VbaModule();
            module.Name = info.Name;
            module.Type = VbaModuleType.ProceduralModule;
            // Guard against null source code.
            module.SourceCode = info.SourceCode ?? string.Empty;
            project.Modules.Add(module);
        }

        // Save the document in macro‑enabled format.
        doc.Save(filePath);
    }

    // Simple DTO to hold module creation data.
    private class VbaModuleInfo
    {
        public string Name { get; set; }
        public string SourceCode { get; set; }
    }
}
