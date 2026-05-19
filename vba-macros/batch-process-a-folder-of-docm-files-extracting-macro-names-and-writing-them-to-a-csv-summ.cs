using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define the folder that will contain the sample DOCM files.
        string docsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        Directory.CreateDirectory(docsFolder);

        // Create sample macro-enabled documents.
        CreateSampleDocM(Path.Combine(docsFolder, "Sample1.docm"), "ModuleA", "Sub MacroA()\n    MsgBox \"Hello A\"\nEnd Sub");
        CreateSampleDocM(Path.Combine(docsFolder, "Sample2.docm"), "ModuleB", "Sub MacroB()\n    MsgBox \"Hello B\"\nEnd Sub");
        CreateSampleDocM(Path.Combine(docsFolder, "Sample3.docm"), "ModuleC", "Sub MacroC()\n    MsgBox \"Hello C\"\nEnd Sub");

        // Prepare the CSV output.
        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "MacroSummary.csv");
        var csvLines = new List<string>();
        csvLines.Add("FileName,ModuleName");

        // Process each DOCM file in the folder.
        foreach (string filePath in Directory.GetFiles(docsFolder, "*.docm"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Check if the document contains macros.
            if (doc.HasMacros && doc.VbaProject != null)
            {
                // Iterate through all VBA modules.
                foreach (VbaModule module in doc.VbaProject.Modules)
                {
                    // Guard against null module name.
                    string moduleName = module?.Name ?? string.Empty;
                    // Add a line to the CSV.
                    csvLines.Add($"{Path.GetFileName(filePath)},{moduleName}");
                }
            }
            else
            {
                // Document has no macros; still record the file with empty module name.
                csvLines.Add($"{Path.GetFileName(filePath)},");
            }
        }

        // Write all lines to the CSV file.
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);
    }

    // Helper method to create a macro-enabled document with a single module.
    private static void CreateSampleDocM(string filePath, string moduleName, string sourceCode)
    {
        // Create a blank document.
        Document doc = new Document();

        // Ensure a VBA project exists.
        VbaProject project = new VbaProject();
        project.Name = "SampleProject";
        doc.VbaProject = project;

        // Create a new module and set its properties.
        VbaModule module = new VbaModule();
        module.Name = moduleName;
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = sourceCode ?? string.Empty;

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(module);

        // Save the document in macro-enabled format.
        doc.Save(filePath);
    }
}
