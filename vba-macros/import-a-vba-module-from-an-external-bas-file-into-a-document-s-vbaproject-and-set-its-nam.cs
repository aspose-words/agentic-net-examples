using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a temporary .bas file with simple VBA code
        string basFilePath = Path.Combine(Path.GetTempPath(), "SampleModule.bas");
        string vbaCode = @"
Sub HelloWorld()
    MsgBox ""Hello from imported module!""
End Sub
";
        File.WriteAllText(basFilePath, vbaCode);

        // Create a new blank document
        Document doc = new Document();

        // Ensure the document has a VBA project
        if (doc.VbaProject == null)
        {
            doc.VbaProject = new VbaProject();
        }

        // Read the VBA source code from the .bas file
        string sourceCode = File.ReadAllText(basFilePath) ?? string.Empty;

        // Create a new VBA module, set its name and source code
        VbaModule importedModule = new VbaModule();
        importedModule.Name = "ImportedModule";
        importedModule.SourceCode = sourceCode;

        // Add the module to the document's VBA project
        doc.VbaProject.Modules.Add(importedModule);

        // Save the document in a macro‑enabled format
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DocumentWithImportedMacro.docm");
        doc.Save(outputPath, SaveFormat.Docm);

        // Simple validation: check that the module exists and its name is set correctly
        bool moduleExists = false;
        foreach (VbaModule module in doc.VbaProject.Modules)
        {
            if (module.Name == "ImportedModule")
            {
                moduleExists = true;
                break;
            }
        }

        Console.WriteLine(moduleExists
            ? "VBA module imported successfully."
            : "Failed to import VBA module.");
    }
}
