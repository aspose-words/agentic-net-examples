using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class RetrieveVbaModuleSource
{
    public static void Main()
    {
        // Define output directories and file names.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        string docPath = Path.Combine(artifactsDir, "SampleMacro.docm");
        string outputPath = Path.Combine(artifactsDir, "ModuleSource.txt");
        string targetModuleName = "SampleModule";

        // -----------------------------------------------------------------
        // 1. Create a new macro‑enabled document and add a VBA module.
        // -----------------------------------------------------------------
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject vbaProject = new VbaProject();
        vbaProject.Name = "SampleProject";
        doc.VbaProject = vbaProject;

        // Create a new VBA module with some sample code.
        VbaModule vbaModule = new VbaModule();
        vbaModule.Name = targetModuleName;
        vbaModule.Type = VbaModuleType.ProceduralModule;
        vbaModule.SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from VBA!""
End Sub
";

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(vbaModule);

        // Save the document in a macro‑enabled format.
        doc.Save(docPath, SaveFormat.Docm);

        // -----------------------------------------------------------------
        // 2. Load the document and retrieve the source code of the target module.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // Ensure the document actually contains macros.
        if (!loadedDoc.HasMacros || loadedDoc.VbaProject == null)
        {
            Console.WriteLine("The document does not contain a VBA project.");
            return;
        }

        // Access the module by name; if not found, fall back to the first module.
        VbaModule targetModule = loadedDoc.VbaProject.Modules[targetModuleName];
        if (targetModule == null && loadedDoc.VbaProject.Modules.Count > 0)
            targetModule = loadedDoc.VbaProject.Modules[0];

        // Guard against null source code.
        string sourceCode = targetModule?.SourceCode ?? string.Empty;

        // Write the source code to a text file.
        File.WriteAllText(outputPath, sourceCode);

        // Optional: indicate completion (no interactive input required).
        Console.WriteLine($"VBA module source written to: {outputPath}");
    }
}
